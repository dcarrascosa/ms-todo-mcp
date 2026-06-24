#!/usr/bin/env python3
"""MCP server exposing Microsoft To Do via the Microsoft Graph To Do API.

Tools (all delegated, acting as the signed-in user):
    todo_list_lists     - list your task lists
    todo_list_tasks     - list tasks in a list (by id or name)
    todo_create_task    - create a task
    todo_update_task    - update fields of a task
    todo_complete_task  - mark a task completed
    todo_delete_task    - delete a task
    todo_create_list    - create a new task list

Graph reference: https://learn.microsoft.com/graph/api/resources/todo-overview
"""

from __future__ import annotations

import json
import os
from enum import Enum
from typing import Any, Optional

from mcp.server.fastmcp import FastMCP
from pydantic import BaseModel, ConfigDict, Field, field_validator

from .auth import AuthError, token_provider_from_env
from .graph import GraphClient, GraphError

mcp = FastMCP("ms_todo_mcp")

# --------------------------------------------------------------------------
# Client singleton (built lazily so the module imports without credentials)
# --------------------------------------------------------------------------
_CLIENT: Optional[GraphClient] = None


def _client() -> GraphClient:
    global _CLIENT
    if _CLIENT is None:
        _CLIENT = GraphClient(token_provider_from_env())
    return _CLIENT


# --------------------------------------------------------------------------
# Enums
# --------------------------------------------------------------------------
class ResponseFormat(str, Enum):
    MARKDOWN = "markdown"
    JSON = "json"


class Importance(str, Enum):
    LOW = "low"
    NORMAL = "normal"
    HIGH = "high"


class TaskStatus(str, Enum):
    NOT_STARTED = "notStarted"
    IN_PROGRESS = "inProgress"
    COMPLETED = "completed"
    WAITING_ON_OTHERS = "waitingOnOthers"
    DEFERRED = "deferred"


# --------------------------------------------------------------------------
# Formatting / helper utilities
# --------------------------------------------------------------------------
def _format_error(exc: Exception) -> str:
    if isinstance(exc, AuthError):
        return f"Error: authentication failed. {exc}"
    if isinstance(exc, GraphError):
        status = exc.status
        if status == 401:
            return (
                "Error: not authenticated or token expired. "
                "Run 'ms-todo-mcp login' to sign in again."
            )
        if status == 403:
            return (
                "Error: permission denied. The app needs the delegated 'Tasks.ReadWrite' "
                "permission (and admin consent in your tenant)."
            )
        if status == 404:
            return "Error: not found. Check the list_id / task_id is correct."
        if status == 429:
            return "Error: rate limit exceeded. Wait a moment and retry."
        return f"Error: Microsoft Graph request failed (HTTP {status}): {exc}"
    return f"Error: unexpected {type(exc).__name__}: {exc}"


def _to_graph_datetime(value: str) -> str:
    """Normalise a date/datetime string into Graph's naive dateTime form."""
    value = value.strip()
    if value.endswith("Z"):
        value = value[:-1]
    if "T" not in value:  # date only -> midnight
        value = f"{value}T00:00:00"
    return value


def _default_timezone() -> str:
    return os.environ.get("MS_TODO_TIMEZONE", "UTC")


def _list_to_dict(lst: dict[str, Any]) -> dict[str, Any]:
    return {
        "id": lst.get("id"),
        "displayName": lst.get("displayName"),
        "isOwner": lst.get("isOwner"),
        "isShared": lst.get("isShared"),
        "wellknownListName": lst.get("wellknownListName"),
    }


def _task_to_dict(task: dict[str, Any]) -> dict[str, Any]:
    due = task.get("dueDateTime") or {}
    body = task.get("body") or {}
    return {
        "id": task.get("id"),
        "title": task.get("title"),
        "status": task.get("status"),
        "importance": task.get("importance"),
        "isReminderOn": task.get("isReminderOn"),
        "dueDateTime": due.get("dateTime"),
        "dueTimeZone": due.get("timeZone"),
        "body": body.get("content") or None,
        "createdDateTime": task.get("createdDateTime"),
        "lastModifiedDateTime": task.get("lastModifiedDateTime"),
        "completedDateTime": (task.get("completedDateTime") or {}).get("dateTime"),
    }


def _tasks_markdown(tasks: list[dict[str, Any]], header: str) -> str:
    if not tasks:
        return f"{header}\n\n_No tasks._"
    lines = [header, ""]
    marks = {"completed": "[x]", "notStarted": "[ ]", "inProgress": "[~]"}
    for t in tasks:
        d = _task_to_dict(t)
        mark = marks.get(d["status"] or "", "[ ]")
        extras = []
        if d["dueDateTime"]:
            extras.append(f"due {d['dueDateTime']}")
        if d["importance"] and d["importance"] != "normal":
            extras.append(d["importance"])
        suffix = f" ({', '.join(extras)})" if extras else ""
        lines.append(f"- {mark} {d['title']}{suffix}  \n  `{d['id']}`")
    return "\n".join(lines)


def _render(payload: dict[str, Any], markdown: str, fmt: ResponseFormat) -> str:
    if fmt == ResponseFormat.JSON:
        return json.dumps(payload, indent=2, ensure_ascii=False)
    return markdown


async def _resolve_list_id(
    list_id: Optional[str], list_name: Optional[str]
) -> str:
    """Return a concrete list id from an explicit id or a display name."""
    if list_id:
        return list_id
    if not list_name:
        raise ValueError("Provide either 'list_id' or 'list_name'.")
    items, _ = await _client().get_collection("/me/todo/lists", limit=100)
    target = list_name.strip().casefold()
    for lst in items:
        if (lst.get("displayName") or "").casefold() == target:
            return lst["id"]
    available = ", ".join(repr(i.get("displayName")) for i in items) or "(none)"
    raise ValueError(f"No list named {list_name!r}. Available lists: {available}.")


# --------------------------------------------------------------------------
# Input models
# --------------------------------------------------------------------------
class _Base(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")


class ListListsInput(_Base):
    limit: int = Field(default=50, description="Max lists to return", ge=1, le=100)
    response_format: ResponseFormat = Field(
        default=ResponseFormat.MARKDOWN, description="'markdown' or 'json'"
    )


class ListTasksInput(_Base):
    list_id: Optional[str] = Field(
        default=None, description="Task list id (from todo_list_lists). Use this OR list_name."
    )
    list_name: Optional[str] = Field(
        default=None, description="Task list display name, e.g. 'Tasks'. Resolved to an id."
    )
    status: Optional[TaskStatus] = Field(
        default=None, description="Filter by task status (e.g. 'notStarted', 'completed')."
    )
    limit: int = Field(default=50, description="Max tasks to return", ge=1, le=200)
    response_format: ResponseFormat = Field(default=ResponseFormat.MARKDOWN)


class CreateTaskInput(_Base):
    title: str = Field(..., description="Task title", min_length=1, max_length=255)
    list_id: Optional[str] = Field(default=None, description="Target list id. Use this OR list_name.")
    list_name: Optional[str] = Field(default=None, description="Target list display name.")
    body: Optional[str] = Field(default=None, description="Optional note/description text.")
    due_date: Optional[str] = Field(
        default=None,
        description="Due date 'YYYY-MM-DD' or datetime 'YYYY-MM-DDTHH:MM:SS'.",
    )
    time_zone: Optional[str] = Field(
        default=None, description="IANA time zone for due_date (defaults to MS_TODO_TIMEZONE)."
    )
    importance: Optional[Importance] = Field(default=None, description="low | normal | high")
    response_format: ResponseFormat = Field(default=ResponseFormat.MARKDOWN)

    @field_validator("title")
    @classmethod
    def _title_not_blank(cls, v: str) -> str:
        if not v.strip():
            raise ValueError("title cannot be empty")
        return v


class UpdateTaskInput(_Base):
    task_id: str = Field(..., description="Id of the task to update", min_length=1)
    list_id: Optional[str] = Field(default=None, description="List id. Use this OR list_name.")
    list_name: Optional[str] = Field(default=None, description="List display name.")
    title: Optional[str] = Field(default=None, description="New title", max_length=255)
    body: Optional[str] = Field(default=None, description="New note/description text.")
    due_date: Optional[str] = Field(default=None, description="New due date/datetime, or '' to clear.")
    time_zone: Optional[str] = Field(default=None, description="IANA time zone for due_date.")
    importance: Optional[Importance] = Field(default=None, description="low | normal | high")
    status: Optional[TaskStatus] = Field(default=None, description="New status.")
    response_format: ResponseFormat = Field(default=ResponseFormat.MARKDOWN)


class TaskRefInput(_Base):
    task_id: str = Field(..., description="Id of the task", min_length=1)
    list_id: Optional[str] = Field(default=None, description="List id. Use this OR list_name.")
    list_name: Optional[str] = Field(default=None, description="List display name.")
    response_format: ResponseFormat = Field(default=ResponseFormat.MARKDOWN)


class CreateListInput(_Base):
    display_name: str = Field(..., description="Name for the new task list", min_length=1, max_length=255)
    response_format: ResponseFormat = Field(default=ResponseFormat.MARKDOWN)


# --------------------------------------------------------------------------
# Tools
# --------------------------------------------------------------------------
@mcp.tool(
    name="todo_list_lists",
    annotations={
        "title": "List Microsoft To Do lists",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def todo_list_lists(params: ListListsInput) -> str:
    """List the signed-in user's Microsoft To Do task lists.

    Use this first to discover list ids and names for the other tools.

    Returns (json): {"count": int, "lists": [{"id", "displayName", "isOwner",
    "isShared", "wellknownListName"}]}. On error: "Error: ...".
    """
    try:
        items, has_more = await _client().get_collection("/me/todo/lists", limit=params.limit)
        lists = [_list_to_dict(x) for x in items]
        payload = {"count": len(lists), "has_more": has_more, "lists": lists}
        md = "# Task lists\n\n" + (
            "\n".join(f"- {l['displayName']}  \n  `{l['id']}`" for l in lists) or "_No lists._"
        )
        return _render(payload, md, params.response_format)
    except Exception as exc:  # noqa: BLE001
        return _format_error(exc)


@mcp.tool(
    name="todo_list_tasks",
    annotations={
        "title": "List tasks in a To Do list",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def todo_list_tasks(params: ListTasksInput) -> str:
    """List tasks within a Microsoft To Do list.

    Identify the list by `list_id` (preferred) or `list_name`. Optionally
    filter by `status` (e.g. 'notStarted' to hide completed items).

    Returns (json): {"list_id", "count", "has_more", "tasks": [task...]} where
    each task has: id, title, status, importance, dueDateTime, dueTimeZone,
    body, createdDateTime, lastModifiedDateTime, completedDateTime.
    On error: "Error: ...".
    """
    try:
        list_id = await _resolve_list_id(params.list_id, params.list_name)
        q: dict[str, Any] = {}
        if params.status:
            q["$filter"] = f"status eq '{params.status.value}'"
        items, has_more = await _client().get_collection(
            f"/me/todo/lists/{list_id}/tasks", params=q, limit=params.limit
        )
        payload = {
            "list_id": list_id,
            "count": len(items),
            "has_more": has_more,
            "tasks": [_task_to_dict(t) for t in items],
        }
        md = _tasks_markdown(items, f"# Tasks ({len(items)})")
        return _render(payload, md, params.response_format)
    except Exception as exc:  # noqa: BLE001
        return _format_error(exc)


@mcp.tool(
    name="todo_create_task",
    annotations={
        "title": "Create a To Do task",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    },
)
async def todo_create_task(params: CreateTaskInput) -> str:
    """Create a new task in a Microsoft To Do list.

    Requires delegated 'Tasks.ReadWrite' (app-only is not supported by Graph
    for creating tasks). Identify the list by `list_id` or `list_name`.

    Returns (json): the created task dict (see todo_list_tasks schema).
    On error: "Error: ...".
    """
    try:
        list_id = await _resolve_list_id(params.list_id, params.list_name)
        body: dict[str, Any] = {"title": params.title}
        if params.body:
            body["body"] = {"content": params.body, "contentType": "text"}
        if params.importance:
            body["importance"] = params.importance.value
        if params.due_date:
            body["dueDateTime"] = {
                "dateTime": _to_graph_datetime(params.due_date),
                "timeZone": params.time_zone or _default_timezone(),
            }
        created = await _client().request(
            "POST", f"/me/todo/lists/{list_id}/tasks", json_body=body
        )
        payload = _task_to_dict(created)
        md = f"Created task **{payload['title']}**\n\n`{payload['id']}`"
        return _render(payload, md, params.response_format)
    except Exception as exc:  # noqa: BLE001
        return _format_error(exc)


@mcp.tool(
    name="todo_update_task",
    annotations={
        "title": "Update a To Do task",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def todo_update_task(params: UpdateTaskInput) -> str:
    """Update fields of an existing task (title, body, due date, importance, status).

    Only the fields you pass are changed. Pass due_date='' to clear the due date.
    Identify the list by `list_id` or `list_name`.

    Returns (json): the updated task dict. On error: "Error: ...".
    """
    try:
        list_id = await _resolve_list_id(params.list_id, params.list_name)
        body: dict[str, Any] = {}
        if params.title is not None:
            body["title"] = params.title
        if params.body is not None:
            body["body"] = {"content": params.body, "contentType": "text"}
        if params.importance is not None:
            body["importance"] = params.importance.value
        if params.status is not None:
            body["status"] = params.status.value
        if params.due_date is not None:
            if params.due_date == "":
                body["dueDateTime"] = None
            else:
                body["dueDateTime"] = {
                    "dateTime": _to_graph_datetime(params.due_date),
                    "timeZone": params.time_zone or _default_timezone(),
                }
        if not body:
            return "Error: nothing to update — pass at least one field to change."
        updated = await _client().request(
            "PATCH", f"/me/todo/lists/{list_id}/tasks/{params.task_id}", json_body=body
        )
        payload = _task_to_dict(updated)
        md = f"Updated task **{payload['title']}** (status: {payload['status']})"
        return _render(payload, md, params.response_format)
    except Exception as exc:  # noqa: BLE001
        return _format_error(exc)


@mcp.tool(
    name="todo_complete_task",
    annotations={
        "title": "Complete a To Do task",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def todo_complete_task(params: TaskRefInput) -> str:
    """Mark a task as completed.

    Convenience wrapper over todo_update_task that sets status='completed'.
    Identify the list by `list_id` or `list_name`.

    Returns (json): the updated task dict. On error: "Error: ...".
    """
    try:
        list_id = await _resolve_list_id(params.list_id, params.list_name)
        updated = await _client().request(
            "PATCH",
            f"/me/todo/lists/{list_id}/tasks/{params.task_id}",
            json_body={"status": "completed"},
        )
        payload = _task_to_dict(updated)
        md = f"Completed task **{payload['title']}**"
        return _render(payload, md, params.response_format)
    except Exception as exc:  # noqa: BLE001
        return _format_error(exc)


@mcp.tool(
    name="todo_delete_task",
    annotations={
        "title": "Delete a To Do task",
        "readOnlyHint": False,
        "destructiveHint": True,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def todo_delete_task(params: TaskRefInput) -> str:
    """Permanently delete a task. This cannot be undone.

    Identify the list by `list_id` or `list_name` and the task by `task_id`.

    Returns: a confirmation string, or "Error: ...".
    """
    try:
        list_id = await _resolve_list_id(params.list_id, params.list_name)
        await _client().request(
            "DELETE", f"/me/todo/lists/{list_id}/tasks/{params.task_id}"
        )
        payload = {"deleted": True, "task_id": params.task_id, "list_id": list_id}
        return _render(payload, f"Deleted task `{params.task_id}`.", params.response_format)
    except Exception as exc:  # noqa: BLE001
        return _format_error(exc)


@mcp.tool(
    name="todo_create_list",
    annotations={
        "title": "Create a To Do list",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    },
)
async def todo_create_list(params: CreateListInput) -> str:
    """Create a new Microsoft To Do task list.

    Returns (json): {"id", "displayName", ...}. On error: "Error: ...".
    """
    try:
        created = await _client().request(
            "POST", "/me/todo/lists", json_body={"displayName": params.display_name}
        )
        payload = _list_to_dict(created)
        md = f"Created list **{payload['displayName']}**\n\n`{payload['id']}`"
        return _render(payload, md, params.response_format)
    except Exception as exc:  # noqa: BLE001
        return _format_error(exc)
