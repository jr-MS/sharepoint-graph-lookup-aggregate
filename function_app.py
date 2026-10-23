"""One HTTP call that returns SharePoint lookup fields Graph omits on a list read.

Collection GET /sites/{site}/lists/{list}/items drops person and lookup columns
after SharePoint's 12-column list-view threshold. GET /items/{id} returns them.
This function reads the item ids, then reads each item through Graph $batch.
"""

from __future__ import annotations

import json
import os
import time
from concurrent.futures import ThreadPoolExecutor
from typing import Any

import azure.functions as func
import requests
from azure.identity import DefaultAzureCredential

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)

BATCH_SIZE = 20
BATCH_WORKERS = 3
MAX_TOP = 2000
GRAPH = "https://graph.microsoft.com/v1.0"

DEFAULT_FIELDS = [
    "Title",
    "Owner",
    "OwnerLookupId",
    "Category",
    "CategoryLookupId",
]


def _json(payload: dict[str, Any], status: int = 200) -> func.HttpResponse:
    return func.HttpResponse(
        json.dumps(payload, ensure_ascii=False),
        status_code=status,
        mimetype="application/json",
    )


def _body(req: func.HttpRequest) -> dict[str, Any]:
    try:
        data = req.get_json()
    except ValueError:
        data = None
    return data if isinstance(data, dict) else {}


def _graph_headers() -> dict[str, str]:
    token = os.environ.get("GRAPH_TOKEN", "").strip()
    if not token:
        token = DefaultAzureCredential().get_token("https://graph.microsoft.com/.default").token
    return {"Authorization": f"Bearer {token}", "Accept": "application/json"}


def _request_with_retry(method: str, url: str, headers: dict[str, str], **kwargs: Any) -> requests.Response:
    delay = 1.0
    timeout = kwargs.pop("timeout", 90)
    response: requests.Response | None = None
    for _ in range(5):
        response = requests.request(method, url, headers=headers, timeout=timeout, **kwargs)
        if response.status_code == 429 or response.status_code >= 500:
            time.sleep(float(response.headers.get("Retry-After", delay)))
            delay = min(delay * 2, 30)
            continue
        return response
    assert response is not None
    return response


def _flatten(raw: dict[str, Any], selected: list[str]) -> dict[str, Any]:
    fields = raw.get("fields") or {}
    wanted = set(selected)
    row: dict[str, Any] = {"id": raw.get("id")}
    for key, value in fields.items():
        base = key.removesuffix("LookupId").removesuffix("Email")
        if key in wanted or base in wanted:
            row[key] = value
    for name in selected:
        row.setdefault(name, None)
    return row


def _mock_rows(selected: list[str], top: int) -> list[dict[str, Any]]:
    rows = []
    for item_id in range(1, top + 1):
        row: dict[str, Any] = {"id": str(item_id)}
        for name in selected:
            if name.endswith("LookupId"):
                row[name] = str(item_id)
            elif name == "Title":
                row[name] = f"Item {item_id:04d}"
            else:
                row[name] = f"{name} {item_id}"
        rows.append(row)
    return rows


def _graph_aggregate(site_id: str, list_id: str, selected: list[str], top: int) -> dict[str, Any]:
    headers = _graph_headers()
    select = ",".join(selected)
    ids: list[str] = []
    url = f"{GRAPH}/sites/{site_id}/lists/{list_id}/items?$select=id&$top={min(top, 999)}"
    while url and len(ids) < top:
        page = _request_with_retry("GET", url, headers, timeout=60)
        page.raise_for_status()
        payload = page.json()
        ids.extend(str(row["id"]) for row in payload.get("value", []))
        url = payload.get("@odata.nextLink")
    ids = ids[:top]

    def fetch_chunk(chunk: list[str]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
        batch_body = {
            "requests": [
                {
                    "id": item_id,
                    "method": "GET",
                    "url": (
                        f"/sites/{site_id}/lists/{list_id}/items/{item_id}"
                        f"?$expand=fields($select={select})"
                    ),
                }
                for item_id in chunk
            ]
        }
        response = _request_with_retry(
            "POST",
            f"{GRAPH}/$batch",
            {**headers, "Content-Type": "application/json"},
            json=batch_body,
            timeout=90,
        )
        response.raise_for_status()
        items: list[dict[str, Any]] = []
        failures: list[dict[str, Any]] = []
        for row in response.json().get("responses", []):
            if row.get("status") == 200:
                items.append(_flatten(row.get("body") or {}, selected))
            else:
                failures.append(
                    {
                        "id": row.get("id"),
                        "status": row.get("status"),
                        "error": (row.get("body") or {}).get("error"),
                    }
                )
        return items, failures

    chunks = [ids[i : i + BATCH_SIZE] for i in range(0, len(ids), BATCH_SIZE)]
    items: list[dict[str, Any]] = []
    failures: list[dict[str, Any]] = []
    workers = min(BATCH_WORKERS, max(len(chunks), 1))
    with ThreadPoolExecutor(max_workers=workers) as pool:
        for chunk_items, chunk_failures in pool.map(fetch_chunk, chunks):
            items.extend(chunk_items)
            failures.extend(chunk_failures)
    return {
        "items": items,
        "failures": failures,
        "idCount": len(ids),
        "batchCount": len(chunks),
        "batchSize": BATCH_SIZE,
    }


@app.route(route="aggregate", methods=["POST"])
def aggregate(req: func.HttpRequest) -> func.HttpResponse:
    body = _body(req)
    mode = str(body.get("mode") or "graph").lower()
    selected = body.get("fields") or DEFAULT_FIELDS
    if not isinstance(selected, list) or not all(isinstance(name, str) for name in selected):
        return _json({"error": "fields must be an array of column internal names."}, 400)
    try:
        top = int(body.get("top") or (1500 if mode == "graph" else 2))
    except (TypeError, ValueError):
        return _json({"error": "top must be an integer."}, 400)
    top = max(1, min(top, MAX_TOP))
    started = time.perf_counter()

    if mode == "mock":
        rows = _mock_rows(selected, top)
        graph = {
            "items": rows,
            "failures": [],
            "idCount": len(rows),
            "batchCount": 0,
            "batchSize": BATCH_SIZE,
        }
        site_id = None
        list_id = None
    elif mode == "graph":
        site_id = body.get("siteId") or os.environ.get("GRAPH_SITE_ID")
        list_id = body.get("listId") or os.environ.get("GRAPH_LIST_ID")
        if not site_id or not list_id:
            return _json(
                {"error": "mode=graph requires siteId and listId, or app settings GRAPH_SITE_ID and GRAPH_LIST_ID."},
                400,
            )
        try:
            graph = _graph_aggregate(site_id, list_id, selected, top)
        except requests.HTTPError as exc:
            detail = exc.response.text[:2000] if exc.response is not None else str(exc)
            status = exc.response.status_code if exc.response is not None else 502
            return _json({"error": "Graph aggregation failed", "graphStatus": status, "detail": detail}, 502)
        except Exception as exc:
            return _json({"error": "Graph aggregation failed", "detail": str(exc)}, 502)
        rows = graph["items"]
    else:
        return _json({"error": "mode must be 'graph' or 'mock'."}, 400)

    elapsed_ms = round((time.perf_counter() - started) * 1000, 1)
    return _json(
        {
            "status": "ok",
            "elapsedMs": elapsed_ms,
            "summary": {
                "mode": mode,
                "itemCount": len(rows),
                "failureCount": len(graph["failures"]),
                "ownerPresent": all(row.get("Owner") not in (None, "") for row in rows) if rows else False,
            },
            "result": {
                "mode": mode,
                "siteId": site_id,
                "listId": list_id,
                "itemCount": len(rows),
                "idCount": graph["idCount"],
                "batchCount": graph["batchCount"],
                "batchSize": graph["batchSize"],
                "failures": graph["failures"],
                "items": rows,
            },
        }
    )
