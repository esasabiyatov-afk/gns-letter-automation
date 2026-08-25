from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Iterable
from urllib.parse import urljoin, urlparse
from urllib.parse import parse_qs


ALWAYS_DENIED_MARKERS = (
    "/delete",
    "/remove",
    "/restore",
    "/close",
)

APPROVAL_REQUIRED_MARKERS = (
    "/add",
    "/create",
    "/edit",
    "/update",
    "/authorize",
    "/authorization",
    "/change",
    "/save",
    "/submit",
)


@dataclass
class ActionApproval:
    """One-time, short-lived approval for one exact browser/API action."""

    method: str
    path: str
    allowed_fields: frozenset[str]
    expires_at: datetime
    reason: str
    consumed: bool = False

    def consume(self, method: str, path: str, field_names: Iterable[str]) -> None:
        if self.consumed:
            raise PermissionError("Action approval has already been used")
        now = datetime.now(timezone.utc)
        expires_at = self.expires_at
        if expires_at.tzinfo is None:
            expires_at = expires_at.replace(tzinfo=timezone.utc)
        if now >= expires_at:
            raise PermissionError("Action approval has expired")
        if method.upper() != self.method.upper() or path != self.path:
            raise PermissionError("Action does not match the approved method and path")
        unexpected = set(field_names) - set(self.allowed_fields)
        if unexpected:
            raise PermissionError(f"Fields were not approved: {', '.join(sorted(unexpected))}")
        self.consumed = True


@dataclass(frozen=True)
class ReadOnlyPolicy:
    """Fail-closed policy with narrowly scoped, one-time write approvals."""

    base_url: str
    allowed_post_paths: frozenset[str] = field(default_factory=frozenset)
    always_denied_markers: tuple[str, ...] = ALWAYS_DENIED_MARKERS
    approval_required_markers: tuple[str, ...] = APPROVAL_REQUIRED_MARKERS
    view_only_get_routes: tuple[tuple[str, tuple[tuple[str, str], ...]], ...] = ()

    def validate(
        self,
        method: str,
        url_or_path: str,
        *,
        field_names: Iterable[str] = (),
        approval: ActionApproval | None = None,
    ) -> str:
        absolute = urljoin(self.base_url.rstrip("/") + "/", url_or_path)
        base = urlparse(self.base_url)
        target = urlparse(absolute)
        if (target.scheme, target.netloc) != (base.scheme, base.netloc):
            raise PermissionError("Read-only adapter cannot access another host")

        normalized_method = method.upper().strip()
        path = target.path
        path_lower = path.lower()
        if normalized_method == "DELETE" or any(marker in path_lower for marker in self.always_denied_markers):
            raise PermissionError(f"Route is always blocked: {path}")

        # Some legacy MVC display pages unfortunately contain words such as
        # "Edit" in their GET route.  They are allowed only when both the
        # exact path and the required view-mode query values match.
        if normalized_method == "GET":
            query = {key.casefold(): [value.casefold() for value in values]
                     for key, values in parse_qs(target.query).items()}
            for safe_path, requirements in self.view_only_get_routes:
                if path != safe_path:
                    continue
                if all(expected.casefold() in query.get(key.casefold(), [])
                       for key, expected in requirements):
                    return absolute

        requires_approval = (
            normalized_method not in {"GET", "HEAD"}
            or any(marker in path_lower for marker in self.approval_required_markers)
        )
        if not requires_approval:
            return absolute
        if normalized_method == "POST" and path in self.allowed_post_paths:
            return absolute
        if approval is None:
            raise PermissionError(f"HTTP {normalized_method} {path} requires one-time user approval")
        approval.consume(normalized_method, path, field_names)
        return absolute
