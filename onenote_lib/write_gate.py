"""Confirmation gate and backups for page writes.

A write tool cannot ask the user anything itself, so the gate makes the
review step structural: onenote_update_page only accepts a token that
onenote_diff_page issued, and a token is bound to the exact change it was
shown for. Editing the markdown after seeing the diff invalidates the token.
"""

import difflib
import hashlib
import secrets
import time
from dataclasses import dataclass
from pathlib import Path

from onenote_lib.config import config

TOKEN_TTL_SECONDS = 600


@dataclass
class _Pending:
    page_id: str
    fingerprint: str
    expires_at: float


_pending: dict[str, _Pending] = {}


def _fingerprint(page_id: str, markdown: str, mode: str) -> str:
    payload = f"{page_id}\0{mode}\0{markdown}".encode()
    return hashlib.sha256(payload).hexdigest()


def issue_token(page_id: str, markdown: str, mode: str) -> str:
    """Issue a token for one specific proposed change."""
    _purge_expired()
    token = secrets.token_urlsafe(16)
    _pending[token] = _Pending(
        page_id=page_id,
        fingerprint=_fingerprint(page_id, markdown, mode),
        expires_at=time.time() + TOKEN_TTL_SECONDS,
    )
    return token


def consume_token(token: str, page_id: str, markdown: str, mode: str) -> None:
    """Validate and burn a token.

    Raises:
        PermissionError: The token is missing, expired, already used, or was
            issued for different content.
    """
    _purge_expired()
    pending = _pending.get(token)
    if pending is None:
        raise PermissionError(
            "No valid confirm_token. Call onenote_diff_page first, show the diff "
            "to the user, and pass the token it returns."
        )
    if pending.page_id != page_id or pending.fingerprint != _fingerprint(page_id, markdown, mode):
        # Deliberately left valid: a mismatched attempt must not invalidate a
        # token the user legitimately approved.
        raise PermissionError(
            "This confirm_token was issued for a different page or different "
            "content. Call onenote_diff_page again for the change you intend."
        )
    del _pending[token]


def _purge_expired() -> None:
    now = time.time()
    for token in [t for t, p in _pending.items() if p.expires_at < now]:
        del _pending[token]


def render_diff(before: str, after: str, page_name: str) -> str:
    """Unified diff of the page's markdown, before and after."""
    lines = difflib.unified_diff(
        before.splitlines(),
        after.splitlines(),
        fromfile=f"{page_name} (current)",
        tofile=f"{page_name} (proposed)",
        lineterm="",
    )
    return "\n".join(lines)


def backup_page(page_id: str, page_xml: str) -> Path:
    """Save the pre-edit page XML. Returns the file written."""
    safe_id = "".join(c for c in page_id if c.isalnum() or c in "-_")[-40:]
    directory = config.resolved_backup_dir() / safe_id
    directory.mkdir(parents=True, exist_ok=True)
    path = directory / f"{time.strftime('%Y%m%dT%H%M%S')}.xml"
    path.write_text(page_xml, encoding="utf-8")
    return path
