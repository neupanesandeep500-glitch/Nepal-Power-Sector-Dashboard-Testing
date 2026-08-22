"""
security.py

Cross-cutting security, reliability, and ops hardening for the Flask
server underneath the Dash app. Kept as an independent module (same
pattern as visitor_counter.py / server_state.py) so it can be wired in
with a single `security.init_app(server)` call and doesn't entangle
with the Dash layout/callback code in app.py.

Covers:
- Security response headers (CSP, X-Frame-Options, etc.)
- Hardened session cookie flags
- Timing-safe admin password check + login rate limiting / lockout
- /healthz liveness+readiness endpoint for uptime monitors / Render
- Friendly, non-leaking 404 / 500 error pages (no raw tracebacks to
  visitors — the exception still goes to the server log)
"""

import os
import time
import hmac
import logging
import threading
from collections import defaultdict, deque

from flask import request, jsonify, render_template_string

log = logging.getLogger("security")
if not log.handlers:
    logging.basicConfig(level=logging.INFO)

DEBUG_MODE = os.environ.get("FLASK_DEBUG", "").lower() in ("1", "true", "yes")
FORCE_HTTPS = os.environ.get("FORCE_HTTPS", "1").lower() in ("1", "true", "yes")

# ── Login rate limiting (simple in-memory sliding window) ──────────────────
# Good enough for a single-dyno deployment (Render web service). If this
# ever runs multi-process/multi-instance, swap for a shared store
# (Redis, the same Google Sheet used for the visitor counter, etc.) —
# the call sites (record_failed_login / is_login_locked_out) are the
# only two touch points that would need to change.
_LOGIN_ATTEMPTS_LOCK = threading.Lock()
_LOGIN_ATTEMPTS = defaultdict(deque)  # ip -> deque[timestamp of failed attempt]
LOGIN_MAX_ATTEMPTS = int(os.environ.get("ADMIN_LOGIN_MAX_ATTEMPTS", "5"))
LOGIN_WINDOW_SECONDS = int(os.environ.get("ADMIN_LOGIN_WINDOW_SECONDS", "600"))  # 10 min


def _client_ip():
    # Render (and most PaaS) sit behind a proxy; trust the first
    # forwarded-for hop if present, else fall back to the socket peer.
    fwd = request.headers.get("X-Forwarded-For", "")
    if fwd:
        return fwd.split(",")[0].strip()
    return request.remote_addr or "unknown"


def is_login_locked_out(ip=None):
    """Return (locked: bool, seconds_remaining: int)."""
    ip = ip or _client_ip()
    now = time.time()
    with _LOGIN_ATTEMPTS_LOCK:
        attempts = _LOGIN_ATTEMPTS[ip]
        while attempts and now - attempts[0] > LOGIN_WINDOW_SECONDS:
            attempts.popleft()
        if len(attempts) >= LOGIN_MAX_ATTEMPTS:
            retry_after = int(LOGIN_WINDOW_SECONDS - (now - attempts[0]))
            return True, max(retry_after, 1)
    return False, 0


def record_failed_login(ip=None):
    ip = ip or _client_ip()
    with _LOGIN_ATTEMPTS_LOCK:
        _LOGIN_ATTEMPTS[ip].append(time.time())


def clear_login_attempts(ip=None):
    ip = ip or _client_ip()
    with _LOGIN_ATTEMPTS_LOCK:
        _LOGIN_ATTEMPTS.pop(ip, None)


def safe_password_check(candidate, expected):
    """Constant-time comparison so login timing can't leak how many
    leading characters of the password guess were correct."""
    if candidate is None or expected is None:
        return False
    return hmac.compare_digest(str(candidate).encode("utf-8"),
                                str(expected).encode("utf-8"))


# ── Error pages ──────────────────────────────────────────────────────────
_ERROR_PAGE = """<!DOCTYPE html>
<html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>{{ title }} — Nepal Power Sector Dashboard</title>
<style>
  body { margin:0; font-family:-apple-system,"Segoe UI",Helvetica,Arial,sans-serif;
    background:#0b1730; color:#e6ecfa; display:flex; align-items:center; justify-content:center;
    min-height:100vh; text-align:center; padding:24px; box-sizing:border-box; }
  .box { max-width:480px; }
  .code { font-size:64px; font-weight:800; color:#ffd166; line-height:1; }
  h1 { font-size:20px; margin:12px 0 8px; }
  p { color:#9fb3c8; line-height:1.6; }
  a.btn { display:inline-block; margin-top:18px; padding:10px 22px; border-radius:6px;
    background:linear-gradient(135deg,#1565c0 0%,#0d47a1 100%); color:#fff; text-decoration:none;
    font-weight:600; }
  a.btn:hover { opacity:0.9; }
</style></head>
<body><div class="box">
  <div class="code">{{ code }}</div>
  <h1>{{ title }}</h1>
  <p>{{ message }}</p>
  <a class="btn" href="/">← Back to Dashboard</a>
</div></body></html>"""


def _render_error(code, title, message):
    return render_template_string(_ERROR_PAGE, code=code, title=title, message=message), code


def render_friendly_error(title, message, code=500):
    """Public helper for individual routes (e.g. the NEA/GIS render
    routes in app.py) that catch their own exceptions and want the
    same branded error page instead of a raw traceback string, while
    still logging the real exception server-side via `log.exception`
    at the call site."""
    return _render_error(code, title, message)


def register_error_handlers(server):
    @server.errorhandler(404)
    def _not_found(e):
        return _render_error(404, "Page not found",
                              "The page you're looking for doesn't exist or may have moved.")

    @server.errorhandler(500)
    def _server_error(e):
        # Full traceback still goes to the server log; visitors never see it.
        log.exception("Unhandled server error on %s", request.path)
        return _render_error(500, "Something went wrong",
                              "An unexpected error occurred. It's been logged — please try again shortly.")

    @server.errorhandler(429)
    def _rate_limited(e):
        return _render_error(429, "Too many attempts",
                              "Please wait a bit before trying again.")


# ── Security headers ─────────────────────────────────────────────────────
def _security_headers(resp):
    resp.headers.setdefault("X-Content-Type-Options", "nosniff")
    resp.headers.setdefault("X-Frame-Options", "SAMEORIGIN")
    resp.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
    resp.headers.setdefault("Permissions-Policy",
                             "geolocation=(), microphone=(), camera=(), payment=()")
    if FORCE_HTTPS and not DEBUG_MODE:
        resp.headers.setdefault("Strict-Transport-Security",
                                 "max-age=63072000; includeSubDomains")
    # Deliberately no strict Content-Security-Policy: this app pulls
    # Google Analytics/gtag, Leaflet tiles from OSM, and inline
    # <script>/<style> blocks baked into index_string/templates. A
    # naive CSP would either break those or require a much larger
    # refactor (nonces on every inline block). Left as a documented
    # follow-up rather than shipping a CSP that's security-theater
    # (report-only, wide-open 'unsafe-inline') or that silently breaks
    # the clock/ticker/charts.
    return resp


def _cache_headers_for_static(resp):
    """Long-lived caching for the handful of routes that serve
    effectively-static bytes (vendor JS, uploaded branding images).
    Safe because those routes are content-addressed by filename and
    branding assets are versioned by the admin re-upload, which
    changes the underlying file — browsers revalidate via
    Last-Modified/ETag that Flask's send_from_directory already sets.
    """
    p = request.path
    if p.startswith("/nea-vendor/") or p.startswith("/assets-"):
        # Override (not setdefault): Flask's send_file/send_from_directory
        # sets its own "no-cache" Cache-Control by default when
        # SEND_FILE_MAX_AGE_DEFAULT isn't configured, which would
        # otherwise shadow this and defeat the point of the override.
        resp.headers["Cache-Control"] = "public, max-age=86400, must-revalidate"
    return resp


# ── Health check ─────────────────────────────────────────────────────────
def register_health_check(server, state_getter):
    @server.route("/healthz")
    def _healthz():
        try:
            snapshot = state_getter() or {}
            ok = True
            detail = {
                "status": "ok",
                "license_data_loaded": bool(snapshot.get("loader")),
                "gis_loaded": bool(snapshot.get("gis_loaded")),
                "last_sync": snapshot.get("last_sync"),
            }
        except Exception as e:  # noqa: BLE001 — health check must never 500
            ok = False
            detail = {"status": "degraded", "error": str(e)}
        return jsonify(detail), (200 if ok else 503)


# ── Wire everything into a Flask server in one call ────────────────────────
def init_app(server, state_getter=None):
    """Call once, right after the Flask server is created.

    state_getter: optional zero-arg callable returning server_state.STATE
    (passed as a callable, not the dict itself, so /healthz always
    reads the live state rather than a snapshot taken at import time).
    """
    server.config.update(
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SAMESITE="Lax",
        SESSION_COOKIE_SECURE=FORCE_HTTPS and not DEBUG_MODE,
        PERMANENT_SESSION_LIFETIME=8 * 3600,
        MAX_CONTENT_LENGTH=25 * 1024 * 1024,  # 25MB upload ceiling app-wide
    )

    @server.after_request
    def _apply_headers(resp):
        resp = _security_headers(resp)
        resp = _cache_headers_for_static(resp)
        return resp

    register_error_handlers(server)
    if state_getter is not None:
        register_health_check(server, state_getter)

    # Try gzip/br compression if flask-compress is installed; degrade
    # gracefully (still fully functional, just uncompressed) if not —
    # avoids making this a hard dependency for local dev setups that
    # haven't re-run `pip install -r requirements.txt` yet.
    try:
        from flask_compress import Compress
        Compress(server)
        log.info("flask-compress enabled (gzip/br response compression)")
    except ImportError:
        log.info("flask-compress not installed — responses will be uncompressed. "
                  "Run `pip install -r requirements.txt` to enable it.")

    admin_pw_is_default = os.environ.get("ADMIN_PASSWORD") is None
    if admin_pw_is_default:
        log.warning(
            "ADMIN_PASSWORD is not set in the environment — the admin panel is "
            "using its built-in fallback password. Set ADMIN_PASSWORD before "
            "deploying publicly."
        )

    return server
