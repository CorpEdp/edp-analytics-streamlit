"""Desktop wrapper for the hosted EDP Analytics Streamlit application."""

from __future__ import annotations

import html
import urllib.error
import urllib.request

import webview


APP_TITLE = "EDP Analytics"
APP_URL = "https://edp-reports-20260915.streamlit.app/"
WINDOW_WIDTH = 1400
WINDOW_HEIGHT = 900
CONNECT_TIMEOUT_SECONDS = 15


def _check_url() -> str | None:
    """Return a user-facing error when the hosted app cannot be reached."""
    request = urllib.request.Request(
        APP_URL,
        headers={"User-Agent": f"{APP_TITLE} desktop wrapper"},
    )
    try:
        with urllib.request.urlopen(request, timeout=CONNECT_TIMEOUT_SECONDS) as response:
            if response.status >= 400:
                return f"The hosted application returned HTTP {response.status}."
    except urllib.error.HTTPError as error:
        return f"The hosted application returned HTTP {error.code}."
    except urllib.error.URLError as error:
        reason = getattr(error, "reason", error)
        return f"The hosted application could not be reached: {reason}."
    except TimeoutError:
        return "The connection timed out. Check your internet connection and try again."
    except OSError as error:
        return f"The hosted application could not be reached: {error}."
    return None


def _error_page(message: str) -> str:
    safe_message = html.escape(message)
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{APP_TITLE}</title>
  <style>
    :root {{ color-scheme: light; font-family: Segoe UI, sans-serif; }}
    body {{ background: #f4f7fb; color: #172033; margin: 0; min-height: 100vh;
           display: grid; place-items: center; }}
    main {{ background: white; border: 1px solid #d9e1ec; border-radius: 12px;
            box-shadow: 0 12px 32px rgba(23, 32, 51, .10); max-width: 620px;
            padding: 36px; margin: 24px; }}
    h1 {{ margin: 0 0 12px; font-size: 26px; }}
    p {{ line-height: 1.55; margin: 10px 0; }}
    .detail {{ color: #52627a; font-size: 14px; word-break: break-word; }}
  </style>
</head>
<body><main>
  <h1>EDP Analytics is unavailable</h1>
  <p>{safe_message}</p>
  <p class="detail">Confirm that you are connected to the internet, then close and reopen the application.</p>
</main></body>
</html>"""


def main() -> None:
    connection_error = _check_url()
    window_args = {
        "title": APP_TITLE,
        "width": WINDOW_WIDTH,
        "height": WINDOW_HEIGHT,
        "resizable": True,
    }
    if connection_error:
        window_args["html"] = _error_page(connection_error)
    else:
        window_args["url"] = APP_URL

    webview.create_window(**window_args)
    webview.start(debug=False)


if __name__ == "__main__":
    main()