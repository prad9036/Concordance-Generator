import os
import json
import time
import shutil
import subprocess
import requests

DEBUG_PORT = 9222
DEBUG_URL = f"http://127.0.0.1:{DEBUG_PORT}"
JSON_URL = f"{DEBUG_URL}/json"
VERSION_URL = f"{DEBUG_URL}/json/version"

PERPLEXITY_URL = "https://www.perplexity.ai/"

TARGET_COOKIES = {
    "pplx.visitor-id",
    "pplx.session-id",
    "pplx.metadata",
    "pplx.edge-vid",
    "pplx.edge-sid",
    "cf_clearance",
    "g_state",
    "_dd_s_v2",
    "__cf_bm",
    "__cflb",
}

WAIT_SECONDS = 60
POLL_INTERVAL = 2


# =========================================================
# Find Edge
# =========================================================

def find_edge():

    paths = [
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        os.path.expandvars(
            r"%LOCALAPPDATA%\Microsoft\Edge\Application\msedge.exe"
        ),
    ]

    for path in paths:
        if os.path.isfile(path):
            return path

    path = shutil.which("msedge.exe")

    if path:
        return path

    raise FileNotFoundError(
        "Microsoft Edge (msedge.exe) was not found."
    )


# =========================================================
# Check CDP
# =========================================================

def edge_debug_running():

    try:

        r = requests.get(
            VERSION_URL,
            timeout=1
        )

        return r.status_code == 200

    except requests.RequestException:

        return False


# =========================================================
# Start Edge
# =========================================================

def start_edge():

    edge = find_edge()

    print("Edge CDP is not running.")
    print("Starting Edge:")
    print(edge)

    profile = os.path.join(
        os.environ.get("TEMP", r"C:\Temp"),
        "edge-perplexity-cdp"
    )

    os.makedirs(profile, exist_ok=True)

    command = [
        edge,

        f"--user-data-dir={profile}",

        "--inprivate",

        f"--remote-debugging-port={DEBUG_PORT}",

        f"--remote-allow-origins=http://127.0.0.1:{DEBUG_PORT}",

        PERPLEXITY_URL,
    ]

    subprocess.Popen(
        command,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )

    print("Waiting for Edge CDP...")

    for i in range(60):

        if edge_debug_running():

            print(
                f"Edge CDP is ready after {i + 1} second(s)."
            )

            return

        time.sleep(1)

    raise RuntimeError(
        "Edge CDP did not become available."
    )


# =========================================================
# Get tabs
# =========================================================

def get_tabs():

    r = requests.get(
        JSON_URL,
        timeout=5
    )

    r.raise_for_status()

    return r.json()


# =========================================================
# Find Perplexity
# =========================================================

def find_perplexity_tab():

    for tab in get_tabs():

        url = tab.get("url", "")

        if (
            "perplexity.ai" in url.lower()
            and tab.get("webSocketDebuggerUrl")
        ):

            return tab

    return None


# =========================================================
# CDP connection
# =========================================================

class CDP:

    def __init__(self, ws):

        self.ws = ws
        self.counter = 0

    def command(self, method, params=None):

        self.counter += 1

        request_id = self.counter

        self.ws.send(
            json.dumps({
                "id": request_id,
                "method": method,
                "params": params or {},
            })
        )

        while True:

            response = json.loads(
                self.ws.recv()
            )

            if response.get("id") == request_id:

                if "error" in response:

                    raise RuntimeError(
                        response["error"]
                    )

                return response


# =========================================================
# Main
# =========================================================

def main():

    # -----------------------------------------------------
    # Start/check Edge
    # -----------------------------------------------------

    if edge_debug_running():

        print(
            f"Edge is already running with CDP "
            f"on port {DEBUG_PORT}."
        )

    else:

        start_edge()

    # -----------------------------------------------------
    # Wait for Perplexity tab
    # -----------------------------------------------------

    print()
    print("Looking for Perplexity tab...")

    tab = None

    for i in range(30):

        tab = find_perplexity_tab()

        if tab:

            break

        time.sleep(1)

    if not tab:

        raise RuntimeError(
            "Perplexity tab was not found."
        )

    print()
    print("Found:")
    print(tab.get("title", ""))
    print(tab.get("url", ""))

    # -----------------------------------------------------
    # Connect to tab
    # -----------------------------------------------------

    from websocket import create_connection

    ws_url = tab["webSocketDebuggerUrl"]

    print()
    print("Connecting to:")
    print(ws_url)

    ws = create_connection(
        ws_url,
        timeout=15,
        origin=f"http://127.0.0.1:{DEBUG_PORT}",
    )

    cdp = CDP(ws)

    try:

        # -------------------------------------------------
        # Enable Network domain
        # -------------------------------------------------

        cdp.command("Network.enable")

        # -------------------------------------------------
        # Wait for cookies
        # -------------------------------------------------

        print()
        print("=" * 60)
        print("WAITING FOR COOKIES")
        print("=" * 60)

        found = {}

        start = time.time()
        last_names = set()

        while time.time() - start < WAIT_SECONDS:

            elapsed = int(time.time() - start)

            result = cdp.command(
                "Network.getAllCookies"
            )

            cookies = result["result"]["cookies"]

            current = {}

            for cookie in cookies:

                name = cookie["name"]

                if name in TARGET_COOKIES:

                    current[name] = cookie["value"]

            # Show newly appearing cookies
            new_names = set(current) - last_names

            for name in sorted(new_names):

                print()
                print(
                    f"[{elapsed:02d}s] COOKIE FOUND: {name}"
                )
                print(current[name])

            last_names = set(current)

            found.update(current)

            # Stop only when all requested cookies exist
            missing = TARGET_COOKIES - set(found)

            print(
                f"\r[{elapsed:02d}s] "
                f"Found {len(found)}/{len(TARGET_COOKIES)} "
                f"cookies; waiting...",
                end="",
                flush=True
            )

            if not missing:

                print()
                print()
                print("All target cookies found.")

                break

            time.sleep(POLL_INTERVAL)

        print()

        # -------------------------------------------------
        # Final result
        # -------------------------------------------------

        print()
        print("=" * 60)
        print("FINAL COOKIE STATUS")
        print("=" * 60)

        for name in sorted(TARGET_COOKIES):

            if name in found:

                print()
                print(f"{name}=")
                print(found[name])

            else:

                print()
                print(f"{name}=<not found>")

        # -------------------------------------------------
        # Cookie string
        # -------------------------------------------------

        cookie_string = "; ".join(
            f"{name}={value}"
            for name, value in found.items()
        )

        print()
        print("=" * 60)
        print("COOKIE STRING")
        print("=" * 60)
        print()
        print(cookie_string)

        # -------------------------------------------------
        # Save
        # -------------------------------------------------

        output = {
            "cookies": found,
            "cookie_string": cookie_string,
            "missing": sorted(
                TARGET_COOKIES - set(found)
            ),
        }

        with open(
            "ppCookies.txt",
            "w",
            encoding="utf-8",
        ) as f:

            json.dump(
                output,
                f,
                indent=2,
                ensure_ascii=False,
            )

        print()
        print(
            "Saved: ppCookies.txt"
        )

    finally:

        ws.close()


if __name__ == "__main__":
    main()
