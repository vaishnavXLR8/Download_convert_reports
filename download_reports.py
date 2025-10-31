import argparse
import os
import sys
import time
import re
from typing import List, Dict, Tuple

import requests
from requests import Response
from dotenv import load_dotenv

PBI_SCOPE = "https://analysis.windows.net/powerbi/api/.default"
TOKEN_URL_TPL = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
BASE_URL = "https://api.powerbi.com/v1.0/myorg"

CHUNK_SIZE = 1024 * 1024  # 1 MiB
PER_REPORT_TIMEOUT_SECS = 180  # 3 minutes


def env(var: str, required: bool = True) -> str:
    val = os.getenv(var)
    if required and not val:
        raise RuntimeError(f"Missing environment variable: {var}")
    return val or ""


def get_access_token() -> str:
    tenant = env("AZURE_TENANT_ID")
    client_id = env("AZURE_CLIENT_ID")
    client_secret = env("AZURE_CLIENT_SECRET")
    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": PBI_SCOPE,
    }
    url = TOKEN_URL_TPL.format(tenant=tenant)
    resp = requests.post(url, data=data, timeout=30)
    if resp.status_code >= 400:
        raise RuntimeError(f"Token request failed ({resp.status_code}): {resp.text}")
    token = resp.json().get("access_token")
    if not token:
        raise RuntimeError("Token response missing access_token")
    return token


def list_reports(token: str, group_id: str) -> List[Dict[str, str]]:
    url = f"{BASE_URL}/groups/{group_id}/reports"
    resp = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
    if resp.status_code >= 400:
        raise RuntimeError(f"List reports failed ({resp.status_code}): {resp.text}")
    data = resp.json()
    return data.get("value", [])


def parse_filename_from_disposition(disposition: str, fallback: str) -> str:
    if not disposition:
        return fallback
    # Try RFC 5987 filename*
    m = re.search(r"filename\*=UTF-8''([^;]+)", disposition, flags=re.IGNORECASE)
    if m:
        try:
            return requests.utils.unquote(m.group(1))
        except Exception:
            pass
    # Try filename="..."
    m = re.search(r'filename="?([^";]+)"?', disposition, flags=re.IGNORECASE)
    if m:
        return m.group(1)
    return fallback


def sanitize_filename(name: str) -> str:
    # Remove invalid characters for Windows filesystems
    name = re.sub(r"[\\/:*?\"<>|]", "_", name)
    name = name.strip()
    return name or "report"


def download_report(token: str, group_id: str, report: Dict[str, str], out_dir: str) -> Tuple[bool, str]:
    report_id = report.get("id")
    report_name = report.get("name", "report")
    if not report_id:
        return False, "Missing report id"

    url = f"{BASE_URL}/groups/{group_id}/reports/{report_id}/Export"
    headers = {"Authorization": f"Bearer {token}"}

    start = time.time()
    try:
        with requests.get(url, headers=headers, stream=True, timeout=30) as resp:
            if resp.status_code >= 400:
                return False, f"Export failed ({resp.status_code}): {resp.text}"

            disp = resp.headers.get("Content-Disposition")
            fallback_name = sanitize_filename(report_name) + ".pbix"
            filename = parse_filename_from_disposition(disp, fallback_name)
            filename = sanitize_filename(filename)

            os.makedirs(out_dir, exist_ok=True)
            path = os.path.join(out_dir, filename)

            with open(path, "wb") as f:
                for chunk in resp.iter_content(chunk_size=CHUNK_SIZE):
                    if chunk:
                        f.write(chunk)
                    # Enforce per-report timeout
                    if time.time() - start > PER_REPORT_TIMEOUT_SECS:
                        f.close()
                        os.remove(path)
                        return False, "Timeout after 3 minutes"
        return True, "OK"
    except requests.Timeout:
        return False, "HTTP request timed out"
    except Exception as e:
        return False, str(e)


def main() -> int:
    load_dotenv()  # load from .env if present

    parser = argparse.ArgumentParser(description="Download all PBIX reports from a Power BI workspace")
    parser.add_argument("--group-id", "-g", required=True, help="Power BI workspace (group) ID")
    parser.add_argument("--output", "-o", default="downloads", help="Output directory (default: ./downloads)")
    args = parser.parse_args()

    try:
        token = get_access_token()
    except Exception as e:
        print(f"Auth error: {e}", file=sys.stderr)
        return 1

    try:
        reports = list_reports(token, args.group_id)
    except Exception as e:
        print(f"Failed to list reports: {e}", file=sys.stderr)
        return 1

    if not reports:
        print("No reports found in this workspace.")
        return 0

    print(f"Found {len(reports)} reports. Downloading to: {os.path.abspath(args.output)}\n")
    successes = []
    failures = []

    for i, r in enumerate(reports, start=1):
        name = r.get("name", r.get("id", "report"))
        print(f"[{i}/{len(reports)}] {name} ...", end=" ")
        ok, msg = download_report(token, args.group_id, r, args.output)
        if ok:
            print("done")
            successes.append(name)
        else:
            print("FAILED")
            failures.append((name, msg))

    print("\nSummary:")
    print(f"  Success: {len(successes)}")
    print(f"  Failed : {len(failures)}")
    if failures:
        for name, reason in failures:
            print(f"   - {name}: {reason}")

    return 0 if not failures else 2


if __name__ == "__main__":
    raise SystemExit(main())
