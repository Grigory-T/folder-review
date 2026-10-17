"""Explicit developer-only OSV check; sends public package names/versions only."""

import argparse
import datetime
import json
import urllib.request
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--components", type=Path, default=Path("dist/components.json"))
    parser.add_argument(
        "--output", type=Path, default=Path("dist/dependency-audit.json")
    )
    args = parser.parse_args()
    inventory = json.loads(args.components.read_text())
    packages = inventory["components"]
    queries = [
        {"package": {"name": p["name"], "ecosystem": "PyPI"}, "version": p["version"]}
        for p in packages
    ]
    service = "https://api.osv.dev/v1/querybatch"
    request = urllib.request.Request(
        service,
        data=json.dumps({"queries": queries}).encode(),
        headers={"Content-Type": "application/json"},
    )
    with urllib.request.urlopen(request, timeout=45) as response:
        results = json.load(response)["results"]
    if len(results) != len(packages):
        raise RuntimeError("Incomplete vulnerability service response")
    records = [
        {
            "name": p["name"],
            "version": p["version"],
            "advisory_ids": [v["id"] for v in result.get("vulns", [])],
        }
        for p, result in zip(packages, results)
    ]
    report = {
        "checked_at_utc": datetime.datetime.now(datetime.UTC).isoformat(),
        "service": service,
        "executable_sha256": inventory["executable_sha256"],
        "packages": records,
        "advisories_found": sum(len(p["advisory_ids"]) for p in records),
        "limits": "Known Python-package advisories only; excludes native/OS libraries and unknown vulnerabilities. Not a malware guarantee.",
    }
    args.output.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    print(
        json.dumps(
            {
                "packages_checked": len(records),
                "advisories_found": report["advisories_found"],
            }
        )
    )
    return 1 if report["advisories_found"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
