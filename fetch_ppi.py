#!/usr/bin/env python3
"""Fetch Pentagon Pizza Index from pizzint.watch and write ppi_latest.json.

This runs hourly via .github/workflows/hourly_ppi.yml. The dashboard page
fetches ppi_latest.json client-side, so updating this file is enough to
refresh the PPI card without regenerating the whole dashboard HTML.
"""
import json
import os
import ssl
from datetime import datetime, timezone, timedelta
from urllib.request import Request, urlopen

API_URL = "https://www.pizzint.watch/api/dashboard-data"
OUT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ppi_latest.json")


def fetch():
    req = Request(
        API_URL,
        headers={
            "User-Agent": "Mozilla/5.0 (tw-stock-dashboard PPI bot)",
            "Accept": "application/json",
        },
    )
    ctx = ssl.create_default_context()
    with urlopen(req, timeout=30, context=ctx) as resp:
        return json.loads(resp.read().decode("utf-8"))


def main():
    data = fetch()
    details = data.get("defcon_details") or {}

    out = {
        "overall_index": data.get("overall_index"),
        "defcon_level": data.get("defcon_level"),
        "defcon_severity": details.get("defcon_severity_decimal"),
        "active_spikes": data.get("active_spikes"),
        "has_active_spikes": data.get("has_active_spikes"),
        "raw_index": details.get("raw_index"),
        "smoothed_index": details.get("smoothed_index"),
        "open_places": details.get("open_places"),
        "total_places": details.get("total_places"),
        "intensity_score": details.get("intensity_score"),
        "breadth_score": details.get("breadth_score"),
        "source_time": details.get("at_time"),
        "fetched_at": datetime.now(timezone(timedelta(hours=8)))
        .isoformat(timespec="seconds"),
    }

    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(out, f, indent=2, ensure_ascii=False)
        f.write("\n")

    print(
        f"✅ PPI: overall={out['overall_index']} defcon={out['defcon_level']} "
        f"spikes={out['active_spikes']} src={out['source_time']}"
    )


if __name__ == "__main__":
    main()
