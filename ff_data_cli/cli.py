from __future__ import annotations

import json
from pathlib import Path
from typing import Optional

import argparse
import csv
from io import StringIO
from urllib.request import urlopen


def echo(text: str) -> None:
    __builtins__.print(text)  # type: ignore[attr-defined]


def write_output(data: object, output: Optional[Path]) -> None:
    if output is None or str(output) == "-":
        echo(json.dumps(data, ensure_ascii=False))
        return
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(data, ensure_ascii=False))
    echo(f"Wrote {len(data) if isinstance(data, (list, dict)) else 'object'} to {output}")


def sleeper_players(output: Optional[Path]):
    """Fetch all Sleeper players as JSON."""
    url = "https://api.sleeper.app/v1/players/nfl"
    with urlopen(url, timeout=60) as resp:
        data = resp.read().decode("utf-8")
    players = json.loads(data)
    write_output(players, output)


def dynastyprocess_player_ids(output: Optional[Path]):
    """Fetch DynastyProcess nflverse player IDs CSV and emit as JSON rows."""
    # Primary official location (nflverse):
    # https://github.com/dynastyprocess/data/tree/master/files
    # Raw CSV URL is versioned, but there is a stable file name
    raw_url = "https://raw.githubusercontent.com/dynastyprocess/data/master/files/playerids.csv"
    with urlopen(raw_url, timeout=60) as resp:
        csv_text = resp.read().decode("utf-8")
    reader = csv.DictReader(StringIO(csv_text))
    rows = [dict(r) for r in reader]
    write_output(rows, output)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(prog="ff", description="Fantasy Football Data CLI")
    subparsers = parser.add_subparsers(dest="command", required=True)

    sp1 = subparsers.add_parser("sleeper-players", help="Fetch all Sleeper players")
    sp1.add_argument("-o", "--output", type=Path, default=None)

    sp2 = subparsers.add_parser("dp-player-ids", help="Fetch DynastyProcess player ID map")
    sp2.add_argument("-o", "--output", type=Path, default=None)

    args = parser.parse_args()
    if args.command == "sleeper-players":
        sleeper_players(args.output)
    elif args.command == "dp-player-ids":
        dynastyprocess_player_ids(args.output)

