from __future__ import annotations

from datetime import datetime

import config


def stage(number: int, title: str) -> None:
    print(f"\n[{number}] {title}")


def detail(message: str) -> None:
    print(f"    {message}")


def result(message: str) -> None:
    print(f"--> {message}")


def run_header(command: str) -> None:
    timestamp = datetime.now().astimezone().isoformat(timespec="seconds")
    print(f"{config.CAMPAIGN_NAME} outreach pipeline | command={command} | {timestamp}")
