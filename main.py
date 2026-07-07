#!/usr/bin/env python3
"""统一入口：按任务类型分发到对应生成脚本。"""

from __future__ import annotations

import argparse
import importlib
from collections.abc import Callable


TaskHandler = Callable[[], None]

TASKS: dict[str, tuple[str, str]] = {
    "contracts": ("生成期货合约导入文件", "scripts.generate_futures_contracts_akshare"),
    "contract-prices": ("生成期货合约价导入文件", "scripts.generate_futures_contract_prices_akshare"),
    "trades": ("生成期货成交记录", "scripts.generate_futures_trades"),
    "spot-prices": ("生成现货市场价导入文件", "scripts.generate_spot_market_price_from_futures_akshare"),
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="期货导入文件生成工具统一入口")
    parser.add_argument(
        "task",
        nargs="?",
        choices=TASKS.keys(),
        help="任务类型；不传则进入菜单选择",
    )
    return parser.parse_args()


def choose_task() -> str:
    print("请选择要执行的任务：")
    task_keys = list(TASKS.keys())
    for idx, key in enumerate(task_keys, start=1):
        print(f"{idx}. {TASKS[key][0]} ({key}，可输入 {idx} 或 {key})")

    while True:
        value = input("请输入序号或任务名: ").strip()
        if value in TASKS:
            return value
        if value.isdigit():
            index = int(value)
            if 1 <= index <= len(task_keys):
                return task_keys[index - 1]
        print("输入无效，请重新输入。")


def main() -> None:
    args = parse_args()
    task_key = args.task or choose_task()
    _, module_name = TASKS[task_key]
    handler: TaskHandler = importlib.import_module(module_name).main
    handler()


if __name__ == "__main__":
    main()
