#!/usr/bin/env python3
"""按模板生成期货成交记录模拟数据（输出 xlsx）。"""

from __future__ import annotations

import random
import re
import secrets
import zipfile
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path
import xml.etree.ElementTree as ET

HEADERS = [
    "结算日期",
    "成交编号",
    "成交日期",
    "成交时间",
    "操作账户",
    "合约代码",
    "买卖方向",
    "开平标志",
    "成交手数",
    "成交价格",
    "成交手续费",
    "持仓类型",
    "委托编号",
]

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent

TRADING_SESSIONS = [
    (time(9, 0, 0), time(10, 15, 0)),
    (time(10, 30, 0), time(11, 30, 0)),
    (time(13, 30, 0), time(15, 0, 0)),
    (time(21, 0, 0), time(23, 0, 0)),
]


@dataclass(frozen=True)
class ProductProfile:
    base_price: float
    min_qty: int
    max_qty: int
    fee_rate: float
    tick_size: float


@dataclass(frozen=True)
class SymbolConfig:
    code: str
    profile: ProductProfile


SHFE_PRODUCT_PROFILES: dict[str, ProductProfile] = {
    "rb": ProductProfile(3560.0, 1, 10, 0.00008, 1.0),
    "ag": ProductProfile(7120.0, 1, 8, 0.00006, 1.0),
    "cu": ProductProfile(78120.0, 1, 4, 0.00005, 10.0),
    "al": ProductProfile(20150.0, 1, 6, 0.00005, 5.0),
    "zn": ProductProfile(22600.0, 1, 5, 0.00005, 5.0),
    "ni": ProductProfile(132000.0, 1, 3, 0.00006, 10.0),
    "au": ProductProfile(560.0, 1, 10, 0.00004, 0.02),
    "hc": ProductProfile(3720.0, 1, 8, 0.00008, 1.0),
    "ru": ProductProfile(14500.0, 1, 8, 0.00006, 5.0),
    "bu": ProductProfile(3540.0, 1, 10, 0.00008, 1.0),
    "fu": ProductProfile(3140.0, 1, 12, 0.00008, 1.0),
    "sn": ProductProfile(272000.0, 1, 3, 0.00005, 10.0),
    "pb": ProductProfile(17500.0, 1, 5, 0.00005, 5.0),
    "sp": ProductProfile(6350.0, 1, 8, 0.00006, 2.0),
    "ss": ProductProfile(12900.0, 1, 6, 0.00005, 5.0),
}


@dataclass(frozen=True)
class UserInputs:
    rows: int
    start_date: str
    end_date: str
    contracts: str
    accounts: list[str]
    output: str


def prompt_with_default(prompt_text: str, default: str) -> str:
    value = input(f"{prompt_text} [{default}]: ").strip()
    return value if value else default


def read_user_inputs() -> UserInputs:
    today_str = date.today().strftime("%Y-%m-%d")
    print("请输入模拟参数（直接回车使用默认值）")
    rows_text = prompt_with_default("生成记录条数", "50")
    start_date = prompt_with_default("开始日期(YYYY-MM-DD)", today_str)
    end_date = prompt_with_default("结束日期(YYYY-MM-DD)", today_str)
    contracts = prompt_with_default("上期所合约(英文逗号分隔)", "RB2610,AG2612,CU2609")
    accounts_text = prompt_with_default("账户列表(英文逗号分隔)", "ACC_SIM_001,ACC_SIM_002,ACC_SIM_003")
    output = input("输出xlsx路径(留空则自动dist+时间戳): ").strip()

    try:
        rows = int(rows_text)
    except ValueError as exc:
        raise ValueError(f"rows 必须是整数，当前输入: {rows_text}") from exc
    accounts = [x.strip() for x in accounts_text.split(",") if x.strip()]
    if not accounts:
        raise ValueError("accounts 不能为空")

    return UserInputs(
        rows=rows,
        start_date=start_date,
        end_date=end_date,
        contracts=contracts,
        accounts=accounts,
        output=output,
    )


def parse_contracts(contracts_text: str) -> list[SymbolConfig]:
    raw_items = [x.strip() for x in contracts_text.split(",") if x.strip()]
    if not raw_items:
        raise ValueError("contracts 不能为空")
    unique_contracts = list(dict.fromkeys(raw_items))
    symbols: list[SymbolConfig] = []
    unsupported: list[str] = []
    pattern = re.compile(r"^([A-Za-z]{1,2})(\d{4})$")
    for code in unique_contracts:
        m = pattern.match(code)
        if not m:
            unsupported.append(code)
            continue
        product = m.group(1).lower()
        normalized_code = f"{m.group(1).upper()}{m.group(2)}"
        profile = SHFE_PRODUCT_PROFILES.get(product)
        if profile is None:
            unsupported.append(code)
            continue
        symbols.append(SymbolConfig(code=normalized_code, profile=profile))
    if unsupported:
        supported_products = ",".join(sorted(SHFE_PRODUCT_PROFILES.keys()))
        raise ValueError(
            f"存在非上期所或格式不正确合约: {unsupported}。"
            f"支持品种前缀: {supported_products}"
        )
    return symbols


def daterange(d1: date, d2: date) -> list[date]:
    days = (d2 - d1).days
    return [d1 + timedelta(days=i) for i in range(days + 1)]


def is_trading_day(d: date) -> bool:
    return d.weekday() < 5


def round_to_tick(price: float, tick: float) -> float:
    return round(round(price / tick) * tick, 2)


def random_trade_times(trade_date: date, count: int) -> list[time]:
    slots: list[time] = []
    for _ in range(count):
        start, end = random.choice(TRADING_SESSIONS)
        start_dt = datetime.combine(trade_date, start)
        end_dt = datetime.combine(trade_date, end)
        seconds = int((end_dt - start_dt).total_seconds())
        slots.append((start_dt + timedelta(seconds=random.randint(0, seconds))).time())
    slots.sort()
    return slots


def format_date(d: date) -> str:
    return d.strftime("%Y-%m-%d")


def format_time(t: time) -> str:
    return t.strftime("%H:%M:%S")


def build_prev_settlement_map(
    symbols: list[SymbolConfig], trade_days: list[date]
) -> dict[tuple[str, date], float]:
    prev_settle_map: dict[tuple[str, date], float] = {}
    for symbol in symbols:
        last_settle = symbol.profile.base_price * (1 + random.uniform(-0.01, 0.01))
        for d in trade_days:
            prev_settle_map[(symbol.code, d)] = round(last_settle, 2)
            change = random.gauss(0.0, 0.008)  # 日波动约 0.8%
            change = max(-0.04, min(0.04, change))
            today_settle = last_settle * (1 + change)
            last_settle = max(symbol.profile.tick_size, today_settle)
    return prev_settle_map


def generate_records(
    rows: int,
    start_date: date,
    end_date: date,
    accounts: list[str],
    symbols: list[SymbolConfig],
) -> list[dict[str, str]]:
    trade_days = [d for d in daterange(start_date, end_date) if is_trading_day(d)]
    if not trade_days:
        raise ValueError("日期范围内没有交易日，请调整 start-date / end-date")

    prev_settlement = build_prev_settlement_map(symbols, trade_days)

    positions: dict[tuple[str, str, str], int] = {}
    result: list[dict[str, str]] = []
    seq = 100000

    sampled_dates = [random.choice(trade_days) for _ in range(rows)]
    sampled_dates.sort()
    day_counts: dict[date, int] = {}
    for d in sampled_dates:
        day_counts[d] = day_counts.get(d, 0) + 1
    time_buckets = {d: random_trade_times(d, cnt) for d, cnt in day_counts.items()}

    for trade_date in sampled_dates:
        settle_date = trade_date
        account = random.choice(accounts)
        symbol = random.choice(symbols)
        profile = symbol.profile

        long_key = (account, symbol.code, "多")
        short_key = (account, symbol.code, "空")
        long_pos = positions.get(long_key, 0)
        short_pos = positions.get(short_key, 0)

        can_close = long_pos > 0 or short_pos > 0
        if can_close and random.random() < 0.45:
            open_close = random.choice(["平今", "平仓"])
            if long_pos > 0 and short_pos > 0:
                close_long = random.random() < 0.5
            else:
                close_long = long_pos > 0
            if close_long:
                side = "卖"
                qty = random.randint(1, long_pos)
                positions[long_key] = long_pos - qty
            else:
                side = "买"
                qty = random.randint(1, short_pos)
                positions[short_key] = short_pos - qty
        else:
            open_close = "开仓"
            side = random.choice(["买", "卖"])
            qty = random.randint(profile.min_qty, profile.max_qty)
            if side == "买":
                positions[long_key] = long_pos + qty
            else:
                positions[short_key] = short_pos + qty

        trade_tm = time_buckets[trade_date].pop(0)

        prev_settle = prev_settlement[(symbol.code, trade_date)]
        move = random.gauss(0.0, 0.0025)  # 围绕前一日结算价的日内偏移
        move = max(-0.015, min(0.015, move))
        raw_price = prev_settle * (1 + move)
        price = round_to_tick(raw_price, profile.tick_size)
        fee = round(price * qty * profile.fee_rate, 2)

        seq += 1
        trade_no = f"TRD{format_date(trade_date).replace('-', '')}{seq:08d}"
        order_no = f"ORD{format_date(trade_date).replace('-', '')}{seq + 900000:08d}"

        result.append(
            {
                "结算日期": format_date(settle_date),
                "成交编号": trade_no,
                "成交日期": format_date(trade_date),
                "成交时间": format_time(trade_tm),
                "操作账户": account,
                "合约代码": symbol.code,
                "买卖方向": side,
                "开平标志": open_close,
                "成交手数": str(qty),
                "成交价格": f"{price:.2f}",
                "成交手续费": f"{fee:.2f}",
                "持仓类型": "投机",
                "委托编号": order_no,
            }
        )

    return result


def validate_open_close_constraints(records: list[dict[str, str]]) -> None:
    """校验每个账户+合约+多空方向，累计开仓手数必须 >= 累计平仓手数。"""
    # key: (account, symbol, side_type['多'|'空']) -> net_open_qty
    net_pos: dict[tuple[str, str, str], int] = {}
    for idx, row in enumerate(records, start=2):
        account = row["操作账户"]
        symbol = row["合约代码"]
        side = row["买卖方向"]
        open_close = row["开平标志"]
        qty = int(row["成交手数"])

        if open_close == "开仓":
            key = (account, symbol, "多" if side == "买" else "空")
            net_pos[key] = net_pos.get(key, 0) + qty
            continue

        # 平仓/平今
        key = (account, symbol, "多" if side == "卖" else "空")
        net_pos[key] = net_pos.get(key, 0) - qty
        if net_pos[key] < 0:
            raise ValueError(
                "平仓数量超过可平仓持仓: "
                f"第{idx}行, 账户={account}, 合约={symbol}, 方向={side}, 开平={open_close}, 手数={qty}"
            )


def col_num_to_name(col_num: int) -> str:
    result = ""
    n = col_num
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _cell_ref(col_index: int, row_index: int) -> str:
    return f"{col_num_to_name(col_index)}{row_index}"


def _extract_header_from_row(row_elem: ET.Element) -> list[str]:
    values: list[str] = []
    for cell in row_elem.findall(f"{{{NS}}}c"):
        inline = cell.find(f"{{{NS}}}is")
        if inline is None:
            values.append("")
            continue
        text_node = inline.find(f".//{{{NS}}}t")
        values.append(text_node.text if text_node is not None and text_node.text is not None else "")
    return values


def _build_data_cell(
    row_elem: ET.Element,
    col_index: int,
    row_index: int,
    value: str,
    style_id: str | None,
    numeric: bool,
) -> None:
    attrs = {"r": _cell_ref(col_index, row_index)}
    if style_id is not None:
        attrs["s"] = style_id
    if numeric:
        c = ET.SubElement(row_elem, f"{{{NS}}}c", attrs)
        v = ET.SubElement(c, f"{{{NS}}}v")
        v.text = value
        return
    attrs["t"] = "inlineStr"
    c = ET.SubElement(row_elem, f"{{{NS}}}c", attrs)
    inline = ET.SubElement(c, f"{{{NS}}}is")
    t = ET.SubElement(inline, f"{{{NS}}}t")
    t.text = value


def write_xlsx_from_template(
    records: list[dict[str, str]],
    template_path: Path,
    output_path: Path,
) -> None:
    if not template_path.exists():
        raise FileNotFoundError(f"模板不存在: {template_path}")

    with zipfile.ZipFile(template_path, "r") as zin:
        sheet_xml = zin.read("xl/worksheets/sheet1.xml")

    root = ET.fromstring(sheet_xml)
    dimension = root.find(f"{{{NS}}}dimension")
    sheet_data = root.find(f"{{{NS}}}sheetData")
    if sheet_data is None:
        raise ValueError("模板 sheet1 缺少 sheetData")

    rows = sheet_data.findall(f"{{{NS}}}row")
    if not rows:
        raise ValueError("模板 sheet1 缺少表头行")
    header_row = rows[0]
    actual_headers = _extract_header_from_row(header_row)
    if actual_headers != HEADERS:
        raise ValueError(f"模板表头不匹配，实际: {actual_headers}")

    for old_row in rows[1:]:
        sheet_data.remove(old_row)

    col_style_map: dict[int, str] = {}
    cols_elem = root.find(f"{{{NS}}}cols")
    if cols_elem is not None:
        for col in cols_elem.findall(f"{{{NS}}}col"):
            min_idx = int(col.attrib.get("min", "0"))
            max_idx = int(col.attrib.get("max", "0"))
            style = col.attrib.get("style")
            if style is None:
                continue
            for idx in range(min_idx, max_idx + 1):
                col_style_map[idx] = style

    numeric_cols = {"成交手数", "成交价格", "成交手续费"}
    for row_no, record in enumerate(records, start=2):
        row_elem = ET.SubElement(sheet_data, f"{{{NS}}}row", {"r": str(row_no)})
        for col_idx, header in enumerate(HEADERS, start=1):
            value = record[header]
            _build_data_cell(
                row_elem=row_elem,
                col_index=col_idx,
                row_index=row_no,
                value=value,
                style_id=col_style_map.get(col_idx),
                numeric=header in numeric_cols,
            )

    last_row = max(1, len(records) + 1)
    if dimension is not None:
        dimension.attrib["ref"] = f"A1:M{last_row}"

    ET.register_namespace("", NS)
    new_sheet_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(template_path, "r") as zin, zipfile.ZipFile(output_path, "w") as zout:
        for item in zin.infolist():
            data = new_sheet_xml if item.filename == "xl/worksheets/sheet1.xml" else zin.read(item.filename)
            zout.writestr(item, data)


def build_output_path(custom_output: str) -> Path:
    if custom_output:
        out = Path(custom_output)
        if out.suffix.lower() != ".xlsx":
            out = out.with_suffix(".xlsx")
        return out if out.is_absolute() else (Path.cwd() / out).resolve()
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return PROJECT_ROOT / "dist" / f"期货成交记录模拟数据_{ts}.xlsx"


def resolve_existing_path(path_text: str, must_exist: bool = True) -> Path:
    raw = Path(path_text)
    candidates: list[Path] = []
    if raw.is_absolute():
        candidates.append(raw)
    else:
        candidates.append((Path.cwd() / raw).resolve())
        candidates.append((PROJECT_ROOT / raw).resolve())

    for p in candidates:
        if p.exists():
            return p
    if must_exist:
        raise FileNotFoundError(f"文件不存在: {path_text}。已尝试: {[str(x) for x in candidates]}")
    return candidates[0]


def main() -> None:
    user_inputs = read_user_inputs()
    seed = secrets.randbits(32)
    random.seed(seed)

    start = datetime.strptime(user_inputs.start_date, "%Y-%m-%d").date()
    end = datetime.strptime(user_inputs.end_date, "%Y-%m-%d").date()
    if start > end:
        raise ValueError("start-date 不能晚于 end-date")
    if user_inputs.rows <= 0:
        raise ValueError("rows 必须大于 0")

    symbols = parse_contracts(user_inputs.contracts)
    records = generate_records(
        rows=user_inputs.rows,
        start_date=start,
        end_date=end,
        accounts=user_inputs.accounts,
        symbols=symbols,
    )
    validate_open_close_constraints(records)

    template_path = resolve_existing_path("template/期货成交记录导入(明细).xlsx", must_exist=True)
    output_path = build_output_path(user_inputs.output)
    write_xlsx_from_template(
        records=records,
        template_path=template_path,
        output_path=output_path,
    )
    print(f"已生成 {len(records)} 条记录: {output_path}")
    print(f"本次随机种子: {seed}")


if __name__ == "__main__":
    main()
