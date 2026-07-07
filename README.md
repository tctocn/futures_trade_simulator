# Futures Trade Simulator

用于生成期货相关 Excel 导入文件的本地脚本集合。脚本基于项目 `template/` 目录中的 Excel 模板输出结果，生成文件默认写入 `dist/` 目录。

## 环境要求

- Python 3.9+
- 推荐使用虚拟环境：`/Users/tongchen/Workspaces/PyCharm/.venv`
- 依赖库：
  - `akshare`
  - `openpyxl`
  - `pandas`

安装依赖：

```bash
/Users/tongchen/Workspaces/PyCharm/.venv/bin/pip install akshare openpyxl pandas
```

## 脚本说明

### 1. 生成期货成交记录

模板：

```text
template/期货成交记录导入(明细).xlsx
```

运行：

```bash
/Users/tongchen/Workspaces/PyCharm/.venv/bin/python scripts/generate_futures_trades.py
```

功能：

- 交互式输入合约、日期范围、记录行数等参数
- 按模板表头生成期货成交明细
- 支持开仓、平仓、平今逻辑
- 输出到 `dist/`

### 2. 生成期货合约导入文件

模板：

```text
template/期货合约导入.xlsx
```

运行：

```bash
/Users/tongchen/Workspaces/PyCharm/.venv/bin/python scripts/generate_futures_contracts_akshare.py
```

功能：

- 交互式输入多个合约，英文逗号分隔
- 通过 AKShare 拉取合约基础信息
- 支持部分 AKShare 手续费接口缺失的合约，使用合约详情接口兜底
- 输出到 `dist/`

示例输入：

```text
NI2602,NI2603,RB2603,RB2605,PG2602,PG2603,PG2605
```

### 3. 生成期货合约价格导入文件

模板：

```text
template/期货合约价导入.xlsx
```

运行：

```bash
/Users/tongchen/Workspaces/PyCharm/.venv/bin/python scripts/generate_futures_contract_prices_akshare.py
```

功能：

- 交互式输入合约和时间范围
- 时间范围默认从当天往前一年到当天
- 通过 AKShare 拉取历史结算价、收盘价
- 输出到 `dist/`

### 4. 生成现货市场价导入文件

模板：

```text
template/现货市场价导入.xlsx
```

运行：

```bash
/Users/tongchen/Workspaces/PyCharm/.venv/bin/python scripts/generate_spot_market_price_from_futures_akshare.py
```

功能：

- 交互式输入合约、品名、材质、规格、产地、省、市
- 时间范围固定为当天往前一年到当天
- 周末和节假日没有合约价格时，向前取最近一个可用结算价
- 公允价值计算：

```text
公允价值 = 合约结算价 + 升贴水
```

升贴水规则：

- 回车不输入：每天按合约结算价的 `[-1%, 1%]` 随机波动
- 输入 `1%`：每天按合约结算价的 `[-1%, 1%]` 随机波动
- 输入 `100`：每天按 `[-100, 100]` 固定金额区间随机波动

## 目录结构

```text
.
├── scripts/      # 生成脚本
├── template/     # Excel 导入模板
├── dist/         # 生成结果，已在 .gitignore 中忽略
└── output/       # 历史输出目录，已在 .gitignore 中忽略
```

## Git 提交建议

建议提交：

- `README.md`
- `.gitignore`
- `scripts/`
- `template/`

不建议提交：

- `dist/`
- `output/`
- `.idea/`
- `.DS_Store`
- 虚拟环境目录
