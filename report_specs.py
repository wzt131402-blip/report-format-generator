#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
报表格式定义
=============

这个文件是「纯数据」，不含任何处理逻辑。每一种报表就是一个 ReportSpec 对象，
描述了它有哪些列、哪些固定行、合计怎么算、要不要下拉框。

想新增一种报表，照抄下面的写法加一个 ReportSpec，然后放进文件末尾的
REPORTS 元组里就行——excel_engine.py 和 app.py 都不用改，
网页上会自动多出这个选项。

常用数字格式对照：
    MONEY  "#,##0.00"    金额，千分位 + 两位小数，如 1,234.50
    INT    "#,##0"       整数，千分位
    PCT    "0.0%"        百分比，输入 0.85 显示 85.0%
    DATE   "yyyy-mm-dd"  日期
    TEXT   "@"           纯文本，防止「0012」这样的编号丢掉前导零
"""

from __future__ import annotations

from excel_engine import Column, ReportSpec, RowSpec

# --- 数字格式简写，避免下面重复写一长串 ---
MONEY = "#,##0.00"
INT = "#,##0"
PCT = "0.0%"
DATE = "yyyy-mm-dd"
TEXT = "@"


# ---------------------------------------------------------------------------
# 1. 财务状况表（资产负债表）
#    行结构固定：每个会计科目都是一行，合计行内置 SUM 公式
# ---------------------------------------------------------------------------

BALANCE_SHEET = ReportSpec(
    key="balance_sheet",
    name="财务状况表",
    description="资产负债表结构，科目行已排好，各级合计自带求和公式，只需填金额",
    sheet_name="财务状况表",
    default_title="财务状况表",
    unit_label="单位：元",
    landscape=False,          # 四列窄表，纵向打印更合适
    seq_column=2,             # 第 2 列是「行次」，自动编号
    columns=(
        Column("项目", width=32, align="left"),
        Column("行次", width=6, number_format=INT),
        Column("期末余额", width=17, number_format=MONEY,
               align="right", summable=True),
        Column("年初余额", width=17, number_format=MONEY,
               align="right", summable=True),
    ),
    rows=(
        # ===== 资产 =====
        RowSpec("一、资产", kind="section"),
        RowSpec("流动资产：", kind="section", indent=1),
        RowSpec("货币资金", indent=2, ref="cash"),
        RowSpec("短期投资", indent=2),
        RowSpec("应收票据", indent=2),
        RowSpec("应收账款", indent=2),
        RowSpec("预付账款", indent=2),
        RowSpec("其他应收款", indent=2),
        RowSpec("存货", indent=2),
        RowSpec("其他流动资产", indent=2, ref="other_ca"),
        RowSpec("流动资产合计", indent=1, kind="total", ref="ca_total",
                total_of=("cash", "other_ca")),

        RowSpec("非流动资产：", kind="section", indent=1),
        RowSpec("长期股权投资", indent=2, ref="lt_eq"),
        RowSpec("固定资产原价", indent=2, ref="fa_gross"),
        RowSpec("减：累计折旧", indent=2, ref="fa_dep"),
        # 账面价值 = 原价 - 累计折旧，自动算，不用手填
        RowSpec("固定资产账面价值", indent=2, ref="fa_net",
                formula="={col}{fa_gross}-{col}{fa_dep}"),
        RowSpec("在建工程", indent=2, ref="cip"),
        RowSpec("无形资产", indent=2, ref="intangible"),
        RowSpec("其他非流动资产", indent=2, ref="other_nca"),
        # 注意：合计只加「账面价值」，不能把原价和累计折旧也加进去，
        # 否则会重复计算，所以这里用 total_refs 精确点名要相加的行
        RowSpec("非流动资产合计", indent=1, kind="total", ref="nca_total",
                total_refs=("lt_eq", "fa_net", "cip", "intangible", "other_nca")),
        RowSpec("资产总计", kind="total", ref="asset_total",
                total_refs=("ca_total", "nca_total")),

        # ===== 负债 =====
        RowSpec("二、负债", kind="section"),
        RowSpec("流动负债：", kind="section", indent=1),
        RowSpec("短期借款", indent=2, ref="st_loan"),
        RowSpec("应付票据", indent=2),
        RowSpec("应付账款", indent=2),
        RowSpec("预收账款", indent=2),
        RowSpec("应付职工薪酬", indent=2),
        RowSpec("应交税费", indent=2),
        RowSpec("其他应付款", indent=2),
        RowSpec("其他流动负债", indent=2, ref="other_cl"),
        RowSpec("流动负债合计", indent=1, kind="total", ref="cl_total",
                total_of=("st_loan", "other_cl")),

        RowSpec("非流动负债：", kind="section", indent=1),
        RowSpec("长期借款", indent=2, ref="lt_loan"),
        RowSpec("长期应付款", indent=2),
        RowSpec("其他非流动负债", indent=2, ref="other_ncl"),
        RowSpec("非流动负债合计", indent=1, kind="total", ref="ncl_total",
                total_of=("lt_loan", "other_ncl")),
        RowSpec("负债合计", kind="total", ref="liab_total",
                total_refs=("cl_total", "ncl_total")),

        # ===== 净资产 =====
        RowSpec("三、净资产（或所有者权益）", kind="section"),
        RowSpec("实收资本（或股本）", indent=1, ref="paid_in"),
        RowSpec("资本公积", indent=1),
        RowSpec("盈余公积", indent=1),
        RowSpec("未分配利润", indent=1),
        RowSpec("其他净资产", indent=1, ref="other_ne"),
        RowSpec("净资产合计", indent=1, kind="total", ref="ne_total",
                total_of=("paid_in", "other_ne")),
        RowSpec("负债和净资产总计", kind="total", ref="grand_total",
                total_refs=("liab_total", "ne_total")),
    ),
    footer_labels=("制表人：", "审核人：", "单位负责人："),
)


# ---------------------------------------------------------------------------
# 2. 收支明细表（流水账）
#    行数不固定，由用户在网页上选择预留多少行
# ---------------------------------------------------------------------------

INCOME_EXPENSE = ReportSpec(
    key="income_expense",
    name="收支明细表",
    description="日常收支流水账，收入/支出自动合计，结算方式带下拉选项",
    sheet_name="收支明细",
    default_title="收支明细表",
    unit_label="单位：元",
    landscape=True,
    seq_column=1,
    columns=(
        Column("序号", width=6, number_format=INT),
        Column("日期", width=13, number_format=DATE),
        Column("摘要", width=26, align="left"),
        Column("收入金额", width=14, number_format=MONEY,
               align="right", summable=True),
        Column("支出金额", width=14, number_format=MONEY,
               align="right", summable=True),
        Column("结算方式", width=12,
               choices=("现金", "银行转账", "支付宝", "微信", "支票", "其他")),
        Column("票据号", width=15, number_format=TEXT),
        Column("经办人", width=10),
        Column("备注", width=20, align="left"),
    ),
    footer_labels=("制表人：", "审核人：", "出纳："),
)


# ---------------------------------------------------------------------------
# 3. 员工工资表
# ---------------------------------------------------------------------------

PAYROLL = ReportSpec(
    key="payroll",
    name="员工工资表",
    description="应发、代扣、实发各栏已分好，底部自动汇总每一列合计",
    sheet_name="工资表",
    default_title="员工工资表",
    unit_label="单位：元",
    landscape=True,
    seq_column=1,
    columns=(
        Column("序号", width=6, number_format=INT),
        Column("工号", width=10, number_format=TEXT),   # 文本格式保住前导零
        Column("姓名", width=10),
        Column("部门", width=12),
        Column("基本工资", width=11, number_format=MONEY,
               align="right", summable=True),
        Column("岗位工资", width=11, number_format=MONEY,
               align="right", summable=True),
        Column("绩效奖金", width=11, number_format=MONEY,
               align="right", summable=True),
        Column("加班补贴", width=11, number_format=MONEY,
               align="right", summable=True),
        Column("其他补贴", width=11, number_format=MONEY,
               align="right", summable=True),
        Column("应发合计", width=12, number_format=MONEY,
               align="right", summable=True),
        Column("社保(个人)", width=11, number_format=MONEY,
               align="right", summable=True),
        Column("公积金(个人)", width=12, number_format=MONEY,
               align="right", summable=True),
        Column("个人所得税", width=11, number_format=MONEY,
               align="right", summable=True),
        Column("其他扣款", width=11, number_format=MONEY,
               align="right", summable=True),
        Column("实发工资", width=12, number_format=MONEY,
               align="right", summable=True),
        Column("领取签名", width=12),
    ),
    footer_labels=("制表人：", "审核人：", "批准人："),
)


# ---------------------------------------------------------------------------
# 4. 库存盘点表
# ---------------------------------------------------------------------------

INVENTORY = ReportSpec(
    key="inventory",
    name="库存盘点表",
    description="账面数与实盘数对照，盘点结果带下拉选项，数量金额自动合计",
    sheet_name="库存盘点",
    default_title="库存盘点表",
    unit_label="单位：元",
    landscape=True,
    seq_column=1,
    columns=(
        Column("序号", width=6, number_format=INT),
        Column("物料编码", width=15, number_format=TEXT),
        Column("品名", width=20, align="left"),
        Column("规格型号", width=15),
        Column("单位", width=7),
        Column("存放位置", width=13),
        Column("账面数量", width=11, number_format=INT,
               align="right", summable=True),
        Column("实盘数量", width=11, number_format=INT,
               align="right", summable=True),
        Column("差异数量", width=11, number_format=INT,
               align="right", summable=True),
        Column("单价", width=11, number_format=MONEY, align="right"),
        Column("差异金额", width=13, number_format=MONEY,
               align="right", summable=True),
        Column("盘点结果", width=11,
               choices=("相符", "盘盈", "盘亏", "损毁", "待查")),
        Column("原因说明", width=22, align="left"),
    ),
    footer_labels=("盘点人：", "复核人：", "仓库主管："),
)


# ---------------------------------------------------------------------------
# 5. 项目进度表
# ---------------------------------------------------------------------------

PROJECT_TRACKER = ReportSpec(
    key="project_tracker",
    name="项目进度表",
    description="任务清单与进度跟踪，阶段、状态、风险等级均为下拉选择",
    sheet_name="项目进度",
    default_title="项目进度表",
    unit_label="",            # 进度表不涉及金额，右上角留空
    landscape=True,
    seq_column=1,
    columns=(
        Column("序号", width=6, number_format=INT),
        Column("任务名称", width=26, align="left"),
        Column("所属阶段", width=13,
               choices=("启动", "规划", "执行", "监控", "收尾")),
        Column("负责人", width=10),
        Column("协作人", width=14),
        Column("计划开始", width=13, number_format=DATE),
        Column("计划结束", width=13, number_format=DATE),
        Column("实际完成", width=13, number_format=DATE),
        Column("完成度", width=10, number_format=PCT),
        Column("状态", width=11,
               choices=("未开始", "进行中", "已完成", "已延期", "已暂停", "已取消")),
        Column("风险等级", width=10, choices=("高", "中", "低")),
        Column("问题与备注", width=28, align="left"),
    ),
    has_total_row=False,      # 进度表求「合计」没有意义
    stripe=True,
    footer_labels=("填表人：", "项目经理：", "审批人："),
)


# ---------------------------------------------------------------------------
# 报表注册表
# ---------------------------------------------------------------------------

# 网页上的下拉列表就按这个顺序显示；新增报表后记得加到这里
REPORTS: tuple[ReportSpec, ...] = (
    BALANCE_SHEET,
    INCOME_EXPENSE,
    PAYROLL,
    INVENTORY,
    PROJECT_TRACKER,
)


def all_specs() -> list[ReportSpec]:
    """返回全部报表定义。"""
    return list(REPORTS)


def get_spec(key: str) -> ReportSpec | None:
    """按 key 查找报表定义，找不到返回 None。"""
    for spec in REPORTS:
        if spec.key == key:
            return spec
    return None


def spec_summary(spec: ReportSpec) -> str:
    """
    生成一段人类可读的结构说明，用于网页上预览这张表长什么样。
    """
    lines = [f"共 {len(spec.columns)} 列："]
    lines.append("  " + " | ".join(c.header for c in spec.columns))

    if spec.is_fixed_rows:
        n_item = sum(1 for r in spec.rows if r.kind == "item")
        n_sec = sum(1 for r in spec.rows if r.kind == "section")
        n_tot = sum(1 for r in spec.rows if r.kind == "total")
        lines.append(
            f"固定 {len(spec.rows)} 行：可填科目 {n_item} 项、"
            f"分类标题 {n_sec} 行、自动合计 {n_tot} 行"
        )
    else:
        lines.append("明细行数由你在下方自选，末尾自动带一行「合计」"
                     if spec.has_total_row else "明细行数由你在下方自选")

    dropdowns = [c.header for c in spec.columns if c.choices]
    if dropdowns:
        lines.append("带下拉选项的列：" + "、".join(dropdowns))

    sums = [c.header for c in spec.columns if c.summable]
    if sums:
        lines.append("自动求和的列：" + "、".join(sums))

    return "\n".join(lines)
