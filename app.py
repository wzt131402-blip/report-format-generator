#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
报表格式生成器 —— Streamlit 网页版
===================================

这是一个小工具网页：选一种报表 → 填几项基本信息 → 下载一个「排版已完成、
内容为空」的 Excel 文件。之后你只需要往格子里填数据，不用再调格式。

启动方式（在本文件所在目录打开终端）：
    py -m streamlit run app.py

启动后浏览器会自动打开 http://localhost:8501
如果没自动打开，手动把这个地址复制到浏览器即可。

文件分工：
    excel_engine.py   排版引擎，负责把定义变成带格式的 Excel
    report_specs.py   报表定义，想加新报表改这里
    app.py            本文件，只负责网页界面，不含排版逻辑
"""

from __future__ import annotations

import os
import re
from dataclasses import replace
from io import BytesIO

import streamlit as st
from openpyxl import load_workbook

import excel_engine as eng
import report_specs as rs


# ---------------------------------------------------------------------------
# 小工具函数
# ---------------------------------------------------------------------------

# Windows / macOS 的文件名里不允许出现这些字符
_ILLEGAL_CHARS = re.compile(r'[\\/:*?"<>|\r\n\t]+')

# xlsx 的标准 MIME 类型，下载按钮要用
XLSX_MIME = ("application/vnd.openxmlformats-officedocument."
             "spreadsheetml.sheet")


def safe_filename(*parts: str, ext: str = ".xlsx") -> str:
    """
    把几段文字拼成一个合法的文件名。

    例如 safe_filename("财务状况表", "某某公司", "2024 年度")
      -> "财务状况表_某某公司_2024 年度.xlsx"
    自动去掉文件名里的非法字符，避免下载失败。
    """
    joined = "_".join(p.strip() for p in parts if p and p.strip())
    cleaned = _ILLEGAL_CHARS.sub("", joined).strip(" ._")
    return (cleaned or "报表") + ext


def sheet_to_grid(wb, sheet_name: str, max_rows: int = 80) -> list[list]:
    """
    把工作表内容读成二维列表，供 st.dataframe 在网页上预览。
    这样下载之前就能先看到表格大概长什么样。

    注意：合计单元格里存的是公式文本（如 =SUM(C6:C13)），
    预览时看到的是公式而不是计算结果——用 Excel 打开后会自动算出数字。
    """
    ws = wb[sheet_name]
    n_cols = ws.max_column
    grid = []
    for row in ws.iter_rows(min_row=1,
                            max_row=min(ws.max_row, max_rows),
                            values_only=True):
        # 每行补齐到相同列数，否则 Streamlit 会因为行长不一致而显示错乱
        values = list(row) + [None] * (n_cols - len(row))
        # 全部转成文本：同一列里既有标题文字又有数字时，
        # Streamlit 底层的 pyarrow 会因为类型混杂而报错
        grid.append(["" if v is None else str(v) for v in values[:n_cols]])
    return grid


def column_table(spec: rs.ReportSpec) -> list[dict]:
    """把列定义整理成表格数据，显示在「结构说明」标签页里。"""
    return [
        {
            "列号": i,
            "表头": c.header,
            "列宽": c.width,
            "数字格式": c.number_format,
            "参与合计": "是" if c.summable else "—",
            "下拉选项": "、".join(c.choices) if c.choices else "—",
        }
        for i, c in enumerate(spec.columns, start=1)
    ]


def row_table(spec: rs.ReportSpec) -> list[dict]:
    """把固定行报表的科目结构整理成表格数据。"""
    kind_name = {"item": "可填科目", "section": "分类标题", "total": "自动合计"}
    out = []
    seq = 0
    for r in spec.rows:
        if r.kind == "item":
            seq += 1
        out.append({
            "行次": seq if r.kind == "item" else "—",
            "项目": "　" * r.indent + r.label,
            "类型": kind_name.get(r.kind, r.kind),
        })
    return out


# ---------------------------------------------------------------------------
# 访问密码
# ---------------------------------------------------------------------------
#
# 密码不写在代码里！原因：部署到 Streamlit 官方云要求仓库是公开的，
# 代码谁都能看到，密码写死在代码里等于没有密码。
#
# 密码的存放位置（二选一）：
#   1. Streamlit 官方云：网站后台 App settings -> Secrets 里填
#        access_password = "你的密码"
#   2. 本机 / 自己的服务器：在项目下建 .streamlit/secrets.toml，内容同上
#      （该文件已被 .gitignore 排除，不会被提交）
#      或者设置环境变量 REPORT_APP_PASSWORD
#
# 两处都没配置时，视为「不加限制」，方便本地开发。

# 环境变量名，作为 secrets.toml 的备用方案
PASSWORD_ENV_NAME = "REPORT_APP_PASSWORD"


def expected_password() -> str:
    """读取配置的访问密码；没配置则返回空字符串。"""
    value = None
    try:
        # st.secrets 在没有配置文件时访问会抛异常，所以整体包一层
        if "access_password" in st.secrets:
            value = st.secrets["access_password"]
    except Exception:
        value = None
    if value:
        return str(value)
    return os.environ.get(PASSWORD_ENV_NAME, "") or ""


def require_login() -> bool:
    """
    访问密码闸门。返回 True 表示放行，可以继续渲染工具主体；
    返回 False 表示停在登录页，调用方应当立刻 st.stop()。
    """
    password = expected_password()

    if not password:
        # 没配密码：本机用无所谓，但放到公网上就是裸奔，给个醒目提醒
        st.info(
            "当前**未设置访问密码**，任何打开此页面的人都能直接使用。\n\n"
            "如果只是本机自用可以忽略；若要部署到公网，请在 Streamlit 后台的 "
            "Secrets 里配置 `access_password`，或设置环境变量 "
            f"`{PASSWORD_ENV_NAME}`。"
        )
        return True

    if st.session_state.get("logged_in"):
        return True

    # ---- 未登录：只显示登录框 ----
    st.title("报表格式生成器")
    st.caption("本工具已启用访问密码，输入正确密码后即可使用。")
    entered = st.text_input("访问密码", type="password",
                            autocomplete="current-password")
    if st.button("进入", type="primary"):
        if entered == password:
            st.session_state["logged_in"] = True
            st.rerun()          # 重新跑一遍脚本，这次会走到工具主体
        else:
            st.error("密码不正确，请重试。")
    return False


# ---------------------------------------------------------------------------
# 页面配置与标题
# ---------------------------------------------------------------------------

st.set_page_config(page_title="报表格式生成器", layout="wide")

# 访问密码闸门：没通过就停在这里，后面的工具代码一行都不会执行
if not require_login():
    st.stop()

st.title("报表格式生成器")
st.caption("选一种报表 → 填几项基本信息 → 下载排好版的空白 Excel，"
           "之后只需往里填数据，不用再调格式。")


# ---------------------------------------------------------------------------
# 侧边栏：所有输入项
# ---------------------------------------------------------------------------

specs = rs.all_specs()
name_to_spec = {s.name: s for s in specs}

with st.sidebar:
    st.header("① 选择报表类型")
    picked = st.selectbox(
        "报表",
        list(name_to_spec.keys()),
        index=0,
        help="需要新的报表种类？在 report_specs.py 里照着现有写法加一个即可，"
             "网页上会自动出现。",
    )
    spec = name_to_spec[picked]
    st.info(spec.description)

    st.divider()
    st.header("② 表头信息")
    title = st.text_input("表格大标题", value=spec.default_title,
                          help="显示在第一行正中间")
    org = st.text_input("编制单位", value="", placeholder="例如：某某有限公司")
    period = st.text_input("所属期间", value="",
                           placeholder="例如：2024 年度 / 2024 年 3 月")
    unit_label = st.text_input("右上角单位说明", value=spec.unit_label,
                               placeholder="例如：单位：元")

    st.divider()
    st.header("③ 生成选项")

    if spec.is_fixed_rows:
        st.caption(f"本报表行结构固定，共 {len(spec.rows)} 行科目，无需设置行数。")
        detail_rows = len(spec.rows)
    else:
        detail_rows = st.number_input(
            "预留明细行数",
            min_value=1, max_value=5000, value=30, step=10,
            help="预先铺多少行带格式的空行。不够用时可以在 Excel 里插入新行，"
                 "格式会自动延续。",
        )

    with_total = st.checkbox(
        "末尾加「合计」行（自动求和）",
        value=spec.has_total_row,
        disabled=not spec.has_total_row,
        help="" if spec.has_total_row else "这种报表求合计没有意义，已固定关闭",
    )
    with_remark = st.checkbox("末尾加「备注：」行", value=spec.remark_line)

    st.divider()
    st.header("④ 用自己的空白 Excel（可选）")
    uploaded = st.file_uploader(
        "上传一个 .xlsx 文件",
        type=["xlsx"],
        help="不上传也可以，程序会自动新建一个工作簿。"
             "上传的话，会在你这个文件的基础上重排格式——"
             "你硬盘上的原文件不会被修改，生成的是可下载的新文件。",
    )

    # 设了密码才显示退出按钮；共用电脑时用完可以手动退出
    if expected_password() and st.session_state.get("logged_in"):
        st.divider()
        if st.button("退出登录"):
            st.session_state["logged_in"] = False
            st.rerun()


# 允许用户在界面上改「单位说明」，用 dataclasses.replace 生成一份改过的定义，
# 不去动 report_specs.py 里的原始对象
spec = replace(spec, unit_label=unit_label)


# ---------------------------------------------------------------------------
# 处理上传的文件
# ---------------------------------------------------------------------------

base_bytes = None
if uploaded is not None:
    try:
        raw = uploaded.getvalue()
        # 先试读一次，格式不对就马上告诉用户，别等到生成时才失败
        load_workbook(BytesIO(raw))
        base_bytes = raw
        st.sidebar.success(f"已读取：{uploaded.name}")
    except Exception as exc:
        st.sidebar.error(
            f"「{uploaded.name}」打不开，已改为新建空白工作簿。\n\n"
            f"原因：{type(exc).__name__}\n\n"
            "请确认上传的是 .xlsx 格式（Excel 2007 及以后）。"
            "老的 .xls 需要先在 Excel 里「另存为 .xlsx」。"
        )


# ---------------------------------------------------------------------------
# 生成工作簿
# ---------------------------------------------------------------------------

try:
    wb = eng.build_workbook(
        spec,
        base_bytes=base_bytes,
        title=title,
        org=org,
        period=period,
        detail_rows=int(detail_rows),
        with_total=with_total,
        with_remark=with_remark,
    )
    data = eng.workbook_to_bytes(wb)
except Exception as exc:
    st.error("生成失败，请把下面的信息反馈给维护者：")
    st.exception(exc)
    st.stop()


# ---------------------------------------------------------------------------
# 主区域：下载按钮 + 三个标签页
# ---------------------------------------------------------------------------

filename = safe_filename(spec.name, org, period)

left, right = st.columns([1, 2])
with left:
    st.download_button(
        label=f"下载 {filename}",
        data=data,
        file_name=filename,
        mime=XLSX_MIME,
        type="primary",
    )
with right:
    st.caption(
        f"工作表：{spec.sheet_name}　|　"
        f"{wb[spec.sheet_name].max_row} 行 × {len(spec.columns)} 列　|　"
        f"文件大小 {len(data) / 1024:.1f} KB　|　"
        f"打印方向：{'横向' if spec.landscape else '纵向'}"
    )

st.divider()

tab_preview, tab_struct, tab_help = st.tabs(["格式预览", "结构说明", "填表说明"])

with tab_preview:
    st.caption("下面是生成结果的预览。合计格显示的是公式文本，"
               "用 Excel 打开后会自动算出数字。")
    grid = sheet_to_grid(wb, spec.sheet_name)
    if grid:
        # 用报表真实的表头当列名，比默认的 0、1、2、3 好认
        # 字典的键不能重复，万一有两个同名表头就补上序号区分
        seen: dict[str, int] = {}
        keys = []
        for c in spec.columns:
            seen[c.header] = seen.get(c.header, 0) + 1
            keys.append(c.header if seen[c.header] == 1
                        else f"{c.header}({seen[c.header]})")
        records = [dict(zip(keys, row)) for row in grid]
        st.dataframe(records, hide_index=True,
                     height=min(600, 35 * len(records) + 40))

with tab_struct:
    st.subheader("列定义")
    st.dataframe(column_table(spec), hide_index=True)

    st.subheader("行结构")
    if spec.is_fixed_rows:
        st.dataframe(row_table(spec), hide_index=True,
                     height=min(600, 35 * len(spec.rows) + 40))
    else:
        st.write(f"明细流水型报表，预留 **{detail_rows}** 行空白行供填写"
                 + ("，末尾自动带一行「合计」。" if with_total else "。"))

    with st.expander("纯文字版结构概要"):
        st.code(rs.spec_summary(spec), language=None)

with tab_help:
    st.markdown(f"""
### 拿到文件之后怎么用

1. **直接填数据**。带浅色边框的空白格就是留给你的，表头、科目名、
   列宽、配色、打印设置都已经排好，不用再调格式。

2. **合计不用自己算**。
   - 明细型报表（收支、工资、盘点）末尾的「合计」行已经写好 `=SUM()` 公式，
     数字一填进去就自动出结果。
   - 财务状况表的「流动资产合计」「资产总计」等行也都内置了公式，
     其中「固定资产账面价值」会自动等于「固定资产原价 − 累计折旧」。

3. **明细行不够用怎么办**。
   在**合计行的上一行**点右键 →「插入」→「整行」，这样新行会落在
   `SUM` 的区间里面，合计范围会自动扩大。
   ⚠️ 不要在合计行**下面**插行，那样新行不会被计入合计。
   插入后新行会自动继承上一行的边框和数字格式。

4. **带下拉箭头的列**（如结算方式、盘点结果、项目状态），
   点单元格右侧的小箭头选就行，手打错别字会被 Excel 拦下来。

5. **打印已经设置好**：
   {('横向' if spec.landscape else '纵向')} A4、自动缩放为一页宽、
   跨页时每页重复表头。直接 Ctrl+P 即可。

---

### 几点说明

- **你上传的原文件不会被改动**。程序只在内存里操作副本，
  下载得到的是一个全新的文件。
- **不上传文件也完全可以用**，程序会自动新建一个工作簿。
- **想增加新的报表种类**：打开 `report_specs.py`，
  照着里面任意一个 `ReportSpec` 的样子复制一份、改改列和行，
  然后加进文件末尾的 `REPORTS` 元组，刷新网页就能看到了，
  不需要动 `excel_engine.py` 和 `app.py`。
- **想换配色或字体**：改 `excel_engine.py` 顶部那几个
  `COLOR_` 开头的常量和 `FONT_NAME` 即可。
""")
