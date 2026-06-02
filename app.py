#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
供應鏈開發機器人
採購需求表單 + 全球供應商自動搜尋（Serper.dev Google Search API）+ Excel 比較報表
"""

import re
import time
from io import BytesIO
from datetime import date, datetime

import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup

# ── 頁面設定 ─────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="供應鏈開發機器人",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── 側欄：API Key 設定 ────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ 設定")
    try:
        serper_key = st.secrets["SERPER_API_KEY"]
        st.success("API Key 已載入", icon="✅")
    except Exception:
        serper_key = st.text_input(
            "Serper API Key",
            type="password",
            value="70d583b38a2c1eebf6ec3a792f0cd1480643cadd",
            help="前往 serper.dev 取得免費 API Key（每月 2,500 次）",
        )
    st.divider()
    st.caption("© 供應鏈開發機器人 v2.0")

st.title("🔍 供應鏈開發機器人")
st.caption("填寫採購需求，自動搜尋全球供應商並匯出比較報表")

SUPPLIER_COLS = [
    '供應商名稱', '聯絡人', '聯絡方式', '品名', '型號',
    '規格概述', '數量', '單價', '交期', '付款方式', '付款條件', '來源', '連結',
]
EMPTY_SUPPLIER = {c: '' for c in SUPPLIER_COLS}

MARKET_QUERIES = {
    '台灣':    '{kw} 台灣 供應商 製造商 廠商',
    '中國大陸': '{kw} 中國 製造商 供應商 工廠 報價',
    '日本':    '{kw} Japan manufacturer supplier price',
    '韓國':    '{kw} Korea manufacturer supplier',
    '美國':    '{kw} USA manufacturer supplier wholesale price',
    '歐洲':    '{kw} Europe manufacturer supplier',
    '東南亞':  '{kw} Southeast Asia manufacturer supplier',
    '全球':    '{kw} global supplier manufacturer wholesale',
}

MARKET_LOCALE = {
    '台灣': ('zh-tw', 'tw'), '中國大陸': ('zh-cn', 'cn'),
    '日本': ('ja',    'jp'), '韓國':    ('ko',    'kr'),
    '美國': ('en',    'us'), '歐洲':    ('en',    'gb'),
    '東南亞': ('en',  'sg'), '全球':    ('en',    'us'),
}

def clean(text: str) -> str:
    return re.sub(r'\s+', ' ', text).strip() if text else ''

# ── Serper 搜尋 ───────────────────────────────────────────────────────────────

def search_serper(keyword: str, market: str, api_key: str) -> list:
    if not api_key:
        st.warning("請在左側側欄輸入 Serper API Key")
        return []

    query = MARKET_QUERIES.get(market, '{kw} supplier').replace('{kw}', keyword)
    hl, gl = MARKET_LOCALE.get(market, ('en', 'us'))

    try:
        r = requests.post(
            'https://google.serper.dev/search',
            headers={'X-API-KEY': api_key, 'Content-Type': 'application/json'},
            json={'q': query, 'num': 10, 'hl': hl, 'gl': gl},
            timeout=20,
        )
        r.raise_for_status()
        data = r.json()
    except Exception as e:
        st.warning(f"{market} 搜尋失敗：{e}")
        return []

    results = []

    # 一般搜尋結果
    for item in data.get('organic', []):
        title   = clean(item.get('title', ''))
        snippet = clean(item.get('snippet', ''))
        link    = item.get('link', '')
        company = re.split(r'[|\-–—]', title)[0].strip()
        if not company:
            continue
        s = dict(EMPTY_SUPPLIER)
        s.update({
            '供應商名稱': company,
            '品名':      keyword,
            '規格概述':  snippet[:150],
            '來源':      f'{market} Google 搜尋',
            '連結':      link,
        })
        results.append(s)

    # 購物結果（如有）
    for item in data.get('shopping', []):
        title  = clean(item.get('title', ''))
        price  = clean(item.get('price', ''))
        source = clean(item.get('source', ''))
        link   = item.get('link', '')
        if not title:
            continue
        s = dict(EMPTY_SUPPLIER)
        s.update({
            '供應商名稱': source,
            '品名':      title,
            '單價':      price,
            '規格概述':  clean(item.get('snippet', ''))[:100],
            '來源':      f'{market} Google 購物',
            '連結':      link,
        })
        results.append(s)

    return results


def search_all_markets(keyword: str, markets: list, api_key: str) -> list:
    all_results = []
    progress = st.progress(0)
    status   = st.empty()

    for i, market in enumerate(markets):
        status.text(f'正在搜尋 {market} 供應商…')
        results = search_serper(keyword, market, api_key)
        all_results.extend(results)
        progress.progress((i + 1) / len(markets))
        time.sleep(0.5)

    status.empty()
    progress.empty()
    return all_results


# ── Excel 匯出 ────────────────────────────────────────────────────────────────

def to_excel(form_data: dict, suppliers_df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        wb      = writer.book
        hdr_fmt = wb.add_format({'bold': True, 'bg_color': '#1F4E79',
                                  'font_color': 'white', 'border': 1, 'align': 'center'})
        lbl_fmt = wb.add_format({'bold': True, 'bg_color': '#D6E4F0', 'border': 1})
        cell_fmt= wb.add_format({'border': 1, 'text_wrap': True})

        # 採購需求工作表
        ws1 = wb.add_worksheet('採購需求')
        ws1.set_column(0, 0, 18)
        ws1.set_column(1, 1, 50)
        for row, (k, v) in enumerate(form_data.items()):
            ws1.write(row, 0, k, lbl_fmt)
            ws1.write(row, 1, str(v), cell_fmt)

        # 供應商比較工作表
        ws2 = wb.add_worksheet('供應商比較')
        col_widths = [25, 12, 20, 22, 14, 35, 10, 12, 10, 12, 12, 16, 40]
        for col, (header, w) in enumerate(zip(suppliers_df.columns, col_widths)):
            ws2.write(0, col, header, hdr_fmt)
            ws2.set_column(col, col, w)
        ws2.freeze_panes(1, 0)
        for row, record in enumerate(suppliers_df.itertuples(index=False), 1):
            for col, val in enumerate(record):
                ws2.write(row, col, str(val) if val else '', cell_fmt)

    return output.getvalue()


# ── 主介面 ────────────────────────────────────────────────────────────────────
tab1, tab2 = st.tabs(["📋 採購需求表單", "📊 供應商搜尋結果 & 匯出"])

# ══════════════════════════════ 表單 ═════════════════════════════════════════
with tab1:
    st.subheader("採購需求表單")
    st.info("填寫完成後點擊「開始搜尋」，系統依目標市場自動搜尋全球供應商。", icon="ℹ️")

    with st.form("procurement_form"):
        col1, col2 = st.columns(2)

        with col1:
            applicant    = st.text_input("申購人 / 部門 ＊", placeholder="例：採購部 / 王小明")
            category     = st.selectbox("產品類別 ＊", [
                "請選擇…", "電子零件", "機械設備", "原材料", "化學材料",
                "包裝材料", "辦公用品", "資訊設備", "空調 / 除濕設備", "其他",
            ])
            product_spec = st.text_area("品名規格 ＊",
                placeholder="例：工業用除濕機，除濕量 50L/day，220V", height=110)
            quantity     = st.text_input("數量 ＊", placeholder="例：10 台")
            purpose      = st.text_area("用途", placeholder="例：廠房防潮", height=80)

        with col2:
            budget            = st.text_input("預算", placeholder="例：NT$50,000 / 台以內")
            target_market     = st.multiselect(
                "目標市場", list(MARKET_QUERIES.keys()),
                default=['台灣', '中國大陸', '全球'],
            )
            existing_supplier = st.text_area("既有供應商",
                placeholder="例：大金工業、三菱電機", height=68)
            existing_quote    = st.text_input("既有報價",
                placeholder="例：NT$45,000 / 台（大金）")
            required_date     = st.date_input("需求日期", value=date.today())

        submitted = st.form_submit_button(
            "🔍 開始搜尋全球供應商", type="primary", use_container_width=True)

    if submitted:
        if not applicant or category == "請選擇…" or not product_spec or not quantity:
            st.error("請填寫所有必填欄位（＊）")
        else:
            st.session_state['form_data'] = {
                '申購人/部門':  applicant,
                '產品類別':    category,
                '品名規格':    product_spec,
                '數量':        quantity,
                '用途':        purpose,
                '預算':        budget,
                '目標市場':    ', '.join(target_market),
                '既有供應商':  existing_supplier,
                '既有報價':    existing_quote,
                '需求日期':    str(required_date),
                '填單時間':    datetime.now().strftime('%Y-%m-%d %H:%M'),
            }
            keyword   = product_spec.split('\n')[0][:60]
            suppliers = search_all_markets(keyword, target_market, serper_key)
            st.session_state['suppliers'] = suppliers

            if suppliers:
                st.success(
                    f"✅ 找到 **{len(suppliers)}** 筆供應商資料！"
                    "請切換到「供應商搜尋結果 & 匯出」分頁查看。"
                )
            else:
                st.warning("未找到相關供應商，請嘗試修改品名規格或換個市場。")

# ══════════════════════════════ 結果 ═════════════════════════════════════════
with tab2:
    st.subheader("供應商搜尋結果")

    if 'suppliers' not in st.session_state or not st.session_state['suppliers']:
        st.info("請先在「採購需求表單」填寫需求並點擊搜尋。", icon="👈")
    else:
        suppliers = st.session_state['suppliers']
        form_data = st.session_state.get('form_data', {})

        st.caption(
            f"共找到 **{len(suppliers)}** 筆供應商。"
            "可直接在表格中補充聯絡人、單價、交期等資訊，完成後匯出 Excel。"
        )

        df = pd.DataFrame(suppliers, columns=SUPPLIER_COLS)
        edited_df = st.data_editor(
            df,
            use_container_width=True,
            num_rows="dynamic",
            hide_index=True,
            column_config={
                '連結':     st.column_config.LinkColumn('連結', display_text="🔗 開啟"),
                '規格概述': st.column_config.TextColumn('規格概述', width='large'),
                '單價':     st.column_config.TextColumn('單價',     width='small'),
                '數量':     st.column_config.TextColumn('數量',     width='small'),
                '交期':     st.column_config.TextColumn('交期',     width='small'),
            },
        )

        st.divider()
        col1, col2 = st.columns(2)

        with col1:
            today    = datetime.now().strftime('%Y%m%d')
            cat      = form_data.get('產品類別', '').replace('/', '-')
            filename = f"供應商比較_{cat}_{today}.xlsx"
            st.download_button(
                label="📥 一鍵匯出 Excel 報表",
                data=to_excel(form_data, edited_df),
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )
        with col2:
            if st.button("🔄 清除結果，重新搜尋", use_container_width=True):
                for key in ['suppliers', 'form_data']:
                    st.session_state.pop(key, None)
                st.rerun()
