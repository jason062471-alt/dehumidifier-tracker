#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
供應鏈開發機器人
採購需求表單 + 全球供應商自動搜尋 + Excel 比較報表
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
    initial_sidebar_state="collapsed",
)

st.title("🔍 供應鏈開發機器人")
st.caption("填寫採購需求，自動搜尋全球供應商並匯出比較報表")

HEADERS = {
    'User-Agent': (
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
        'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36'
    ),
    'Accept-Language': 'zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7',
}

SUPPLIER_COLS = [
    '供應商名稱', '聯絡人', '聯絡方式', '品名', '型號',
    '規格概述', '數量', '單價', '交期', '付款方式', '付款條件', '來源', '連結',
]

EMPTY_SUPPLIER = {c: '' for c in SUPPLIER_COLS}

def clean(text: str) -> str:
    return re.sub(r'\s+', ' ', text).strip() if text else ''

# ── 供應商搜尋函式 ────────────────────────────────────────────────────────────

def safe_get(url, params=None):
    try:
        r = requests.get(url, params=params, headers=HEADERS, timeout=20)
        r.raise_for_status()
        return r
    except Exception:
        return None


def search_made_in_china(keyword: str) -> list:
    results = []
    slug = re.sub(r'\s+', '-', keyword.strip())
    r = safe_get(f"https://www.made-in-china.com/products-search/hot-china-products/{slug}.html")
    if not r:
        return results
    soup = BeautifulSoup(r.text, 'lxml')
    for item in soup.select('.product-item, .pro-item, [class*="product"]')[:12]:
        name_el    = item.select_one('h4, h3, .pro-name, [class*="name"]')
        price_el   = item.select_one('.price, [class*="price"]')
        company_el = item.select_one('.company-name, [class*="company"]')
        link_el    = item.select_one('a[href]')
        name    = clean(name_el.text)    if name_el    else ''
        price   = clean(price_el.text)   if price_el   else ''
        company = clean(company_el.text) if company_el else ''
        href    = link_el['href']        if link_el    else ''
        if href and not href.startswith('http'):
            href = 'https://www.made-in-china.com' + href
        if name or company:
            s = dict(EMPTY_SUPPLIER)
            s.update({'供應商名稱': company, '品名': name, '單價': price,
                      '來源': 'Made-in-China', '連結': href})
            results.append(s)
    return results


def search_taiwantrade(keyword: str) -> list:
    results = []
    r = safe_get("https://www.taiwantrade.com/search", params={'q': keyword, 'lang': 'zh-tw'})
    if not r:
        return results
    soup = BeautifulSoup(r.text, 'lxml')
    for item in soup.select('.product-card, .search-result-item, [class*="product"]')[:12]:
        name_el    = item.select_one('h3, h4, [class*="name"], [class*="title"]')
        company_el = item.select_one('[class*="company"], [class*="supplier"]')
        price_el   = item.select_one('[class*="price"]')
        link_el    = item.select_one('a[href]')
        name    = clean(name_el.text)    if name_el    else ''
        company = clean(company_el.text) if company_el else ''
        price   = clean(price_el.text)   if price_el   else ''
        href    = link_el['href']        if link_el    else ''
        if href and not href.startswith('http'):
            href = 'https://www.taiwantrade.com' + href
        if name or company:
            s = dict(EMPTY_SUPPLIER)
            s.update({'供應商名稱': company, '品名': name, '單價': price,
                      '來源': 'Taiwan Trade', '連結': href})
            results.append(s)
    return results


def search_globalsources(keyword: str) -> list:
    results = []
    slug = re.sub(r'\s+', '-', keyword.strip().lower())
    r = safe_get(f"https://www.globalsources.com/manufacturers/{slug}.html")
    if not r:
        return results
    soup = BeautifulSoup(r.text, 'lxml')
    for item in soup.select('.product-item, [class*="product"], [class*="supplier"]')[:12]:
        name_el    = item.select_one('h3, h4, [class*="name"]')
        company_el = item.select_one('[class*="company"], [class*="supplier-name"]')
        price_el   = item.select_one('[class*="price"]')
        link_el    = item.select_one('a[href]')
        name    = clean(name_el.text)    if name_el    else ''
        company = clean(company_el.text) if company_el else ''
        price   = clean(price_el.text)   if price_el   else ''
        href    = link_el['href']        if link_el    else ''
        if href and not href.startswith('http'):
            href = 'https://www.globalsources.com' + href
        if name or company:
            s = dict(EMPTY_SUPPLIER)
            s.update({'供應商名稱': company, '品名': name, '單價': price,
                      '來源': 'Global Sources', '連結': href})
            results.append(s)
    return results


def search_thomasnet(keyword: str) -> list:
    """美國供應商目錄"""
    results = []
    r = safe_get("https://www.thomasnet.com/search", params={'what': keyword, 'where': 'USA'})
    if not r:
        return results
    soup = BeautifulSoup(r.text, 'lxml')
    for item in soup.select('.supplier-row, [class*="supplier"], .profile-card')[:10]:
        company_el = item.select_one('h2, h3, [class*="name"], [class*="title"]')
        spec_el    = item.select_one('[class*="description"], [class*="spec"], p')
        link_el    = item.select_one('a[href]')
        company = clean(company_el.text) if company_el else ''
        spec    = clean(spec_el.text)    if spec_el    else ''
        href    = link_el['href']        if link_el    else ''
        if href and not href.startswith('http'):
            href = 'https://www.thomasnet.com' + href
        if company:
            s = dict(EMPTY_SUPPLIER)
            s.update({'供應商名稱': company, '規格概述': spec[:80],
                      '來源': 'ThomasNet 美國', '連結': href})
            results.append(s)
    return results


# ── Excel 匯出 ────────────────────────────────────────────────────────────────

def to_excel(form_data: dict, suppliers_df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        hdr_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#1F4E79',
            'font_color': 'white', 'border': 1, 'align': 'center',
        })
        cell_fmt = workbook.add_format({'border': 1, 'text_wrap': True})

        # ── 採購需求工作表 ──
        ws1 = workbook.add_worksheet('採購需求')
        label_fmt = workbook.add_format({'bold': True, 'bg_color': '#D6E4F0', 'border': 1})
        ws1.set_column(0, 0, 18)
        ws1.set_column(1, 1, 45)
        for row, (k, v) in enumerate(form_data.items()):
            ws1.write(row, 0, k, label_fmt)
            ws1.write(row, 1, str(v), cell_fmt)

        # ── 供應商比較工作表 ──
        ws2 = workbook.add_worksheet('供應商比較')
        for col, header in enumerate(suppliers_df.columns):
            ws2.write(0, col, header, hdr_fmt)
            ws2.set_column(col, col, 18)
        ws2.set_column(suppliers_df.columns.get_loc('規格概述'), suppliers_df.columns.get_loc('規格概述'), 30)
        ws2.set_column(suppliers_df.columns.get_loc('供應商名稱'), suppliers_df.columns.get_loc('供應商名稱'), 25)
        ws2.freeze_panes(1, 0)
        for row, record in enumerate(suppliers_df.itertuples(index=False), 1):
            for col, val in enumerate(record):
                ws2.write(row, col, str(val) if val else '', cell_fmt)

    return output.getvalue()


# ── 主介面 ────────────────────────────────────────────────────────────────────
tab1, tab2 = st.tabs(["📋 採購需求表單", "📊 供應商搜尋結果 & 匯出"])

# ═══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.subheader("採購需求表單")
    st.info("填寫完成後點擊「開始搜尋」，系統自動搜尋全球供應商。", icon="ℹ️")

    with st.form("procurement_form"):
        col1, col2 = st.columns(2)

        with col1:
            applicant    = st.text_input("申購人 / 部門 ＊", placeholder="例：採購部 / 王小明")
            category     = st.selectbox("產品類別 ＊", [
                "請選擇…", "電子零件", "機械設備", "原材料", "化學材料",
                "包裝材料", "辦公用品", "資訊設備", "空調 / 除濕設備", "其他",
            ])
            product_spec = st.text_area("品名規格 ＊", placeholder="例：工業用除濕機，除濕量 50L/day，220V", height=110)
            quantity     = st.text_input("數量 ＊", placeholder="例：10 台")
            purpose      = st.text_area("用途", placeholder="例：廠房防潮", height=80)

        with col2:
            budget            = st.text_input("預算", placeholder="例：NT$50,000 / 台以內")
            target_market     = st.multiselect("目標市場", [
                "台灣", "中國大陸", "日本", "韓國", "美國", "歐洲", "東南亞", "全球",
            ])
            existing_supplier = st.text_area("既有供應商", placeholder="例：大金工業、三菱電機", height=68)
            existing_quote    = st.text_input("既有報價", placeholder="例：NT$45,000 / 台（大金）")
            required_date     = st.date_input("需求日期", value=date.today())

        submitted = st.form_submit_button("🔍 開始搜尋全球供應商", type="primary", use_container_width=True)

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

            keyword = product_spec.split('\n')[0][:60]
            sources = [
                ("Made-in-China（中國製造）", search_made_in_china),
                ("Taiwan Trade（台灣外貿）",  search_taiwantrade),
                ("Global Sources（全球採購）", search_globalsources),
                ("ThomasNet（美國供應商）",   search_thomasnet),
            ]

            all_suppliers = []
            progress_bar = st.progress(0)
            status_text  = st.empty()

            for i, (src_name, func) in enumerate(sources):
                status_text.text(f"正在搜尋 {src_name}…")
                results = func(keyword)
                all_suppliers.extend(results)
                progress_bar.progress((i + 1) / len(sources))
                time.sleep(1)

            status_text.empty()
            progress_bar.empty()
            st.session_state['suppliers'] = all_suppliers

            if all_suppliers:
                st.success(f"✅ 找到 **{len(all_suppliers)}** 筆供應商資料！請切換到「供應商搜尋結果 & 匯出」分頁。")
            else:
                st.warning("未找到相關供應商，請嘗試修改關鍵字後重新搜尋。")

# ═══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.subheader("供應商搜尋結果")

    if 'suppliers' not in st.session_state or not st.session_state['suppliers']:
        st.info("請先在「採購需求表單」填寫需求並點擊搜尋。", icon="👈")
    else:
        suppliers  = st.session_state['suppliers']
        form_data  = st.session_state.get('form_data', {})

        st.caption(
            f"共找到 **{len(suppliers)}** 筆供應商。"
            "您可直接在表格中新增/編輯資料（聯絡人、報價、交期等），完成後匯出 Excel。"
        )

        df = pd.DataFrame(suppliers, columns=SUPPLIER_COLS)
        edited_df = st.data_editor(
            df,
            use_container_width=True,
            num_rows="dynamic",
            hide_index=True,
            column_config={
                '連結': st.column_config.LinkColumn('連結', display_text="開啟"),
                '單價': st.column_config.TextColumn('單價', width='small'),
                '數量': st.column_config.TextColumn('數量', width='small'),
                '交期': st.column_config.TextColumn('交期', width='small'),
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
