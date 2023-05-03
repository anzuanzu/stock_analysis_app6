import pandas as pd
import numpy as np
import streamlit as st
import yfinance as yf
from datetime import datetime
from datetime import timedelta
import xlsxwriter
import io
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle

def get_stock_info(stock_id, avg_eps):
    stock_history = yf.download(
        f"{stock_id}.TW", start=datetime.now() - timedelta(days=7), end=datetime.now())
    if stock_history.shape[0] >= 2:
        current_price = stock_history.iloc[-1]["Close"]
        price_change = current_price / stock_history.iloc[-2]["Close"] - 1
        pe_ratio = current_price / avg_eps
        return current_price, price_change, pe_ratio
    else:
        return np.nan, np.nan, np.nan

def calculate_std(row, recent_years):
    eps_values = row[recent_years]
    return np.std(eps_values)


def analyze_data(file, current_year):
    data = pd.read_excel(file, engine='openpyxl')

    years = [f"{year}年度每股盈餘(元)" for year in range(
        current_year - 1, current_year - 6, -1)]
    dividend_years = [f"{year}合計股利" for year in range(
        current_year - 1, current_year - 6, -1)]

    data["盈餘標準差"] = data.apply(lambda row: calculate_std(row, years), axis=1)
    data["近5年平均EPS(元)"] = data[years].mean(axis=1)
    data["近5年平均合計股利(元)"] = data[dividend_years].mean(axis=1)
    data["股利發放率"] = data["近5年平均合計股利(元)"] / data["近5年平均EPS(元)"]

    return data


st.title("股票分析工具")

uploaded_file = st.file_uploader("選擇一個 .xlsx 文件", type="xlsx")
current_year = st.number_input(
    "請輸入當前年份：", min_value=1900, max_value=9999, value=2023, step=1)

filtered_data = pd.DataFrame()  # 初始化 filtered_data

if uploaded_file is not None:
    data = analyze_data(uploaded_file, current_year)

    stock_id = st.text_input("輸入股票代號：")
    if stock_id:
        display_data = data[data["代號"].apply(
            lambda x: str(x).startswith(stock_id))]
        if not display_data.empty:
            st.write(display_data[["代號", "名稱", "盈餘標準差",
                     "近5年平均EPS(元)", "近5年平均合計股利(元)", "股利發放率"]])
        else:
            st.warning("找不到該股票代號。")

    # 增加篩選功能
    st.header("篩選功能")
    eps_min_filter = st.number_input("近5年平均EPS(元)最小值：", value=1.0, step=0.1)
    eps_max_filter = st.number_input("近5年平均EPS(元)最大值：", value=1000.0, step=0.1)
    div_min_filter = st.number_input("近5年平均合計股利(元)最小值：", value=1.0, step=0.1)
    div_max_filter = st.number_input(
        "近5年平均合計股利(元)最大值：", value=1000.0, step=0.1)
    std_min_filter = st.number_input("盈餘標準差最小值：", value=0.0, step=0.1)
    std_max_filter = st.number_input("盈餘標準差最大值：", value=1.0, step=0.1)
    payout_min_filter = st.number_input("股利發放率最小值：", value=0.7, step=0.01)
    payout_max_filter = st.number_input("股利發放率最大值：", value=1.0, step=0.01)
    filtered_data = data[(data["近5年平均EPS(元)"] >= eps_min_filter) & (data["近5年平均EPS(元)"] <= eps_max_filter) &
                         (data["近5年平均合計股利(元)"] >= div_min_filter) & (data["近5年平均合計股利(元)"] <= div_max_filter) &
                         (data["盈餘標準差"] >= std_min_filter) & (data["盈餘標準差"] <= std_max_filter) &
                         (data["股利發放率"] >= payout_min_filter) & (data["股利發放率"] <= payout_max_filter)]
if st.button("篩選"):
    if not filtered_data.empty:
        # 获取最新报价、涨跌幅和本益比
        filtered_data["最新報價"], filtered_data["漲跌幅"], filtered_data["本益比"] = zip(
            *filtered_data.apply(lambda x: get_stock_info(x["代號"], x["近5年平均EPS(元)"]), axis=1))
        st.write(filtered_data[[
                 "代號", "名稱", "最新報價", "漲跌幅", "本益比", "盈餘標準差", "近5年平均EPS(元)", "近5年平均合計股利(元)", "股利發放率"]])
if not filtered_data.empty:
    # 匯出功能
    st.header("匯出結果")
    export_format = st.radio("選擇匯出格式：", ["XLSX"])

    if st.button("匯出", key="export_button"):
        if export_format == "XLSX":
            output = io.BytesIO()
            writer = pd.ExcelWriter(output, engine='xlsxwriter')
            filtered_data.to_excel(writer, sheet_name='Filtered Stocks', index=False)
            writer.close()
            output.seek(0)
            st.download_button(
                    label="下載XLSX文件",
                    data=output,
                    file_name='filtered_stocks.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
    else:
        st.warning("沒有符合篩選條件的股票。")

else:
    st.warning("請選擇一個文件進行分析。")
