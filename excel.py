import pandas as pd
import yfinance as yf
from openpyxl import load_workbook
from openpyxl.styles import Font
import os

# 設定輸入與輸出的檔案名稱
input_filename = '投資績效.xlsx'
output_filename = '投資績效_計算結果.xlsx'

# 設定要查詢的時間範圍
# 注意：yfinance 的結束日期不包含當天，因此設定至 10-16 以包含 10-15 的資料
start_date = '2024-10-02'
end_date = '2024-10-16' 

def main():
    if not os.path.exists(input_filename):
        print(f"找不到檔案：{input_filename}，請確認檔案與程式在同一個資料夾。")
        return

    print(f"正在讀取 {input_filename}...")
    # 假設 Excel 檔案格式為：
    # 包含 '股票代號'、'持有股數' 及 '持有成本' 的欄位
    df = pd.read_excel(input_filename)
    
    if '股票代號' not in df.columns or '持有股數' not in df.columns:
        print("Excel 表頭必須包含 '股票代號' 與 '持有股數'。")
        return

    symbols = df['股票代號'].dropna().astype(str).tolist()
    # 自動補上台灣股市的後綴 (.TW 或 .TWO)，以符合 Yahoo Finance 格式
    formatted_symbols = []
    for sym in symbols:
        if not sym.endswith('.TW') and not sym.endswith('.TWO'):
            formatted_symbols.append(f"{sym}.TW")
        else:
            formatted_symbols.append(sym)
            
    print(f"正在查詢股票代號：{formatted_symbols}")
    
    try:
        # 下載歷史收盤價資料
        stock_data = yf.download(formatted_symbols, start=start_date, end=end_date)['Close']
    except Exception as e:
        print(f"下載股價時發生錯誤：{e}")
        return

    if len(formatted_symbols) == 1:
        stock_data = stock_data.to_frame(name=formatted_symbols[0])

    result_df = df.copy()
    
    # 逐日比對並計算盈虧
    for date in stock_data.index:
        date_str = date.strftime('%m/%d')
        price_col_name = f"{date_str} 價格"
        pnl_col_name = f"{date_str} 盈虧"
        
        result_df[price_col_name] = 0.0
        result_df[pnl_col_name] = 0.0
        
        for idx, row in df.iterrows():
            original_sym = str(row['股票代號'])
            sym = original_sym if original_sym.endswith('.TW') or original_sym.endswith('.TWO') else f"{original_sym}.TW"
            shares = row['持有股數']
            
            if sym in stock_data.columns and not pd.isna(stock_data.loc[date, sym]):
                price = stock_data.loc[date, sym]
                result_df.at[idx, price_col_name] = round(price, 2)
                
                # 若有持有成本則計算真實盈虧，否則僅計算當日市值變化
                if '持有成本' in df.columns:
                    cost = row['持有成本']
                    result_df.at[idx, pnl_col_name] = round((price - cost) * shares, 2)
                else:
                    # 假設無成本，先顯示當日持股市值
                    result_df.at[idx, pnl_col_name] = round(price * shares, 2)

    print(f"計算完成，正在儲存至 {output_filename} ...")
    result_df.to_excel(output_filename, index=False)
    
    # 使用 openpyxl 將負數格式化為紅色
    wb = load_workbook(output_filename)
    ws = wb.active
    red_font = Font(color="FF0000")
    
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            col_name = ws.cell(row=1, column=cell.column).value
            # 針對包含 '盈虧' 的欄位進行檢查
            if col_name and '盈虧' in str(col_name):
                try:
                    if float(cell.value) < 0:
                        cell.font = red_font
                except (ValueError, TypeError):
                    pass
                    
    wb.save(output_filename)
    print("檔案處理完畢！請查看輸出檔案。")

if __name__ == "__main__":
    main()