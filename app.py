import streamlit as st
import openpyxl
from PIL import Image as PILImage, ImageOps
import io
import os

st.title("現場報告書 生成アプリ (テンプレート方式)")

# 報告書の種類を選択
report_type = st.radio("作成する報告書を選択", ["作業報告書", "トラブル報告書"])

# 入力項目
manage_no = st.text_input("管理番号　例：E99-9999")
client = st.text_input("顧客/担当者　例：株式会社○○○　※正式名称記載")
equipment = st.text_input("設備名称　例：スタッカークレーン　1号機")
date_val = st.date_input("作成日　※カレンダーより選択")
work_title = st.text_input("内容　例：フォークタイムオーバー　調査")
workers = st.text_input("作業者　※苗字のみ")

if report_type == "トラブル報告書":
    history = st.text_area("１．経緯")
    check_status = st.text_area("２．現場状況")
    investigation = st.text_area("３．調査内容・状況")
    result = st.text_area("４．調査結果")
    action = st.text_area("５．対応内容・処置内容")
    remarks = st.text_area("６．備考・まとめ")
    img_t1 = st.file_uploader("状況写真1", type=["png", "jpg"])
    img_t2 = st.file_uploader("状況写真2", type=["png", "jpg"])

st.subheader("交換・処置写真")
img_b = st.file_uploader("前（旧品）", type=["png", "jpg"])
img_a = st.file_uploader("後（新品）", type=["png", "jpg"])

# 【重要】結合セル対応の転記関数
def write_cell(ws, cell_name, value):
    cell = ws[cell_name]
    # 結合セルなら左上のセルに書き込む
    if isinstance(cell, openpyxl.cell.cell.MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                cell = merged_range.start_cell
                break
    cell.value = value

def edit_excel(template_name):
    if not os.path.exists(template_name):
        return None
        
    wb = openpyxl.load_workbook(template_name)
    ws = wb.active
    
    # 書き込み (templateの実際のセル番地に合わせてください)
    write_cell(ws, "A2", manage_no)
    write_cell(ws, "E3", str(date_val))
    write_cell(ws, "E3", client)
    write_cell(ws, "C6", equipment)
    write_cell(ws, "C8", work_title)
    write_cell(ws, "G7", workers)

    if report_type == "トラブル報告書":
        write_cell(ws, "C15", history)
        write_cell(ws, "C17", check_status)
        write_cell(ws, "C19", investigation)
        write_cell(ws, "C21", result)
        write_cell(ws, "C23", action)
        write_cell(ws, "C25", remarks)

    # 写真貼り付け
    def add_img(file, cell):
        if file:
            img = PILImage.open(file)
            img = ImageOps.exif_transpose(img)
            img.thumbnail((200, 200))
            path = f"tmp_{cell}.png"
            img.save(path)
            xl_img = openpyxl.drawing.image.Image(path)
            ws.add_image(xl_img, cell)
            return path
        return None
    
    img_paths = [add_img(img_b, "B20"), add_img(img_a, "F20")]
    
    out = io.BytesIO()
    wb.save(out)
    
    for p in img_paths:
        if p and os.path.exists(p): os.remove(p)
        
    return out

# ダウンロードボタン
if st.button("Excelを作成"):
    template = "template_sagyo.xlsx" if report_type == "作業報告書" else "template_trouble.xlsx"
    out = edit_excel(template)
    if out:
        st.download_button("📥 ダウンロード", out.getvalue(), f"{manage_no}_報告書.xlsx")
    else:
        st.error(f"ファイル {template} が見つかりません。")
