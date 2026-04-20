import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor, Inches  # 確保導入 Inches
import io

# --- 1. 字體與格式函數 ---
def set_font_kai_11(run):
    run.font.name = '標楷體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run.font.size = Pt(11)
    run.font.bold = False

def set_font_kai_bold_14(run):
    run.font.name = '標楷體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run.font.size = Pt(14)
    run.font.bold = True

# --- 2. 數據抓取函數庫 ---

def fetch_and_aggregate_lighting(file):
    try:
        xl = pd.ExcelFile(file)
        target_sheets = [s for s in xl.sheet_names if "表九之二" in s]
        if not target_sheets: return None
        aggregated_data = {}
        for sheet in target_sheets:
            df = pd.read_excel(file, sheet_name=sheet, header=None)
            for i in range(6, len(df)):
                kind_raw = str(df.iloc[i, 1]).strip()
                spec = str(df.iloc[i, 5]).strip()
                count_str = str(df.iloc[i, 9]).strip()
                hours_str = str(df.iloc[i, 11]).strip()
                if kind_raw == "nan" or "註" in kind_raw or "合計" in kind_raw: continue
                kind = kind_raw.split('.')[-1].strip() if '.' in kind_raw else kind_raw
                try:
                    count = int(float(count_str.replace(',', '')))
                    hours = int(float(hours_str.replace(',', '')))
                    key = (kind, spec, hours)
                    aggregated_data[key] = aggregated_data.get(key, 0) + count
                except: continue
        return aggregated_data
    except: return None

def fetch_chiller_spec(file):
    try:
        xl = pd.ExcelFile(file)
        target_sheets = [s for s in xl.sheet_names if "空調系統(三)" in s]
        all_chillers = []
        for sheet in target_sheets:
            df = pd.read_excel(file, sheet_name=sheet, header=None)
            for i in range(6, len(df)):
                name_raw = str(df.iloc[i, 1]).strip() # B: 設備名稱
                sn_raw = str(df.iloc[i, 2]).strip()   # C: 設備編號
                
                # --- 核心過濾邏輯 ---
                # 1. 必須包含 "主機" 
                if "主機" not in name_raw: continue
                # 2. 排除說明的關鍵字
                if any(x in name_raw for x in ["效率", "標準", "請依", "IE1", "IE2"]): continue
                # 3. 如果編號是 nan，代表這不是真正的設備列
                if sn_raw == "nan" or sn_raw == "": continue
                
                # --- 通過過濾後才開始解析數據 ---
                name = name_raw.split('.')[-1].strip() if '.' in name_raw else name_raw
                sn = sn_raw
                form = str(df.iloc[i, 5]).strip()
                inv = "變頻" if str(df.iloc[i, 7]).strip() == "有" else "定頻"
                volt = str(df.iloc[i, 11]).strip()
                pwr = str(df.iloc[i, 12]).strip()
                yr = str(df.iloc[i, 13]).strip().replace('民國','').replace('年','')
                cap = str(df.iloc[i, 14]).strip()
                unit = str(df.iloc[i, 15]).strip().upper()
                qty = str(df.iloc[i, 21]).strip()
                
                # 如果容量也是 nan，代表這一列無效
                if cap == "nan" or cap == "": continue

                try:
                    v = float(cap.replace(',',''))
                    rt = round(v/3.517,1) if "KW" in unit else (round(v/3024,1) if "KCAL" in unit else v)
                except: rt = cap
                
                all_chillers.append([name, sn, form, volt, pwr, yr, rt, qty, inv])
        return all_chillers
    except: return None

def fetch_pump_and_cooling_data(file):
    try:
        xl = pd.ExcelFile(file)
        target_sheets = [s for s in xl.sheet_names if "空調系統(三)" in s]
        data = {"冰水泵": [], "區域水泵": [], "冷卻水泵": [], "冷卻水塔": []}
        has_sec = False
        for sheet in target_sheets:
            df = pd.read_excel(file, sheet_name=sheet, header=None)
            for i in range(6, len(df)):
                name = str(df.iloc[i, 1]).strip()
                sn = str(df.iloc[i, 2]).strip()
                inv = "變頻" if str(df.iloc[i, 7]).strip() == "有" else "定頻"
                cap_val = str(df.iloc[i, 14]).strip()
                unit = str(df.iloc[i, 15]).strip().upper()
                hp_val = str(df.iloc[i, 18]).strip()
                qty = str(df.iloc[i, 21]).strip()
                if cap_val == "nan" or hp_val == "nan": continue

                if "冰水泵" in name or "區域水泵" in name or "冷卻水泵" in name:
                    target = "冰水泵" if "冰水泵" in name else ("區域水泵" if "區域水泵" in name else "冷卻水泵")
                    if target == "區域水泵": has_sec = True
                    try:
                        f_lpm = float(cap_val.replace(',','')) * 3.785 if "GPM" in unit else float(cap_val.replace(',',''))
                        hp = float(hp_val)
                        head = round((hp * 462.52) / (f_lpm / 60 * 9.8), 1) if f_lpm > 0 else 0
                        data[target].append([sn, hp_val, int(f_lpm), head, qty, inv])
                    except: continue
                elif "冷卻水塔" in name:
                    data["冷卻水塔"].append([sn, hp_val, cap_val, qty, inv])
        return data, has_sec
    except: return None, False
def fetch_other_systems(file):
    try:
        xl = pd.ExcelFile(file)
        # 模糊搜尋所有包含「表九之三」的分頁 (支援多建築物)
        target_sheets = [s for s in xl.sheet_names if "表九之三" in s]
        if not target_sheets: return None
        
        all_data = []
        for sheet in target_sheets:
            df = pd.read_excel(file, sheet_name=sheet, header=None)
            # 數據從 index 6 開始 (B欄=1, C欄=2, J欄=9, K欄=10, T欄=19, X欄=23)
            for i in range(6, len(df)):
                sys_raw = str(df.iloc[i, 1]).strip()    # B: 系統名稱
                name_raw = str(df.iloc[i, 2]).strip()   # C: 設備名稱
                volt = str(df.iloc[i, 9]).strip()       # J: 電壓
                pwr = str(df.iloc[i, 10]).strip()       # K: 功率值
                qty = str(df.iloc[i, 19]).strip()       # T: 數量
                hours = str(df.iloc[i, 23]).strip()     # X: 運轉時數

                # 過濾：排除空行、合計行、或註解行
                if sys_raw in ["nan", "None", ""] or "合計" in sys_raw or "註" in sys_raw:
                    continue
                if name_raw in ["nan", "None", ""]:
                    continue

                # 清洗：系統名稱排除「1. 」數字開頭
                sys_name = sys_raw.split('.')[-1].strip() if '.' in sys_raw else sys_raw
                
                # 數值格式清洗 (去除小數點後多餘的0)
                def clean_num(v):
                    return v.replace('.0', '') if v.endswith('.0') else v

                all_data.append([sys_name, name_raw, clean_num(volt), clean_num(pwr), clean_num(qty), clean_num(hours)])
        
        return all_data
    except:
        return None
# --- 3. Word 生成函數庫 ---

def add_lighting_table(doc, data):
    p = doc.add_paragraph(); run = p.add_run("2. 照明系統："); set_font_kai_bold_14(run)
    table = doc.add_table(rows=2, cols=4); table.style = 'Table Grid'
    table.cell(0,1).merge(table.cell(0,2)); table.cell(0,0).merge(table.cell(1,0)); table.cell(0,3).merge(table.cell(1,3))
    hdrs = [(0,0,"燈具種類"),(0,1,"燈具形式"),(0,3,"運轉時數\n(小時/年)"),(1,1,"容量規格"),(1,2,"數量")]
    for r,c,t in hdrs:
        cp = table.cell(r,c).paragraphs[0]; cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_font_kai_11(cp.add_run(t))
    for (k,s,h), count in sorted(data.items()):
        row = table.add_row().cells
        for i, v in enumerate([k,s,str(count),str(h)]):
            cp = row[i].paragraphs[0]; cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_font_kai_11(cp.add_run(v))

def add_ac_mode_table(doc, ac_data):
    p3 = doc.add_paragraph(); run3 = p3.add_run("3. 空調系統："); set_font_kai_bold_14(run3)
    p1 = doc.add_paragraph(); p1.paragraph_format.left_indent = Pt(20)
    run1 = p1.add_run("(1) 空調主機開啟模式："); set_font_kai_bold_14(run1)
    table = doc.add_table(rows=4, cols=6); table.style = 'Table Grid'
    headers = ["季節", "主機總容量\n(RT)", "冰機總開啟台數", "負載率\n(%)", "合計容量\n(RT)", "出水溫度設定 (°C)"]
    for i, h in enumerate(headers):
        cp = table.cell(0,i).paragraphs[0]; cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_font_kai_11(cp.add_run(h))
    for r_idx, row_vals in enumerate(ac_data, 1):
        for c_idx, val in enumerate(row_vals):
            cp = table.cell(r_idx, c_idx).paragraphs[0]; cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_font_kai_11(cp.add_run(str(val)))

def add_chiller_spec_table(doc, chiller_data):
    p2 = doc.add_paragraph(); p2.paragraph_format.left_indent = Pt(20)
    run2 = p2.add_run("(2) 冰水主機規格："); set_font_kai_bold_14(run2)
    table = doc.add_table(rows=1, cols=9); table.style = 'Table Grid'
    headers = ["設備名稱", "設備編號", "型式", "電壓\n(V)", "功率值\n(kW)", "製造年份", "容量\n(RT)", "現有數量", "備註"]
    for i, h in enumerate(headers):
        cp = table.cell(0,i).paragraphs[0]; cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_font_kai_11(cp.add_run(h))
    for r_vals in chiller_data:
        row = table.add_row().cells
        for i, v in enumerate(r_vals):
            cp = row[i].paragraphs[0]; cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_font_kai_11(cp.add_run(str(v).replace('.0','')))

def add_pump_section(doc, data, has_sec):
    p3 = doc.add_paragraph(); p3.paragraph_format.left_indent = Pt(20)
    set_font_kai_bold_14(p3.add_run("(3) 冰水管路系統："))
    side_txt = "採一/二次側系統設計" if has_sec else "採一次側系統設計"
    ps = doc.add_paragraph(); ps.paragraph_format.left_indent = Pt(40)
    set_font_kai_11(ps.add_run(f"{side_txt}，設備規格說明如下表所示"))
    for p_name in ["冰水泵", "區域水泵"]:
        items = data.get(p_name, [])
        if not items: continue
        table = doc.add_table(rows=2, cols=6); table.style = 'Table Grid'
        table.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_cell = table.cell(0,0).merge(table.cell(0,5))
        set_font_kai_11(title_cell.paragraphs[0].add_run(p_name))
        title_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        h = ["設備編號", "額定馬力\n(HP)", "額定流量\n(LPM)", "揚程\n(m)", "數量\n(台)", "備註"]
        for i, txt in enumerate(h):
            cp = table.cell(1,i).paragraphs[0]; cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_font_kai_11(cp.add_run(txt))
        for r_vals in items:
            row = table.add_row().cells
            for i, v in enumerate(r_vals):
                cp = row[i].paragraphs[0]; cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_font_kai_11(cp.add_run(str(v).replace('.0','')))
        doc.add_paragraph()

def add_cooling_section(doc, data):
    p4 = doc.add_paragraph(); p4.paragraph_format.left_indent = Pt(20)
    set_font_kai_bold_14(p4.add_run("(4) 冷卻水系統："))
    for p_name in ["冷卻水泵", "冷卻水塔"]:
        items = data.get(p_name, [])
        if not items: continue
        num_cols = 6 if p_name == "冷卻水泵" else 5
        table = doc.add_table(rows=2, cols=num_cols); table.style = 'Table Grid'
        table.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_cell = table.cell(0,0).merge(table.cell(0,num_cols-1))
        set_font_kai_11(title_cell.paragraphs[0].add_run(p_name))
        title_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        h = ["設備編號", "額定馬力\n(HP)", "額定流量\n(LPM)", "揚程\n(m)", "數量\n(台)", "備註"] if p_name == "冷卻水泵" else ["設備編號", "額定馬力\n(HP)", "額定容量\n(RT)", "數量\n(台)", "備註"]
        for i, txt in enumerate(h):
            cp = table.cell(1,i).paragraphs[0]; cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_font_kai_11(cp.add_run(txt))
        for r_vals in items:
            row = table.add_row().cells
            for i, v in enumerate(r_vals):
                cp = row[i].paragraphs[0]; cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_font_kai_11(cp.add_run(str(v).replace('.0','')))
        doc.add_paragraph()
    p5 = doc.add_paragraph(); p5.paragraph_format.left_indent = Pt(20)
    set_font_kai_bold_14(p5.add_run("(5) 空調附屬設備："))
    set_font_kai_11(p5.add_run("空氣側採用空調箱或小型送風機供應空調至現場使用。"))
def add_other_systems_table(doc, other_data):
    # 標題 4.其他系統： (標楷加粗 14號)
    p = doc.add_paragraph()
    run = p.add_run("4.其他系統：")
    set_font_kai_bold_14(run)

    # 建立表格 (6欄)
    table = doc.add_table(rows=2, cols=6)
    table.style = 'Table Grid'
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 跨列合併處理表頭
    table.cell(0, 0).merge(table.cell(1, 0)) # 系統名稱
    table.cell(0, 1).merge(table.cell(1, 1)) # 設備名稱
    table.cell(0, 2).merge(table.cell(0, 3)) # 設備電功率 (跨兩欄)
    table.cell(0, 4).merge(table.cell(1, 4)) # 現有數量
    table.cell(0, 5).merge(table.cell(1, 5)) # 運轉時數

    headers = [
        (table.cell(0, 0), "系統名稱"),
        (table.cell(0, 1), "設備名稱"),
        (table.cell(0, 2), "設備電功率"),
        (table.cell(1, 2), "電壓(伏特)"),
        (table.cell(1, 3), "功率值(瓩)"),
        (table.cell(0, 4), "現有數量\n(台)"),
        (table.cell(0, 5), "運轉時數\n(小時/年)")
    ]

    for cell, text in headers:
        cell.vertical_alignment = WD_ALIGN_PARAGRAPH.CENTER
        cp = cell.paragraphs[0]
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_font_kai_11(cp.add_run(text))

    # 填入數據 (黑字標楷 11號)
    for row_vals in other_data:
        cells = table.add_row().cells
        for i, val in enumerate(row_vals):
            cp = cells[i].paragraphs[0]
            cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_font_kai_11(cp.add_run(str(val)))
def add_site_photos_table(doc):
    # 1. 大標題：現場節能輔導照片 (標楷、黑色、18號、加粗)
    p_main = doc.add_paragraph()
    p_main.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_main = p_main.add_run("現場節能輔導照片")
    run_main.font.name = '標楷體'
    run_main._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run_main.font.size = Pt(18)
    run_main.font.bold = True

    # 2. 定義設備清單 (兩兩一組)
    items = [
        ["大樓外觀", "冰水主機"],
        ["水泵群", "冷卻水塔"],
        ["高壓變電站", "空調監控系統"]
    ]

    # 3. 建立表格 (共 6 列，3列放圖，3列放字)
    table = doc.add_table(rows=0, cols=2)
    table.style = 'Table Grid'
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for row_pair in items:
        # --- 新增圖片行 ---
        image_row = table.add_row()
        image_row.height = Inches(2.2) # 設定圖片框高度
        for cell in image_row.cells:
            cell.width = Inches(3.0)

        # --- 新增文字行 ---
        text_row = table.add_row()
        for i, name in enumerate(row_pair):
            cell = text_row.cells[i]
            cell.vertical_alignment = WD_ALIGN_PARAGRAPH.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT # 文字靠左對齊
            run = p.add_run(name)
            # 設定字體 (標楷、黑色、12號)
            run.font.name = '標楷體'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
            run.font.size = Pt(12)
            run.font.bold = False
# 自動調整欄寬 (總寬約 6 英吋)
    for row in table.rows:
        for cell in row.cells:
            cell.width = Inches(3.5)
    # 在表格最後加上分節符號
    doc.add_page_break()
# --- 4. Streamlit 介面 ---
st.subheader("⚙️ 設備系統資料庫")
c0, c1, c2, c3, c4 = st.columns([0.8, 1.5, 1.2, 1.2, 1.5])
with c1:
    # 將最小值 min_value 設為 0，這樣就不會被限制在只能調大
    rt_s = st.number_input("夏容量", min_value=0, value=600, key="rt_s")
    rt_sp = st.number_input("春容量", min_value=0, value=450, key="rt_sp")
    rt_w = st.number_input("冬容量", min_value=0, value=450, key="rt_w")
with c2:
    ct_s = st.number_input("夏台", min_value=0, value=1, key="ct_s")
    ct_sp = st.number_input("春台", min_value=0, value=1, key="ct_sp")
    ct_w = st.number_input("冬台", min_value=0, value=1, key="ct_w")
with c3:
    ld_s = st.number_input("夏負", min_value=0, max_value=100, value=70, key="ld_s")
    ld_sp = st.number_input("春負", min_value=0, max_value=100, value=70, key="ld_sp")
    ld_w = st.number_input("冬負", min_value=0, max_value=100, value=60, key="ld_w")
with c4:
    tp_s = st.number_input("夏溫", min_value=0, value=7, key="tp_s")
    tp_sp = st.number_input("春溫", min_value=0, value=7, key="tp_sp")
    tp_w = st.number_input("冬溫", min_value=0, value=7, key="tp_w")

ac_rows = [
    ["夏季", rt_s, ct_s, f"{ld_s}%", round(rt_s*ld_s/100, 1), tp_s],
    ["春秋", rt_sp, ct_sp, f"{ld_sp}%", round(rt_sp*ld_sp/100, 1), tp_sp],
    ["冬季", rt_w, ct_w, f"{ld_w}%", round(rt_w*ld_w/100, 1), tp_w]
]

st.markdown("---")
up_file = st.file_uploader("請上傳 Excel", type=["xlsx"])
final_file = up_file if up_file else st.session_state.get('global_excel')

# --- 5. 下載與打包整合邏輯 ---
st.markdown("---")

if final_file:
    # 建立一個生成 Word 的函數 (供按鈕使用)
    def create_word_blob():
        doc = Document()
        # 1. 照明系統
        l_data = fetch_and_aggregate_lighting(final_file)
        if l_data:
            add_lighting_table(doc, l_data)
            doc.add_paragraph()
        # 2. 空調模式
        add_ac_mode_table(doc, ac_rows)
        # 3. 冰水主機規格
        c_data = fetch_chiller_spec(final_file)
        if c_data:
            add_chiller_spec_table(doc, c_data)
            doc.add_paragraph()
        # 4. 泵浦與冷卻水系統
        p_data, has_sec = fetch_pump_and_cooling_data(final_file)
        if p_data:
            add_pump_section(doc, p_data, has_sec)
            add_cooling_section(doc, p_data)
        # 5. 其他系統
        o_data = fetch_other_systems(final_file)
        if o_data:
            add_other_systems_table(doc, o_data)
        # 6. 照片頁面
        doc.add_page_break()
        add_site_photos_table(doc)

        buf = io.BytesIO()
        doc.save(buf)
        return buf.getvalue()

    # --- 關鍵修正：兩步操作 ---
    col_btn1, col_btn2 = st.columns(2)

    with col_btn1:
        # 第一步：確認數值並存入打包中心
        if st.button("🔄 確認數值並同步至打包中心", use_container_width=True):
            word_data = create_word_blob()
            if 'report_warehouse' not in st.session_state:
                st.session_state['report_warehouse'] = {}
            # 存入打包中心
            st.session_state['report_warehouse']['設備系統報告'] = word_data
            st.success("✅ 數據已鎖定，左側打包下載已更新！")
            st.rerun() # 強制重新整理讓左側數量更新

    with col_btn2:
        # 第二步：下載 (只有同步過後才能下載)
        if 'report_warehouse' in st.session_state and '設備系統報告' in st.session_state['report_warehouse']:
            st.download_button(
                label="📥 下載目前的 Word 報告",
                data=st.session_state['report_warehouse']['設備系統報告'],
                file_name="設備報告.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        else:
            st.button("📥 請先點擊左側確認同步", disabled=True, use_container_width=True)

else:
    st.info("💡 請先上傳 Excel 檔案以啟用報告生成功能。")
