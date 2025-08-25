# patent_billing_generator.py ------ 兼容列数 + 空值容错

import os, re
import streamlit as st
from pathlib import Path
from datetime import date
import pandas as pd
from docx import Document
from openpyxl import load_workbook
import tempfile
import zipfile
import io

# ------------------ 工具函数 ------------------

def number_to_upper(amount: int) -> str:
    CN_NUM = ['零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖']
    CN_UNIT = ['', '拾', '佰', '仟', '万', '拾万', '佰万', '仟万', '亿']
    if amount == 0:
        return "零元整"
    s = str(amount)
    result = []
    for i, ch in enumerate(s[::-1]):
        digit = int(ch)
        if digit != 0:
            result.append(f"{CN_NUM[digit]}{CN_UNIT[i]}")
        else:
            if result and not result[-1].startswith('零'):
                result.append('零')
    upper_str = ''.join(reversed(result))
    upper_str = re.sub(r'零{2,}', '零', upper_str)
    upper_str = re.sub(r'零元', '元', upper_str)
    upper_str = re.sub(r'零万', '万', upper_str)
    upper_str = re.sub(r'亿万', '亿', upper_str)
    if not upper_str.endswith('元'):
        upper_str += "元整"
    return upper_str

def sanitize_filename(filename: str) -> str:
    illegal_chars = r'[<>:\"/\\|?*\x00-\x1F]'
    filename = re.sub(illegal_chars, '_', filename)
    return filename.rstrip('. ')

# ------------------ 处理单个分割 ------------------

def process_split_group(split_no, sub_df: pd.DataFrame, output_dir: Path,
                      word_template_path: Path, company_name: str):
    print(f"\n>>> 处理分割号 {split_no}，共 {len(sub_df)} 条")
    applicant = str(sub_df["申请人"].iloc[0]) if "申请人" in sub_df.columns else ""
    # 空值→0 再求和
    official_total = pd.to_numeric(sub_df["官费"], errors="coerce").fillna(0).astype(int).sum()
    agent_total = pd.to_numeric(sub_df["代理费"], errors="coerce").fillna(0).astype(int).sum()
    grand_total = official_total + agent_total

    # 序号列处理：无论原表有没有"序号"，都重建
    sub_df = sub_df.rename(columns={"分割号": "序号"})
    if "序号" in sub_df.columns:
        sub_df = sub_df.drop(columns=["序号"])
    sub_df.insert(0, "序号", range(1, len(sub_df) + 1))

    # Word 模板
    if not word_template_path.exists():
        raise FileNotFoundError("Word template not found")
    doc = Document(word_template_path)

    # 正文占位符
    for p in doc.paragraphs:
        p.text = p.text.replace("{{申请人}}", applicant) \
                      .replace("{{合计}}", str(grand_total)) \
                      .replace("{{大写}}", number_to_upper(grand_total)) \
                      .replace("{{日期}}", date.today().strftime("%Y年%m月%d日"))

    # 表格处理
    if not doc.tables:
        raise ValueError("模板中未找到表格")
    tbl = doc.tables[0]

    # 表头
    hdr_cells = tbl.rows[0].cells
    for idx, col_name in enumerate(sub_df.columns):
        if idx >= len(hdr_cells):
            tbl.add_column(width=None)
            hdr_cells = tbl.rows[0].cells
        hdr_cells[idx].text = str(col_name)

    # 数据行
    for _, row in sub_df.iterrows():
        new_cells = tbl.add_row().cells
        for idx, col_name in enumerate(sub_df.columns):
            if idx >= len(new_cells):
                break
            new_cells[idx].text = str(row[col_name] or "")

    # 合计行
    try:
        off_idx = sub_df.columns.get_loc("官费")
        agt_idx = sub_df.columns.get_loc("代理费")
        sum_idx = agt_idx + 1
    except KeyError:
        off_idx = 0
        agt_idx = 1
        sum_idx = 2

    row = tbl.add_row()
    cells = row.cells
    merge_start = cells[0]
    merge_end = cells[off_idx - 1] if off_idx > 0 else cells[0]
    if merge_start != merge_end:
        merge_start.merge(merge_end)
    merge_start.text = "合计"
    for p in merge_start.paragraphs:
        p.alignment = 2

    if off_idx < len(cells):
        cells[off_idx].text = str(official_total)
    if agt_idx < len(cells):
        cells[agt_idx].text = str(agent_total)
    if sum_idx < len(cells):
        cells[sum_idx].text = str(grand_total)

    filename = sanitize_filename(f"请款单（{applicant}-{grand_total}元-{company_name}-{date.today().strftime('%Y-%m-%d')}).docx")
    doc.save(output_dir / filename)
    print(f"✅ 已生成请款单：{filename}")
    return {
        "分割号": split_no,
        "申请人": applicant,
        "总官费": official_total,
        "总代理费": agent_total,
        "总计": grand_total,
        "文件名": filename,
    }

# ------------------ 生成发票申请汇总 Excel ------------------

def generate_invoice_excel(rows: list, output_dir: Path, excel_template_path: Path, company_name: str):
    if not rows:
        print("⚠ 无数据可汇总")
        return None
    if not excel_template_path.exists():
        st.error("Excel模板文件不存在，请上传正确路径")
        return None

    wb = load_workbook(excel_template_path)
    ws = wb[wb.sheetnames[0]]

    start_row = 2
    while ws.cell(row=start_row, column=1).value is not None:
        start_row += 1

    for r in rows:
        ws.cell(row=start_row, column=2, value="普通发票（电子）")
        ws.cell(row=start_row, column=3, value=r["申请人"])
        ws.cell(row=start_row, column=7, value=r["总官费"])
        ws.cell(row=start_row, column=8, value=r["总官费"])
        ws.cell(row=start_row, column=9, value=r["总计"])
        ws.cell(row=start_row, column=17, value=date.today().strftime("%Y年%m月%d日"))
        start_row += 1

        ws.cell(row=start_row, column=2, value="专用发票（电子）")
        ws.cell(row=start_row, column=3, value=r["申请人"])
        ws.cell(row=start_row, column=7, value=r["总代理费"])
        ws.cell(row=start_row, column=8, value=r["总代理费"])
        ws.cell(row=start_row, column=9, value=r["总计"])
        ws.cell(row=start_row, column=17, value=date.today().strftime("%Y年%m月%d日"))
        start_row += 1

    excel_filename = f"发票申请表-{company_name}-{date.today().strftime('%Y-%m-%d')}.xlsx"
    wb.save(output_dir / excel_filename)
    print(f"📊 发票申请表已生成：{output_dir / excel_filename}")
    return excel_filename

# ------------------ Streamlit 界面 ------------------

def main():
    st.set_page_config(
        page_title="专利请款单生成器",
        page_icon="📄",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        color: #1E88E5;
        text-align: center;
        margin-bottom: 2rem;
        font-weight: bold;
    }
    .blue-card {
        background-color: #E3F2FD;
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #1E88E5;
        margin-bottom: 1rem;
    }
    .success-card {
        background-color: #E8F5E9;
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #4CAF50;
        margin-bottom: 1rem;
    }
    .download-section {
        background-color: #F5F5F5;
        padding: 1.5rem;
        border-radius: 10px;
        border: 2px solid #BBDEFB;
    }
    .company-selector {
        background-color: #E8EAF6;
        padding: 1rem;
        border-radius: 10px;
        margin-top: 2rem;
        text-align: center;
    }
    .generate-button {
        background-color: #1E88E5 !important;
        color: white !important;
        border: none !important;
        padding: 0.5rem 1.5rem !important;
        font-size: 1rem !important;
        border-radius: 8px !important;
        margin: 1rem auto !important;
        display: block !important;
        width: fit-content !important;
    }
    </style>
    """, unsafe_allow_html=True)

    st.markdown('<h1 class="main-header">📄 专利请款单生成器</h1>', unsafe_allow_html=True)

    st.markdown('<div class="blue-card">', unsafe_allow_html=True)
    st.subheader("📤 上传文件")
    col1, col2 = st.columns(2)
    with col1:
        word_template = st.file_uploader("上传Word请款单模板", type=["docx"],
                                         help="请上传包含 {{申请人}}、{{合计}}、{{大写}}、{{日期}} 占位符的Word模板")
    with col2:
        excel_data = st.file_uploader("上传需请款专利清单Excel", type=["xlsx"],
                                      help="Excel必须包含 '分割号'、'官费'、'代理费' 列")
    st.markdown('</div>', unsafe_allow_html=True)

    st.info("📋 发票申请表模板已内置在系统中，无需上传")

    st.markdown('<div class="company-selector">', unsafe_allow_html=True)
    st.subheader("🔸 选择命名格式")
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        company_name = st.radio(
            "",
            ["深佳", "集佳"],
            horizontal=True,
            label_visibility="collapsed",
            key="company_selector"
        )
    st.markdown('</div>', unsafe_allow_html=True)

    col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
    with col_btn2:
        if st.button("🚀 生成文件", use_container_width=True, type="primary", key="generate_btn"):
            if not word_template or not excel_data:
                st.error("请上传所有必须的文件！")
                return

            with tempfile.TemporaryDirectory() as temp_dir:
                temp_path = Path(temp_dir)
                output_dir = temp_path / "output"
                output_dir.mkdir(exist_ok=True)

                word_template_path = temp_path / "word_template.docx"
                with open(word_template_path, "wb") as f:
                    f.write(word_template.getbuffer())

                excel_data_path = temp_path / "data.xlsx"
                with open(excel_data_path, "wb") as f:
                    f.write(excel_data.getbuffer())

                try:
                    invoice_template_path = Path(__file__).parent / "发票申请表.xlsx" \
                        if "__file__" in locals() else Path.cwd() / "发票申请表.xlsx"

                    df = pd.read_excel(excel_data_path, dtype=str).fillna("")
                    if "分割号" not in df.columns or "官费" not in df.columns or "代理费" not in df.columns:
                        st.error("Excel 必须包含 '分割号'、'官费'、'代理费' 列")
                        return

                    invoice_rows = []
                    success_count = 0
                    error_count = 0
                    progress_bar = st.progress(0)
                    total_groups = len(df.groupby("分割号"))

                    for i, (split_no, sub) in enumerate(df.groupby("分割号")):
                        try:
                            result = process_split_group(split_no, sub, output_dir,
                                                         word_template_path, company_name)
                            invoice_rows.append(result)
                            success_count += 1
                        except Exception as e:
                            error_count += 1
                            st.warning(f"⚠ 处理分割号 {split_no} 出错：{str(e)}")
                        progress_bar.progress((i + 1) / total_groups)

                    try:
                        excel_filename = generate_invoice_excel(invoice_rows, output_dir,
                                                                invoice_template_path, company_name)
                    except Exception as e:
                        st.error(f"⌛ 生成发票申请表失败：{str(e)}")
                        excel_filename = None

                    if 'generated_files' not in st.session_state:
                        st.session_state.generated_files = {}

                    all_files = {}
                    docx_files = list(output_dir.glob("*.docx"))
                    xlsx_files = list(output_dir.glob("*.xlsx"))
                    for file in docx_files + xlsx_files:
                        with open(file, "rb") as f:
                            all_files[file.name] = f.read()

                    st.session_state.generated_files = all_files
                    st.session_state.company_name = company_name

                    st.markdown('<div class="success-card">', unsafe_allow_html=True)
                    st.success(f"🎉 处理完成：成功生成 {success_count} 个请款单，{error_count} 个失败")
                    st.markdown('</div>', unsafe_allow_html=True)

                except Exception as e:
                    st.error(f"⌛ 处理过程中出现错误：{str(e)}")

    if 'generated_files' in st.session_state and st.session_state.generated_files:
        st.markdown("---")
        st.markdown('<div class="download-section">', unsafe_allow_html=True)
        st.subheader("📥 下载生成的文件")

        col_zip1, col_zip2, col_zip3 = st.columns([1, 2, 1])
        with col_zip2:
            if st.button("📦 一键下载全部", use_container_width=True, type="secondary", key="download_all"):
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    for filename, file_content in st.session_state.generated_files.items():
                        zip_file.writestr(filename, file_content)
                zip_buffer.seek(0)
                company = st.session_state.get('company_name', '公司')
                zip_filename = f"请款单文件_{company}_{date.today().strftime('%Y%m%d')}.zip"
                st.download_button(
                    label="⬇️ 下载ZIP文件",
                    data=zip_buffer,
                    file_name=zip_filename,
                    mime="application/zip",
                    key="download_zip",
                    use_container_width=True
                )

        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            st.write("**📄 请款单文件:**")
            docx_files = {k: v for k, v in st.session_state.generated_files.items() if k.endswith('.docx')}
            if docx_files:
                for filename, file_content in docx_files.items():
                    st.download_button(
                        label=f"下载 {filename}",
                        data=file_content,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"doc_{filename}",
                        use_container_width=True
                    )
            else:
                st.info("暂无请款单文件")

        with col_dl2:
            st.write("**📊 发票申请表:**")
            xlsx_files = {k: v for k, v in st.session_state.generated_files.items() if k.endswith('.xlsx')}
            if xlsx_files:
                for filename, file_content in xlsx_files.items():
                    st.download_button(
                        label=f"下载 {filename}",
                        data=file_content,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"xlsx_{filename}",
                        use_container_width=True
                    )
            else:
                st.info("暂无发票申请表文件")
        st.markdown('</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
