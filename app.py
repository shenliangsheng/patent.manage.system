# patent_billing_generator.py ------ 兼容列数 + 序号列 + 发票类型 + 空值容错

import os, re
import streamlit as st
from pathlib import Path
from datetime import date
import pandas as pd
from docx import Document
from openpyxl import load_workbook
import tempfile

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
    # ------------------ 4-3 合计行（合并单元格 + 右对齐） ------------------
    # 先确定官费、代理费、小计三列的索引
    try:
        off_idx = sub_df.columns.get_loc("官费")
        agt_idx = sub_df.columns.get_loc("代理费")
        sum_idx = agt_idx + 1  # 小计紧跟代理费右侧
    except KeyError:
        off_idx = 0
        agt_idx = 1
        sum_idx = 2  # 兜底

    # 插入新行
    row = tbl.add_row()
    cells = row.cells

    # 合并左侧所有列（从第 0 列到 off_idx-1）
    merge_start = cells[0]
    merge_end = cells[off_idx - 1] if off_idx > 0 else cells[0]
    if merge_start != merge_end:
        merge_start.merge(merge_end)

    # 写入"合计"并右对齐
    merge_start.text = "合计"
    for p in merge_start.paragraphs:
        p.alignment = 2  # WD_ALIGN_PARAGRAPH.RIGHT

    # 填写官费、代理费、小计
    if off_idx < len(cells):
        cells[off_idx].text = str(official_total)
    if agt_idx < len(cells):
        cells[agt_idx].text = str(agent_total)
    if sum_idx < len(cells):
        cells[sum_idx].text = str(grand_total)

    filename = sanitize_filename(f"{applicant}-{grand_total}元-{company_name}-{date.today().strftime('%Y%m%d')}.docx")
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

def generate_invoice_excel(rows: list, output_dir: Path, excel_template_path: Path):
    if not rows:
        print("⚠️ 无数据可汇总")
        return

    if not excel_template_path.exists():
        raise FileNotFoundError("Excel template not found")

    wb = load_workbook(excel_template_path)
    ws = wb.active
    start_row = ws.max_row + 1

    for r in rows:
        # 官费行
        ws[f'B{start_row}'] = "普通发票（电子）"
        ws[f'C{start_row}'] = r["申请人"]
        ws[f'G{start_row}'] = r["总官费"]
        ws[f'H{start_row}'] = r["总官费"]
        ws[f'I{start_row}'] = r["总计"]
        ws[f'Q{start_row}'] = date.today().strftime("%Y年%m月%d日")
        start_row += 1

        # 代理费行
        ws[f'B{start_row}'] = "专用发票（电子）"
        ws[f'C{start_row}'] = r["申请人"]
        ws[f'G{start_row}'] = r["总代理费"]
        ws[f'H{start_row}'] = r["总代理费"]
        ws[f'I{start_row}'] = r["总计"]
        ws[f'Q{start_row}'] = date.today().strftime("%Y年%m月%d日")
        start_row += 1

    excel_filename = f"发票申请表-{date.today().strftime('%Y%m%d')}.xlsx"
    wb.save(output_dir / excel_filename)
    print(f"🉑 发票申请表已生成：{output_dir / excel_filename}")
    return excel_filename

# ------------------ Streamlit 界面 ------------------

def main():
    st.set_page_config(page_title="专利请款单生成器", page_icon="📄", layout="wide")
    st.title("📄 专利请款单生成器")
    
    # 公司选择
    company_name = st.radio("选择公司名称:", ["深佳", "集佳"], horizontal=True)
    
    # 文件上传
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("上传模板文件")
        word_template = st.file_uploader("上传Word请款单模板", type=["docx"])
        excel_template = st.file_uploader("上传Excel发票申请表模板", type=["xlsx"])
    
    with col2:
        st.subheader("上传数据文件")
        excel_data = st.file_uploader("上传需请款专利清单Excel", type=["xlsx"])
    
    if st.button("生成请款单和发票申请表", type="primary"):
        if not all([word_template, excel_template, excel_data]):
            st.error("请上传所有必需的文件！")
            return
        
        # 创建临时目录处理文件
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            output_dir = temp_path / "output"
            output_dir.mkdir(exist_ok=True)
            
            # 保存上传的文件
            word_template_path = temp_path / "word_template.docx"
            with open(word_template_path, "wb") as f:
                f.write(word_template.getbuffer())
            
            excel_template_path = temp_path / "excel_template.xlsx"
            with open(excel_template_path, "wb") as f:
                f.write(excel_template.getbuffer())
            
            excel_data_path = temp_path / "data.xlsx"
            with open(excel_data_path, "wb") as f:
                f.write(excel_data.getbuffer())
            
            try:
                # 读取数据
                df = pd.read_excel(excel_data_path, dtype=str).fillna("")
                
                if "分割号" not in df.columns or "官费" not in df.columns or "代理费" not in df.columns:
                    st.error("Excel 必须包含 '分割号'、'官费'、'代理费' 列")
                    return
                
                invoice_rows = []
                success_count = 0
                
                for split_no, sub in df.groupby("分割号"):
                    try:
                        result = process_split_group(split_no, sub, output_dir, word_template_path, company_name)
                        invoice_rows.append(result)
                        success_count += 1
                        st.success(f"成功处理分割号 {split_no}: {result['文件名']}")
                    except Exception as e:
                        st.warning(f"处理分割号 {split_no} 出错：{e}")
                
                # 生成发票申请表
                try:
                    excel_filename = generate_invoice_excel(invoice_rows, output_dir, excel_template_path)
                    st.success(f"发票申请表已生成: {excel_filename}")
                except Exception as e:
                    st.error(f"生成发票申请表失败：{e}")
                
                # 提供下载
                st.subheader("📥 下载生成的文件")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.write("请款单文件:")
                    for file in output_dir.glob("*.docx"):
                        with open(file, "rb") as f:
                            st.download_button(
                                label=f"下载 {file.name}",
                                data=f,
                                file_name=file.name,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                
                with col2:
                    st.write("发票申请表:")
                    for file in output_dir.glob("*.xlsx"):
                        with open(file, "rb") as f:
                            st.download_button(
                                label=f"下载 {file.name}",
                                data=f,
                                file_name=file.name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                
                st.success(f"处理完成！成功生成 {success_count} 个请款单")
                
            except Exception as e:
                st.error(f"处理过程中出现错误：{e}")

if __name__ == "__main__":
    main()