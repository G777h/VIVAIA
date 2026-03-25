import streamlit as st
import pandas as pd
import numpy as np
import requests
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
from PIL import Image as PILImage
import zipfile

# 设置网页布局为宽屏
st.set_page_config(page_title="VIVAIA 门店库存与销量报表生成器", layout="wide")


def process_data(inventory_file, products_file, sales_file):
    # 1. 加载数据，加入编码容错处理
    def load_csv(file):
        try:
            df = pd.read_csv(file, encoding='utf-8-sig', low_memory=False)
        except UnicodeDecodeError:
            file.seek(0)  # 重置文件指针
            df = pd.read_csv(file, encoding='gbk', low_memory=False)
        return df

    inventory_df = load_csv(inventory_file)
    products_df = load_csv(products_file)
    sales_df = load_csv(sales_file)

    # 库存负数清洗
    inventory_df['On hand (current)'] = pd.to_numeric(inventory_df['On hand (current)'], errors='coerce').fillna(0)
    inventory_df['On hand (current)'] = inventory_df['On hand (current)'].clip(lower=0)

    inventory_df['SKC'] = inventory_df['SKU'].astype(str).str[:-3]
    products_df['SKC'] = products_df['Handle'].astype(str).str[:-3]
    sales_df['SKC'] = sales_df['Lineitem sku'].astype(str).str[:-3]

    products_unique = products_df.drop_duplicates(subset=['SKC'], keep='first')[['SKC', 'Image Src']]

    # 销售表预处理
    sales_df['Lineitem quantity'] = pd.to_numeric(sales_df['Lineitem quantity'], errors='coerce').fillna(0)
    sales_valid = sales_df.dropna(subset=['Paid at']).copy()
    sales_valid['Paid at'] = pd.to_datetime(sales_valid['Paid at'], utc=True)
    now = pd.Timestamp.now(tz='UTC')
    sales_valid['days_ago'] = (now - sales_valid['Paid at']).dt.days

    target_locations = {
        'MELBOURNE': 'VIVAIA MELBOURNE CENTRAL',
        'QVB': 'VIVAIA QVB',
        'BONDI': 'VIVAIA BONDI JUNCTION'
    }

    fixed_sizes = ['EU35', 'EU35.5', 'EU36', 'EU36.5', 'EU37', 'EU37.5', 'EU38', 'EU38.5',
                   'EU39', 'EU39.5', 'EU40', 'EU40.5', 'EU41', 'EU41.5', 'EU42', 'EU42.5',
                   'EU43', 'EU43.5', 'EU44', 'EU44.5', 'EU45', 'EU45.5', 'EU46']

    current_time_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    generated_files = {}

    for file_prefix, loc_name in target_locations.items():
        loc_inv = inventory_df[inventory_df['Location'] == loc_name].copy()
        if loc_inv.empty:
            continue

        split_title = loc_inv['Title'].astype(str).str.split('/', expand=True)
        loc_inv['Category'] = split_title[0] if 0 in split_title.columns else ''
        loc_inv['Collection'] = split_title[1] if 1 in split_title.columns else ''
        loc_inv['color'] = split_title[3] if 3 in split_title.columns else ''
        loc_inv['Size'] = split_title[4] if 4 in split_title.columns else ''
        loc_inv['Size'] = loc_inv['Size'].str.strip()

        # 1. 计算总库存
        grouped_inv = loc_inv.groupby('SKC').agg({
            'Category': 'first',
            'Collection': 'first',
            'color': 'first',
            'On hand (current)': 'sum'
        }).rename(columns={'On hand (current)': 'Stock'}).reset_index()

        # 2. 尺码透视表
        size_pivot = loc_inv.pivot_table(index='SKC', columns='Size', values='On hand (current)', aggfunc='sum',
                                         fill_value=0)
        size_pivot = size_pivot.reindex(columns=fixed_sizes, fill_value=0).reset_index()

        # 3. 联结主表
        merged = pd.merge(grouped_inv, products_unique, on='SKC', how='left')
        merged = pd.merge(merged, size_pivot, on='SKC', how='left')
        merged[fixed_sizes] = merged[fixed_sizes].fillna(0).astype(int)

        # 4. 销量计算
        loc_sales = sales_valid[sales_valid['Location'] == loc_name].copy()
        sales_30 = loc_sales[loc_sales['days_ago'] <= 30].groupby('SKC')['Lineitem quantity'].sum().reset_index(
            name='30_days')
        sales_60 = loc_sales[loc_sales['days_ago'] <= 60].groupby('SKC')['Lineitem quantity'].sum().reset_index(
            name='60_days')
        sales_90 = loc_sales[loc_sales['days_ago'] <= 90].groupby('SKC')['Lineitem quantity'].sum().reset_index(
            name='90_days')

        merged = pd.merge(merged, sales_30, on='SKC', how='left')
        merged = pd.merge(merged, sales_60, on='SKC', how='left')
        merged = pd.merge(merged, sales_90, on='SKC', how='left')
        merged[['30_days', '60_days', '90_days']] = merged[['30_days', '60_days', '90_days']].fillna(0).astype(int)

        merged['Monthly Turnover Rate'] = np.where(merged['Stock'] == 0, 'None',
                                                   (merged['30_days'] / merged['Stock']).round(4))
        merged['Quarterly Turnover Rate'] = np.where(merged['Stock'] == 0, 'None',
                                                     (merged['90_days'] / merged['Stock']).round(4))

        # 5. 构建最终表
        final_main = pd.DataFrame({
            'Category': merged['Category'],
            'Collection': merged['Collection'],
            'SKC': merged['SKC'],
            'color': merged['color'],
            'E列_留空': '',
            'F列_留空': '',
            'Pic': '',
            '30 Days Sales': merged['30_days'],
            '60 Days Sales': merged['60_days'],
            '90 Days Sales': merged['90_days'],
            'Monthly Turnover Rate': merged['Monthly Turnover Rate'],
            'Quarterly Turnover Rate': merged['Quarterly Turnover Rate'],
            'Stock': merged['Stock']
        })
        final_df = pd.concat([final_main, merged[fixed_sizes]], axis=1)

        # 6. 写入内存中的 Excel
        excel_buffer = BytesIO()
        final_df.to_excel(excel_buffer, index=False, engine='openpyxl')
        excel_buffer.seek(0)

        # 7. 植入图片
        wb = load_workbook(excel_buffer)
        ws = wb.active
        ws.column_dimensions['G'].width = 10

        total_images = len(merged['Image Src'])
        progress_text = f"正在为 {loc_name} 下载并植入图片..."
        my_bar = st.progress(0, text=progress_text)

        for idx, url in enumerate(merged['Image Src']):
            row = idx + 2
            ws.row_dimensions[row].height = 60
            if pd.notna(url) and str(url).strip() != '':
                try:
                    response = requests.get(url, timeout=5)
                    if response.status_code == 200:
                        img_data = BytesIO(response.content)
                        pil_img = PILImage.open(img_data)
                        pil_img.thumbnail((70, 70))
                        img_byte_arr = BytesIO()
                        pil_img.save(img_byte_arr, format='PNG')
                        img_byte_arr.seek(0)
                        excel_img = OpenpyxlImage(img_byte_arr)
                        ws.add_image(excel_img, f'G{row}')
                except Exception:
                    pass
            # 更新进度条
            my_bar.progress((idx + 1) / total_images, text=progress_text)

        my_bar.empty()  # 处理完毕后清空进度条

        final_buffer = BytesIO()
        wb.save(final_buffer)
        final_buffer.seek(0)

        filename = f'{file_prefix}_{current_time_str}.xlsx'
        generated_files[filename] = final_buffer

    return generated_files


# ================= UI 布局 =================
st.title("📊 VIVAIA 报表自动生成系统")
st.markdown("---")

# 创建左右两列
col1, col2 = st.columns(2)

with col1:
    st.header("1. 上传数据文件")
    inventory_file = st.file_uploader("📁 上传【库存.csv】", type=['csv'])
    products_file = st.file_uploader("📁 上传【产品.csv】", type=['csv'])
    sales_file = st.file_uploader("📁 上传【销售.csv】", type=['csv'])

    start_button = st.button("🚀 开始生成报表", type="primary", use_container_width=True)

with col2:
    st.header("2. 下载结果文件")
    if start_button:
        if not all([inventory_file, products_file, sales_file]):
            st.warning("⚠️ 请先在左侧上传全部三个 CSV 文件！")
        else:
            with st.spinner("🔄 系统正在处理数据，由于需要下载商品图片，这可能需要几分钟的时间，请耐心等待..."):
                try:
                    # 运行核心逻辑
                    result_files = process_data(inventory_file, products_file, sales_file)
                    st.success("✅ 处理完成！请点击下方按钮下载。")

                    # 生成下载按钮
                    for filename, file_buffer in result_files.items():
                        st.download_button(
                            label=f"⬇️ 下载 {filename}",
                            data=file_buffer,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                except Exception as e:
                    st.error(f"❌ 处理过程中发生错误: {str(e)}")
    else:
        st.info("等待上传文件并开始生成...")