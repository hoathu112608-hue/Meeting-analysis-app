import streamlit as st
import pandas as pd
import plotly.express as px
import io

st.set_page_config(page_title="Sales Meeting Desktop App", layout="wide", initial_sidebar_state="expanded")

st.title("📊 Sales Meeting Frequency Desktop App")
st.markdown("Ứng dụng quản lý và phân tích tần số Meeting theo Stage dự án (Offline Version)")

STAGES = ['Chance', 'Get Design info', 'Design In', '1st quotation', 'Approval Sample', 'Final Quotation', 'Reliability Approval', 'Design Win', 'PO']
STAGE_COLS_DIDW = ['Chance.1', 'Get Design info.1', 'Design In.1', '1st quotation.1', 'Approval Sample.1', 'Final Quotation.1', 'Reliability Approval.1', 'Design Win.1', 'PO.1']

@st.cache_data
def process_data(file_act_path, file_didw_path):
    df_act = pd.read_excel(file_act_path)
    df_didw = pd.read_excel(file_didw_path)
    
    df_act.columns = [c.strip('.').strip() for c in df_act.columns]
    df_didw.columns = [c.strip('.').strip() for c in df_didw.columns]
    
    meeting_types = ['meal', 'meeting', 'meeting(on-line)', 'meeting online']
    df_act_clean = df_act[df_act['Type'].astype(str).str.lower().isin(meeting_types)].copy()
    df_act_clean['Date'] = pd.to_datetime(df_act_clean['Date'], errors='coerce')
    
    for col in STAGE_COLS_DIDW:
        if col in df_didw.columns:
            df_didw[col] = pd.to_datetime(df_didw[col], errors='coerce')
            
    assigned_stages = []
    
    for _, act_row in df_act_clean.iterrows():
        sales = act_row.get('Sales Member')
        cust = act_row.get('Customer')
        pjt = act_row.get('Project Name')
        act_date = act_row['Date']
        
        matched_didw = pd.DataFrame()
        if pd.notna(pjt):
            matched_didw = df_didw[(df_didw['Sales Member'] == sales) & (df_didw['Project Name'] == pjt)]
        if matched_didw.empty and pd.notna(cust):
            matched_didw = df_didw[(df_didw['Sales Member'] == sales) & (df_didw['Customer Name'] == cust)]
            
        assigned_stage = 'Non-Relevant'
        
        if not matched_didw.empty and pd.notna(act_date):
            for _, didw_row in matched_didw.iterrows():
                dates = []
                for name, col in zip(STAGES, STAGE_COLS_DIDW):
                    d = didw_row.get(col)
                    if pd.notna(d):
                        dates.append((name, d))
                
                if not dates:
                    continue
                
                found = False
                for i in range(len(dates)):
                    s_name, s_date = dates[i]
                    next_date = dates[i+1][1] if i+1 < len(dates) else None
                    
                    if s_date <= act_date:
                        if next_date is None or act_date < next_date:
                            assigned_stage = s_name
                            found = True
                            break
                if found:
                    break
                    
        assigned_stages.append(assigned_stage)
        
    df_act_clean['Assigned_Stage'] = assigned_stages
    return df_act_clean, df_didw

st.sidebar.header("📁 Nhập File Raw Data")
file_act = st.sidebar.file_uploader("1. File Activity Detail (.xlsx)", type=["xlsx"])
file_didw = st.sidebar.file_uploader("2. File DIDW (.xlsx)", type=["xlsx"])

if file_act and file_didw:
    df_processed, df_didw = process_data(file_act, file_didw)
    
    st.sidebar.header("🔍 Bộ lọc (Filters)")
    sales_list = ["All"] + sorted(df_processed['Sales Member'].dropna().unique().tolist())
    selected_sales = st.sidebar.selectbox("Sales Member", sales_list)
    
    region_list = ["All"] + sorted(df_processed['Region'].dropna().unique().tolist())
    selected_region = st.sidebar.selectbox("Region", region_list)
    
    pjt_list = ["All"] + sorted(df_processed['Project Name'].dropna().unique().tolist())
    selected_pjt = st.sidebar.selectbox("Project Name", pjt_list)
    
    min_date = df_processed['Date'].min().date()
    max_date = df_processed['Date'].max().date()
    
    st.sidebar.markdown("---")
    st.sidebar.markdown("📅 **Khoảng thời gian (Period)**")
    start_date = st.sidebar.date_input("Từ ngày", value=min_date)
    end_date = st.sidebar.date_input("Đến ngày", value=max_date)
    
    df_filtered = df_processed.copy()
    if selected_sales != "All":
        df_filtered = df_filtered[df_filtered['Sales Member'] == selected_sales]
    if selected_region != "All":
        df_filtered = df_filtered[df_filtered['Region'] == selected_region]
    if selected_pjt != "All":
        df_filtered = df_filtered[df_filtered['Project Name'] == selected_pjt]
        
    if start_date and end_date:
        df_filtered = df_filtered[(df_filtered['Date'].dt.date >= start_date) & (df_filtered['Date'].dt.date <= end_date)]
        
    col1, col2, col3 = st.columns(3)
    col1.metric("Tổng số Meeting đã lọc", len(df_filtered))
    col2.metric("Số lượng Sales tham gia", df_filtered['Sales Member'].nunique())
    col3.metric("Số lượng Dự án", df_filtered['Project Name'].nunique())
    
    st.markdown("---")
    st.subheader("📈 Biểu đồ Tần số Meeting theo Stage dự án")
    
    categories_order = STAGES + ['Non-Relevant']
    stage_counts = df_filtered['Assigned_Stage'].value_counts().reindex(categories_order, fill_value=0).reset_index()
    stage_counts.columns = ['Stage', 'Meeting Count']
    
    fig = px.bar(
        stage_counts, x='Stage', y='Meeting Count', text='Meeting Count',
        color='Stage', color_discrete_map={'Non-Relevant': '#EF553B'}
    )
    fig.update_traces(textposition='outside')
    fig.update_layout(xaxis_title="Project Stage", yaxis_title="Số lượng Meeting", showlegend=False)
    st.plotly_chart(fig, use_container_width=True)
    
    st.subheader("📋 Bảng Dữ liệu Chi tiết")
    st.dataframe(df_filtered, use_container_width=True)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        stage_counts.to_excel(writer, sheet_name='Summary Stage', index=False)
        df_filtered.to_excel(writer, sheet_name='Meeting Details', index=False)
    excel_data = output.getvalue()
    
    st.download_button(
        label="📥 Tải Báo Cáo Excel (.xlsx)",
        data=excel_data,
        file_name="Sales_Meeting_Report_Offline.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("👋 Vui lòng tải lên 2 file raw data (`Activity Detail` và `DIDW`) ở cột bên trái để khởi chạy ứng dụng!")
