import io
import pandas as pd
import plotly.express as px
import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows
import streamlit as st

st.set_page_config(
    page_title="Meeting Analysis App",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("📊 Meeting Analysis App")
st.markdown(
    "Management and tracking application for sales meeting & meal frequency by"
    " project stage (Division: LT)"
)

STAGES = [
    "Chance",
    "Get Design info",
    "Design In",
    "1st quotation",
    "Approval Sample",
    "Final Quotation",
    "Reliability Approval",
    "Design Win",
    "PO",
]
STAGE_COLS_DIDW = [
    "Chance.1",
    "Get Design info.1",
    "Design In.1",
    "1st quotation.1",
    "Approval Sample.1",
    "Final Quotation.1",
    "Reliability Approval.1",
    "Design Win.1",
    "PO.1",
]


@st.cache_data
def process_data(file_act, file_didw):
  df_act = pd.read_excel(file_act)
  df_didw = pd.read_excel(file_didw)

  df_act.columns = [c.strip(".").strip() for c in df_act.columns]
  df_didw.columns = [c.strip(".").strip() for c in df_didw.columns]

  # Lọc riêng Division LT
  if "Sector" in df_act.columns:
    df_act = df_act[df_act["Sector"].astype(str).str.upper() == "LT"]

  if "Sector" in df_didw.columns:
    df_didw = df_didw[df_didw["Sector"].astype(str).str.upper() == "LT"]
  elif "Div" in df_didw.columns:
    df_didw = df_didw[df_didw["Div"].astype(str).str.upper().str.contains("LT")]

  meeting_types = ["meeting", "meeting(on-line)", "meeting online"]
  meal_types = ["meal"]
  valid_types = meeting_types + meal_types

  df_act_clean = df_act[
      df_act["Type"].astype(str).str.lower().isin(valid_types)
  ].copy()
  df_act_clean["Date"] = pd.to_datetime(
      df_act_clean["Date"], errors="coerce"
  )

  def get_act_category(t):
    t_str = str(t).lower().strip()
    if t_str in meal_types:
      return "Meal"
    return "Meeting"

  df_act_clean["Activity_Category"] = df_act_clean["Type"].apply(
      get_act_category
  )

  for col in STAGE_COLS_DIDW:
    if col in df_didw.columns:
      df_didw[col] = pd.to_datetime(df_didw[col], errors="coerce")

  assigned_stages = []
  didw_pjt_dict = {}
  didw_cust_dict = {}

  for _, didw_row in df_didw.iterrows():
    sales = didw_row.get("Sales Member")
    pjt = didw_row.get("Project Name")
    cust = didw_row.get("Customer Name")

    stages_dates = []
    for s_name, col in zip(STAGES, STAGE_COLS_DIDW):
      d = didw_row.get(col)
      if pd.notna(d):
        stages_dates.append((s_name, pd.to_datetime(d)))

    if not stages_dates:
      continue

    item = {"stages": stages_dates}
    if pd.notna(sales) and pd.notna(pjt):
      didw_pjt_dict[(str(sales).strip(), str(pjt).strip())] = item
    if pd.notna(sales) and pd.notna(cust):
      if (str(sales).strip(), str(cust).strip()) not in didw_cust_dict:
        didw_cust_dict[(str(sales).strip(), str(cust).strip())] = item

  for _, act_row in df_act_clean.iterrows():
    sales = str(act_row.get("Sales Member")).strip()
    cust = str(act_row.get("Customer")).strip()
    pjt = str(act_row.get("Project Name")).strip()
    act_date = act_row["Date"]

    assigned_stage = "Non-Relevant"

    if pd.notna(act_date):
      matched_item = didw_pjt_dict.get((sales, pjt))
      if not matched_item:
        matched_item = didw_cust_dict.get((sales, cust))

      if matched_item:
        dates = matched_item["stages"]
        for i in range(len(dates)):
          s_name, s_date = dates[i]
          next_date = dates[i + 1][1] if i + 1 < len(dates) else None

          if s_date <= act_date:
            if next_date is None or act_date < next_date:
              assigned_stage = s_name
              break

    assigned_stages.append(assigned_stage)

  df_act_clean["Assigned_Stage"] = assigned_stages
  return df_act_clean, df_didw


st.sidebar.header("📁 Upload Raw Data")
uploaded_act = st.sidebar.file_uploader(
    "Upload Activity Detail (.xlsx)", type=["xlsx"]
)
uploaded_didw = st.sidebar.file_uploader("Upload DIDW (.xlsx)", type=["xlsx"])

default_act = (
    uploaded_act
    if uploaded_act is not None
    else "20261106_Activity Detail(Attendee included).xlsx"
)
default_didw = (
    uploaded_didw if uploaded_didw is not None else "DIDW_12T (41).xlsx"
)

try:
  df_processed, df_didw_raw = process_data(default_act, default_didw)

  st.sidebar.markdown("---")
  st.sidebar.header("🔍 Filters")

  sales_list = ["All"] + sorted(
      df_processed["Sales Member"].dropna().unique().tolist()
  )
  selected_sales = st.sidebar.selectbox("Sales Member", sales_list)

  df_temp_region = df_processed.copy()
  if selected_sales != "All":
    df_temp_region = df_temp_region[
        df_temp_region["Sales Member"] == selected_sales
    ]
  region_list = ["All"] + sorted(
      df_temp_region["Region"].dropna().unique().tolist()
  )
  selected_region = st.sidebar.selectbox("Region", region_list)

  df_temp_pjt = df_temp_region.copy()
  if selected_region != "All":
    df_temp_pjt = df_temp_pjt[df_temp_pjt["Region"] == selected_region]
  pjt_list = ["All"] + sorted(
      df_temp_pjt["Project Name"].dropna().unique().tolist()
  )
  selected_pjt = st.sidebar.selectbox("Project Name", pjt_list)

  min_date = df_processed["Date"].min().date()
  max_date = df_processed["Date"].max().date()

  st.sidebar.markdown("---")
  st.sidebar.markdown("📅 **Period**")
  start_date = st.sidebar.date_input("Start Date", value=min_date)
  end_date = st.sidebar.date_input("End Date", value=max_date)

  if "submitted" not in st.session_state:
    st.session_state.submitted = False

  if st.sidebar.button("✅ OK - Apply Report", type="primary"):
    st.session_state.submitted = True

  if st.session_state.submitted:
    df_filtered = df_processed.copy()
    if selected_sales != "All":
      df_filtered = df_filtered[df_filtered["Sales Member"] == selected_sales]
    if selected_region != "All":
      df_filtered = df_filtered[df_filtered["Region"] == selected_region]
    if selected_pjt != "All":
      df_filtered = df_filtered[df_filtered["Project Name"] == selected_pjt]

    if start_date and end_date:
      df_filtered = df_filtered[
          (df_filtered["Date"].dt.date >= start_date)
          & (df_filtered["Date"].dt.date <= end_date)
      ]

    # --- MAIN DASHBOARD ---
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Filtered Activities (LT)", len(df_filtered))
    col2.metric("Participating Sales", df_filtered["Sales Member"].nunique())
    col3.metric("Total Projects", df_filtered["Project Name"].nunique())

    st.markdown("---")

    categories_order = STAGES + ["Non-Relevant"]

    # 1. Global Overview (LT / All Region)
    st.subheader(
        "📈 Global Overview: Meeting & Meal Frequency by Project Stage (LT)"
    )
    df_grouped_global = (
        df_filtered.groupby(["Assigned_Stage", "Activity_Category"])
        .size()
        .unstack(fill_value=0)
    )
    for col_name in ["Meeting", "Meal"]:
      if col_name not in df_grouped_global.columns:
        df_grouped_global[col_name] = 0
    df_grouped_global = df_grouped_global[["Meeting", "Meal"]].reindex(
        categories_order, fill_value=0
    )
    df_global_plot = df_grouped_global.reset_index()

    fig_global = px.bar(
        df_global_plot,
        x="Assigned_Stage",
        y=["Meeting", "Meal"],
        barmode="group",
        text_auto=True,
        labels={
            "Assigned_Stage": "Project Stage",
            "value": "Count",
            "variable": "Activity Type",
        },
        color_discrete_map={"Meeting": "#636EFA", "Meal": "#EF553B"},
    )
    fig_global.update_layout(
        xaxis_title="Project Stage",
        yaxis_title="Count",
        legend_title="Activity Type",
    )
    st.plotly_chart(fig_global, use_container_width=True, key="global_chart")

    # 2. Breakdown by Region
    st.subheader("🌍 Breakdown by Region (LT)")
    regions = sorted(df_filtered["Region"].dropna().unique().tolist())
    region_data_dict = {}
    if regions:
      for reg in regions:
        st.markdown(f"**Region: {reg}**")
        df_reg = df_filtered[df_filtered["Region"] == reg]
        df_reg_grouped = (
            df_reg.groupby(["Assigned_Stage", "Activity_Category"])
            .size()
            .unstack(fill_value=0)
        )
        for col_name in ["Meeting", "Meal"]:
          if col_name not in df_reg_grouped.columns:
            df_reg_grouped[col_name] = 0
        df_reg_grouped = df_reg_grouped[["Meeting", "Meal"]].reindex(
            categories_order, fill_value=0
        )
        df_reg_plot = df_reg_grouped.reset_index()
        region_data_dict[reg] = df_reg_plot

        fig_reg = px.bar(
            df_reg_plot,
            x="Assigned_Stage",
            y=["Meeting", "Meal"],
            barmode="group",
            text_auto=True,
            labels={
                "Assigned_Stage": "Project Stage",
                "value": "Count",
                "variable": "Activity Type",
            },
            color_discrete_map={"Meeting": "#636EFA", "Meal": "#EF553B"},
        )
        fig_reg.update_layout(
            xaxis_title="Project Stage",
            yaxis_title="Count",
            legend_title="Activity Type",
            height=350,
        )
        st.plotly_chart(
            fig_reg, use_container_width=True, key=f"reg_chart_{reg}"
        )
    else:
      st.info("No region data available.")

    # 3. Breakdown by Sales Member
    st.subheader("👤 Breakdown by Sales Member (LT)")
    sales_members = sorted(df_filtered["Sales Member"].dropna().unique().tolist())
    sales_data_dict = {}
    if sales_members:
      for sales_m in sales_members:
        st.markdown(f"**Sales Member: {sales_m}**")
        df_s = df_filtered[df_filtered["Sales Member"] == sales_m]
        df_s_grouped = (
            df_s.groupby(["Assigned_Stage", "Activity_Category"])
            .size()
            .unstack(fill_value=0)
        )
        for col_name in ["Meeting", "Meal"]:
          if col_name not in df_s_grouped.columns:
            df_s_grouped[col_name] = 0
        df_s_grouped = df_s_grouped[["Meeting", "Meal"]].reindex(
            categories_order, fill_value=0
        )
        df_s_plot = df_s_grouped.reset_index()
        sales_data_dict[sales_m] = df_s_plot

        fig_s = px.bar(
            df_s_plot,
            x="Assigned_Stage",
            y=["Meeting", "Meal"],
            barmode="group",
            text_auto=True,
            labels={
                "Assigned_Stage": "Project Stage",
                "value": "Count",
                "variable": "Activity Type",
            },
            color_discrete_map={"Meeting": "#636EFA", "Meal": "#EF553B"},
        )
        fig_s.update_layout(
            xaxis_title="Project Stage",
            yaxis_title="Count",
            legend_title="Activity Type",
            height=320,
        )
        st.plotly_chart(
            fig_s, use_container_width=True, key=f"sales_chart_{sales_m}"
        )
    else:
      st.info("No sales member data available.")

    st.markdown("---")
    st.subheader("📋 Detailed Data Table")
    st.dataframe(df_filtered, use_container_width=True)

    # --- EXCEL EXPORT CHUYÊN NGHIỆP: NHÚNG TOÀN BỘ BIỂU ĐỒ ---
    wb = openpyxl.Workbook()

    # Sheet 1: Global Summary & Chart
    ws_global = wb.active
    ws_global.title = "Global Summary"
    for r in dataframe_to_rows(df_global_plot, index=False, header=True):
      ws_global.append(r)

    chart_global = BarChart()
    chart_global.type = "col"
    chart_global.style = 10
    chart_global.title = "Global Overview: Meeting & Meal by Stage"
    chart_global.y_axis.title = "Count"
    chart_global.x_axis.title = "Project Stage"
    data_ref_g = Reference(
        ws_global, min_col=2, min_row=1, max_row=len(df_global_plot) + 1, max_col=3
    )
    cats_ref_g = Reference(
        ws_global, min_col=1, min_row=2, max_row=len(df_global_plot) + 1
    )
    chart_global.add_data(data_ref_g, titles_from_data=True)
    chart_global.set_categories(cats_ref_g)
    chart_global.height = 14
    chart_global.width = 24
    ws_global.add_chart(chart_global, "E2")

    # Sheet 2: Region Summaries & Charts
    if region_data_dict:
      ws_region = wb.create_sheet(title="Region Breakdown")
      curr_row = 1
      for reg, df_r_plot in region_data_dict.items():
        ws_region.cell(row=curr_row, column=1, value=f"Region: {reg}")
        curr_row += 1
        for r in dataframe_to_rows(df_r_plot, index=False, header=True):
          ws_region.append(r)
          curr_row += 1

        chart_r = BarChart()
        chart_r.type = "col"
        chart_r.style = 10
        chart_r.title = f"Region: {reg}"
        chart_r.y_axis.title = "Count"
        chart_r.x_axis.title = "Project Stage"
        start_r = curr_row - len(df_r_plot) - 1
        data_ref_r = Reference(
            ws_region, min_col=2, min_row=start_r, max_row=curr_row - 1, max_col=3
        )
        cats_ref_r = Reference(
            ws_region, min_col=1, min_row=start_r + 1, max_row=curr_row - 1
        )
        chart_r.add_data(data_ref_r, titles_from_data=True)
        chart_r.set_categories(cats_ref_r)
        chart_r.height = 10
        chart_r.width = 18
        ws_region.add_chart(chart_r, f"E{start_r}")
        curr_row += 3  # Khoảng cách trống giữa các biểu đồ region

    # Sheet 3: Sales Summaries & Charts
    if sales_data_dict:
      ws_sales = wb.create_sheet(title="Sales Breakdown")
      curr_row_s = 1
      for sales_m, df_s_plot in sales_data_dict.items():
        ws_sales.cell(row=curr_row_s, column=1, value=f"Sales Member: {sales_m}")
        curr_row_s += 1
        for r in dataframe_to_rows(df_s_plot, index=False, header=True):
          ws_sales.append(r)
          curr_row_s += 1

        chart_s = BarChart()
        chart_s.type = "col"
        chart_s.style = 10
        chart_s.title = f"Sales: {sales_m}"
        chart_s.y_axis.title = "Count"
        chart_s.x_axis.title = "Project Stage"
        start_s = curr_row_s - len(df_s_plot) - 1
        data_ref_s = Reference(
            ws_sales, min_col=2, min_row=start_s, max_row=curr_row_s - 1, max_col=3
        )
        cats_ref_s = Reference(
            ws_sales, min_col=1, min_row=start_s + 1, max_row=curr_row_s - 1
        )
        chart_s.add_data(data_ref_s, titles_from_data=True)
        chart_s.set_categories(cats_ref_s)
        chart_s.height = 10
        chart_s.width = 18
        ws_sales.add_chart(chart_s, f"E{start_s}")
        curr_row_s += 3

    # Sheet 4: Filtered Details
    ws_details = wb.create_sheet(title="Filtered Details")
    df_details_export = df_filtered.copy()
    if "Date" in df_details_export.columns:
      df_details_export["Date"] = df_details_export["Date"].astype(str)
    for r in dataframe_to_rows(df_details_export, index=False, header=True):
      ws_details.append(r)

    # Sheet 5: Original Raw Data
    ws_raw = wb.create_sheet(title="Original Raw Data")
    df_raw_export = pd.read_excel(default_act)
    if "Date" in df_raw_export.columns:
      df_raw_export["Date"] = df_raw_export["Date"].astype(str)
    for r in dataframe_to_rows(df_raw_export, index=False, header=True):
      ws_raw.append(r)

    output = io.BytesIO()
    wb.save(output)
    excel_data = output.getvalue()

    st.download_button(
        label="📥 Download Complete Report (.xlsx)",
        data=excel_data,
        file_name="Meeting_Meal_Full_Charts_Report.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )
  else:
    st.info(
        "👆 Vui lòng upload file raw data, tùy chỉnh các bộ lọc và bấm nút **✅"
        " OK - Apply Report** để hiển thị toàn bộ báo cáo."
    )

except Exception as e:
  st.warning(
      "⚠️ Vui lòng upload đầy đủ cả 2 file Excel (`Activity Detail` và `DIDW`)"
      f" hoặc kiểm tra lại định dạng file. Lỗi chi tiết: {e}"
  )
