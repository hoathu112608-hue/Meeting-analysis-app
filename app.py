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
    " project stage"
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


# --- BƯỚC QUAN TRỌNG: GẮN @st.cache_data ĐỂ TĂNG TỐC ĐỘ ---
@st.cache_data
def process_data(file_act, file_didw):
  # Đọc file Excel trực tiếp
  df_act = pd.read_excel(file_act)
  df_didw = pd.read_excel(file_didw)

  df_act.columns = [c.strip(".").strip() for c in df_act.columns]
  df_didw.columns = [c.strip(".").strip() for c in df_didw.columns]

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

  for _, act_row in df_act_clean.iterrows():
    sales = act_row.get("Sales Member")
    cust = act_row.get("Customer")
    pjt = act_row.get("Project Name")
    act_date = act_row["Date"]

    matched_didw = pd.DataFrame()
    if pd.notna(pjt):
      matched_didw = df_didw[
          (df_didw["Sales Member"] == sales)
          & (df_didw["Project Name"] == pjt)
      ]
    if matched_didw.empty and pd.notna(cust):
      matched_didw = df_didw[
          (df_didw["Sales Member"] == sales) & (df_didw["Customer Name"] == cust)
      ]

    assigned_stage = "Non-Relevant"

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
          next_date = dates[i + 1][1] if i + 1 < len(dates) else None

          if s_date <= act_date:
            if next_date is None or act_date < next_date:
              assigned_stage = s_name
              found = True
              break
        if found:
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
  # Gọi hàm đã được cache dữ liệu
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
    col1.metric("Total Filtered Activities", len(df_filtered))
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
        x="Stage",
        y=["Meeting", "Meal"],
        barmode="group",
        text_auto=True,
        labels={"value": "Count", "variable": "Activity Type"},
        color_discrete_map={"Meeting": "#636EFA", "Meal": "#EF553B"},
    )
    fig_global.update_layout(
        xaxis_title="Project Stage",
        yaxis_title="Count",
        legend_title="Activity Type",
    )
    st.plotly_chart(fig_global, use_container_width=True)

    # 2. Breakdown by Region
    st.subheader("🌍 Breakdown by Region")
    df_region_grouped = (
        df_filtered.groupby(["Region", "Assigned_Stage", "Activity_Category"])
        .size()
        .reset_index(name="Count")
    )
    if not df_region_grouped.empty:
      fig_region = px.bar(
          df_region_grouped,
          x="Assigned_Stage",
          y="Count",
          color="Activity_Category",
          barmode="group",
          facet_col="Region",
          category_orders={"Assigned_Stage": categories_order},
          text="Count",
      )
      fig_region.update_layout(
          xaxis_title="Project Stage", yaxis_title="Count"
      )
      st.plotly_chart(fig_region, use_container_width=True)
    else:
      st.info("No data available for Region breakdown with current filters.")

    # 3. Breakdown by Sales Member
    st.subheader("👤 Breakdown by Sales Member")
    df_sales_grouped = (
        df_filtered.groupby(
            ["Sales Member", "Assigned_Stage", "Activity_Category"]
        )
        .size()
        .reset_index(name="Count")
    )
    if not df_sales_grouped.empty:
      fig_sales = px.bar(
          df_sales_grouped,
          x="Assigned_Stage",
          y="Count",
          color="Activity_Category",
          barmode="group",
          facet_row="Sales Member",
          category_orders={"Assigned_Stage": categories_order},
          text="Count",
      )
      fig_sales.update_layout(
          xaxis_title="Project Stage", yaxis_title="Count"
      )
      st.plotly_chart(fig_sales, use_container_width=True)
    else:
      st.info(
          "No data available for Sales Member breakdown with current filters."
      )

    st.markdown("---")
    st.subheader("📋 Detailed Data Table")
    st.dataframe(df_filtered, use_container_width=True)

    # --- EXCEL EXPORT ---
    wb = openpyxl.Workbook()

    ws_summary = wb.active
    ws_summary.title = "Summary & Chart"

    summary_export_df = df_global_plot.copy()
    for r in dataframe_to_rows(
        summary_export_df, index=False, header=True
    ):
      ws_summary.append(r)

    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Meeting & Meal Frequency by Project Stage"
    chart.y_axis.title = "Count"
    chart.x_axis.title = "Project Stage"

    data_ref = Reference(
        ws_summary,
        min_col=2,
        min_row=1,
        max_row=len(summary_export_df) + 1,
        max_col=3,
    )
    cats_ref = Reference(
        ws_summary,
        min_col=1,
        min_row=2,
        max_row=len(summary_export_df) + 1,
    )

    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)

    chart.height = 14
    chart.width = 24
    ws_summary.add_chart(chart, "E2")

    ws_details = wb.create_sheet(title="Filtered Details")
    df_details_export = df_filtered.copy()
    if "Date" in df_details_export.columns:
      df_details_export["Date"] = df_details_export["Date"].astype(str)

    for r in dataframe_to_rows(df_details_export, index=False, header=True):
      ws_details.append(r)

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
        file_name="Meeting_Meal_Analysis_Report.xlsx",
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
