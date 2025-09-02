import streamlit as st
import pandas as pd
import google.generativeai as genai
import plotly.express as px

# 🔐 Configure Gemini API
genai.configure(api_key="AIzaSyB6znGDpOFUEO7DIcK-kZhkqgl_9OW5pJc")  # Replace with your actual key

# Load the Excel file
file_path = "Regression 2.xlsx"
sheet1 = pd.read_excel(file_path, sheet_name="Sheet1", header=None, engine="openpyxl")
sheet2 = pd.read_excel(file_path, sheet_name="Sheet 2", header=None, engine="openpyxl")
sheet3 = pd.read_excel(file_path, sheet_name="Sheet 3", header=0, engine="openpyxl")
sheet4 = pd.read_excel(file_path, sheet_name="Sheet 4", header=0, engine="openpyxl")
sheet5 = pd.read_excel(file_path, sheet_name="Sheet 5", engine="openpyxl")

# Extract button labels from Sheet1
button_labels = sheet1.iloc[1:, 0].dropna().tolist()

st.markdown("<h2 style='text-align: center;'>📊  Regression Coverage Report </h2>", unsafe_allow_html=True)

# Initialize session state
for key in ["selected_label", "selected_tile", "search_mode", "gemini_mode", "chart_mode"]:
    if key not in st.session_state:
        st.session_state[key] = False

# Sidebar
st.sidebar.title("Explore")
if st.sidebar.button("🧠 Analyse with AI"):
    st.session_state.update(selected_label=None, selected_tile=None, search_mode=False, gemini_mode=True, chart_mode=False)
if st.sidebar.button("🔍 Search Test Case"):
    st.session_state.update(selected_label=None, selected_tile=None, search_mode=True, gemini_mode=False, chart_mode=False)
if st.sidebar.button("📊 Summary - Graphical "):
    st.session_state.update(selected_label=None, selected_tile=None, search_mode=False, gemini_mode=False, chart_mode=True)

st.sidebar.title("Select Options")
for label in button_labels:
    if st.sidebar.button(label):
        st.session_state.update(selected_label=label, selected_tile=None, search_mode=False, gemini_mode=False, chart_mode=False)

# Charts UI

# Initialize chart_mode if not already set
if "chart_mode" not in st.session_state:
    st.session_state.chart_mode = True  # or False based on your logic

if st.session_state.chart_mode:
    st.markdown("<h3 style='font-size:20px;'>Summary View - Graphs</h3>", unsafe_allow_html=True)

    # Identify column groups separated by empty columns
    column_groups = []
    current_group = []

    for col in sheet5.columns:
        if sheet5[col].isnull().all():
            if current_group:
                column_groups.append(current_group)
                current_group = []
        else:
            current_group.append(col)
    if current_group:
        column_groups.append(current_group)

    # Process each group
    for group in column_groups:
        group_df = sheet5[group].dropna()
        non_numeric_cols = group_df.select_dtypes(exclude='number').columns.tolist()
        numeric_cols = group_df.select_dtypes(include='number').columns.tolist()

        if non_numeric_cols and numeric_cols:
            x_axis = non_numeric_cols[0]
            for y_axis in numeric_cols:
                fig = px.bar(
                    group_df,
                    y=x_axis,
                    x=y_axis,
                    orientation='h',
                    title=f"{y_axis} vs {x_axis}",
                    labels={x_axis: x_axis, y_axis: y_axis},
                    hover_data={x_axis: True, y_axis: True},
                    text=y_axis,
                    color=y_axis,
                    color_continuous_scale='Viridis'
                )
                fig.update_traces(textposition='outside')
                fig.update_layout(
                    xaxis_title=y_axis,
                    yaxis_title=x_axis,
                    margin=dict(l=120, r=40, t=60, b=60),
                    height=600,
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    font=dict(size=12)
                )
                st.plotly_chart(fig, use_container_width=True)


# Gemini mode UI
if st.session_state.gemini_mode:
    st.markdown("<h3 style='font-size:20px;'> AI Analysis</h3>", unsafe_allow_html=True)
    user_query = st.text_area("Enter your query based on the Excel data:")

    if user_query:
        if st.button("Get AI Response"):
            with st.spinner("Thinking..."):
                try:
                    # Prepare context from Excel sheets
                    context_data = {
                        "Sheet1": sheet1.to_string(index=False),
                        "Sheet2": sheet2.to_string(index=False),
                        "Sheet3": sheet3.to_string(index=False),
                        "Sheet4": sheet4.to_string(index=False),
                    }

                    # Construct prompt with context
                    prompt = (
                        "You are an AI assistant that answers questions strictly based on the following Excel data.\n\n"
                        "Sheet1:\n" + context_data["Sheet1"] + "\n\n"
                        "Sheet2:\n" + context_data["Sheet2"] + "\n\n"
                        "Sheet3:\n" + context_data["Sheet3"] + "\n\n"
                        "Sheet4:\n" + context_data["Sheet4"] + "\n\n"
                        "User Query: " + user_query + "\n\n"
                        "Please answer only using the data above. Do not use any external knowledge."
                    )

                    model = genai.GenerativeModel("gemini-2.5-flash")
                    response = model.generate_content(prompt)
                    st.markdown("AI Response : ")
                    st.write(response.text)
                except Exception as e:
                    st.error(f"Error: {e}")


# Search mode UI
if st.session_state.search_mode:
    st.markdown("<h3 style='font-size:20px;'>Search Test Case - To Get More Details</h3>", unsafe_allow_html=True)
    search_input = st.text_input("Enter exact keyword to search in Column 1:")
    if search_input:
        col1_as_str = sheet4.iloc[:, 0].astype(str).str.strip().str.lower()
        search_term = search_input.strip().lower()

        matched_rows = sheet4[col1_as_str == search_term]

        if not matched_rows.empty:
            st.write(f"Found {len(matched_rows)} matching row(s):")

            wrap_enabled = st.checkbox("Wrap table content", value=False)

            if wrap_enabled:
                st.markdown("""
                    <style>
                    div[data-testid="stTable"] {
                        font-size: 12px;
                        white-space: normal;
                    }
                    </style>
                    """, unsafe_allow_html=True)
            else:
                st.markdown("""
                    <style>
                    div[data-testid="stTable"] {
                        font-size: 12px;
                        white-space: nowrap;
                    }
                    </style>
                    """, unsafe_allow_html=True)

            st.table(matched_rows)
        else:
            st.warning("No matching test cases found.")

# Display tiles
if st.session_state.selected_label:
    selected_label = st.session_state.selected_label
    st.markdown(f"<h3 style='font-size:25px;'>{selected_label} Items</h3>", unsafe_allow_html=True)

    selected_row_index = sheet1[sheet1[0] == selected_label].index[0]
    tile_items = sheet1.iloc[selected_row_index, 1:].dropna().tolist()

    # Use 3 columns for 3 items per row
    cols = st.columns(3)
    for idx, item in enumerate(tile_items):
        percentage = None
        for row in sheet2.itertuples(index=False):
            row_list = list(row)
            if item in row_list:
                item_index = row_list.index(item)
                if item_index + 1 < len(row_list):
                    percentage = row_list[item_index + 1]
                break

        if percentage is not None:
            if percentage >= 10:
                bg_color = "#d4edda"
            elif percentage >= 5:
                bg_color = "#fff3cd"
            else:
                bg_color = "#f8d7da"
            percent_text = f"{percentage}"
        else:
            bg_color = "#e2e3e5"
            percent_text = "N/A"

        # Create a styled text link
        link_html = f"""
        <div style="background-color:{bg_color}; border-radius:8px; padding:10px; margin:5px 0; height:100px; display:flex; align-items:center; justify-content:center;">
            <a href="?selected_tile={item}" style="text-decoration:none; color:#333; font-size:12px; text-align:center;">
                {item}<br><span style='font-size:12px;'>{percent_text}</span>
            </a>
        </div>
        """
        with cols[idx % 3]:
            st.markdown(link_html, unsafe_allow_html=True)

# Handle query param to update session state
query_params = st.query_params
if "selected_tile" in query_params:
    st.session_state.selected_tile = query_params["selected_tile"]
    st.session_state.search_mode = False
    st.session_state.gemini_mode = False


# Show test cases below tiles
if st.session_state.selected_tile:
    item = st.session_state.selected_tile
    st.markdown("---")
    st.markdown(f"<h3 style='font-size:20px;'>Test Cases for {item}</h3>", unsafe_allow_html=True)

    if item in sheet3.columns:
        test_cases = sheet3[item].dropna().tolist()
        test_case_count = len(test_cases)
        st.write(f"**Total Test Cases:** {test_case_count}")

        for i, test_case in enumerate(test_cases, 1):
            st.markdown(f"- {test_case}")

        # Create a DataFrame for export
        df_export = pd.DataFrame({f"Test Cases for {item}": test_cases})

        # Convert to CSV
        csv = df_export.to_csv(index=False).encode('utf-8')

        # Add download button
        st.download_button(
            label="📥 Export Test Cases to CSV",
            data=csv,
            file_name=f"{item}_test_cases.csv",
            mime="text/csv"
        )
    else:
        st.write("No test cases available for this item.")
