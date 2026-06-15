import streamlit as st
import pandas as pd
from io import BytesIO

# =========================================
# CONFIG
# =========================================

EMPLOYEE_ID_COLUMN = "EMPLOYEE ID"
EMPLOYEE_NAME_COLUMN = "EMPLOYEE NAME"
ELEMENT_COLUMN = "ELEMENT"

EDITABLE_COLUMNS = [
    "Sent to finance on",
    "Old Amt",
    "New Amount",
    "Final amount to be paid"
]

# =========================================
# PAGE CONFIG
# =========================================

st.set_page_config(
    page_title="Employee Finance Editor",
    layout="wide"
)

st.title("Employee Finance Editor")

# =========================================
# FILE UPLOAD
# =========================================

uploaded_file = st.file_uploader(
    "Upload XLSB File",
    type=["xlsb"]
)

# =========================================
# PROCESS FILE
# =========================================

if uploaded_file is not None:

    try:

        # =========================================
        # LOAD FILE
        # =========================================

        df = pd.read_excel(
            uploaded_file,
            engine="pyxlsb"
        )

        # Clean columns
        df.columns = df.columns.str.strip()

        # Fill merged cells
        df = df.ffill()

        # =========================================
        # CONVERT EXCEL SERIAL DATES
        # =========================================

        if "Sent to finance on" in df.columns:

            df["Sent to finance on"] = pd.to_datetime(
                df["Sent to finance on"],
                origin="1899-12-30",
                unit="D",
                errors="coerce"
            ).dt.strftime("%d-%b-%y")

    except Exception as e:

        st.error(f"Error reading file: {e}")
        st.stop()

    # =========================================
    # CHECK REQUIRED COLUMNS
    # =========================================

    required_columns = [
        EMPLOYEE_ID_COLUMN,
        EMPLOYEE_NAME_COLUMN,
        ELEMENT_COLUMN
    ] + EDITABLE_COLUMNS

    missing_columns = [
        col for col in required_columns
        if col not in df.columns
    ]

    if missing_columns:

        st.error("Missing Columns:")

        for col in missing_columns:
            st.write(f"- {col}")

        st.stop()

    # =========================================
    # EMPLOYEE INPUT
    # =========================================

    employee_id = st.text_input(
        "Enter Employee ID"
    )

    # =========================================
    # SEARCH EMPLOYEE
    # =========================================

    if employee_id:

        employee_rows = df[
            df[EMPLOYEE_ID_COLUMN]
            .astype(str)
            .str.strip()
            == employee_id
        ]

        # =========================================
        # EMPLOYEE NOT FOUND
        # =========================================

        if employee_rows.empty:

            st.error("Employee not found.")

        else:

            # =========================================
            # EMPLOYEE DETAILS
            # =========================================

            employee_name = employee_rows.iloc[0][
                EMPLOYEE_NAME_COLUMN
            ]

            st.success("Employee Found")

            st.write(f"### Employee Name: {employee_name}")

            # =========================================
            # SHOW ALL EMPLOYEE ENTRIES
            # =========================================

            st.write("## All Employee Entries")

            display_df = employee_rows.copy()

            display_df = display_df.reset_index(
                drop=True
            )

            st.dataframe(
                display_df,
                use_container_width=True,
                height=400
            )

            # =========================================
            # ELEMENT DROPDOWN
            # =========================================

            st.write("## Select Element To Update")

            elements = employee_rows[
                ELEMENT_COLUMN
            ].astype(str).tolist()

            selected_element = st.selectbox(
                "Select Element",
                elements
            )

            # =========================================
            # FIND SELECTED ELEMENT
            # =========================================

            selected_rows = employee_rows[
                employee_rows[ELEMENT_COLUMN]
                .astype(str)
                .str.strip()
                .str.lower()
                == selected_element.lower()
            ]

            row_index = selected_rows.index[0]

            # =========================================
            # UPDATE SECTION
            # =========================================

            st.write("## Update Values")

            updated_values = {}

            for col in EDITABLE_COLUMNS:

                current_value = df.at[
                    row_index,
                    col
                ]

                updated_values[col] = st.text_input(
                    label=col,
                    value=str(current_value)
                )

            # =========================================
            # CONFIRM BUTTON
            # =========================================

            if st.button("Confirm Changes"):

                try:

                    # =========================================
                    # VALIDATE + UPDATE
                    # =========================================

                    for col, value in updated_values.items():

                        # DATE COLUMN
                        if col == "Sent to finance on":

                            converted_date = pd.to_datetime(
                                value,
                                format="%d-%b-%y",
                                errors="raise"
                            )

                            df.at[row_index, col] = (
                                converted_date.strftime(
                                    "%d-%b-%y"
                                )
                            )

                        # NUMERIC COLUMNS
                        else:

                            cleaned_value = (
                                value.replace(",", "")
                            )

                            converted_value = float(
                                cleaned_value
                            )

                            df.at[
                                row_index,
                                col
                            ] = converted_value

                    # =========================================
                    # EXPORT UPDATED EMPLOYEE DATA
                    # =========================================

                    updated_employee_rows = df[
                        df[EMPLOYEE_ID_COLUMN]
                        .astype(str)
                        .str.strip()
                        == employee_id
                    ]

                    output = BytesIO()

                    with pd.ExcelWriter(
                        output,
                        engine="openpyxl"
                    ) as writer:

                        updated_employee_rows.to_excel(
                            writer,
                            index=False
                        )

                    output.seek(0)

                    st.success(
                        "Changes confirmed successfully."
                    )

                    # =========================================
                    # DOWNLOAD BUTTON
                    # =========================================

                    st.download_button(
                        label="Download Updated Excel File",
                        data=output,
                        file_name=f"{employee_id}.xlsx",
                        mime=(
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    )

                except:

                    st.error(
                        "Invalid input detected. "
                        "Please check date and numeric fields."
                    )