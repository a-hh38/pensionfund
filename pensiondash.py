import streamlit as st
import pandas as pd
from updater import update_manual, update_latest
from PIL import Image

logo = Image.open("yes.png")   # Change path if needed

st.sidebar.image(
    logo,
    use_container_width=True
)

st.sidebar.markdown("---")
st.set_page_config(
    page_title="Pension Fund Database Updater",
    layout="wide"
)

st.title("Pension Fund Database Updater")

st.write(
    "Update the Pension Funds Master database."
)

mode = st.radio(
    "Select Update Mode",
    [
        "Latest Update",
        "Manual Update"
    ]
)

if mode == "Manual Update":

    col1, col2 = st.columns(2)

    with col1:

        start_date = st.date_input(
            "Start Month",
            value=pd.Timestamp(2025, 4, 1)
        )

    with col2:

        end_date = st.date_input(
            "End Month",
            value=pd.Timestamp.today().replace(day=1)
        )

    if st.button(
        "Run Manual Update",
        type="primary",
        use_container_width=True
    ):

        if start_date > end_date:

            st.error(
                "Start month must be before End month."
            )

        else:

            with st.spinner(
                "Running all scrapers..."
            ):

                try:

                    df = update_manual(
                        pd.Timestamp(start_date),
                        pd.Timestamp(end_date)
                    )

                    st.success(
                        f"Finished successfully.\n\nRows written: {len(df):,}"
                    )

                    st.dataframe(
                        df,
                        use_container_width=True
                    )

                except Exception as e:

                    st.exception(e)

else:

    st.info(
        "This will append only the months that are missing from the master workbook."
    )

    if st.button(
        "Run Latest Update",
        type="primary",
        use_container_width=True
    ):

        with st.spinner(
            "Checking latest month..."
        ):

            try:

                df = update_latest()

                if df is None or df.empty:

                    st.success(
                        "Database is already up to date."
                    )

                else:

                    st.success(
                        f"Update complete.\n\nTotal rows in master: {len(df):,}"
                    )

                    st.dataframe(
                        df.tail(100),
                        use_container_width=True
                    )

            except Exception as e:

                st.exception(e)

st.divider()

st.caption(
    "Supported Funds: ABSL • Axis • DSP • HDFC • ICICI • Kotak • LIC • SBI • Tata • UTI"
)