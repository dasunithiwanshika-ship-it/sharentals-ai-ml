"""Streamlit Web Application for Toll Processing Automation System.

Supports two modes:
- INITIAL SETUP: Upload LINKT + 365 + Existing Master (when no stored master exists)
- DAILY RUN: Upload LINKT + 365 only (auto-loads stored master from data/current_master.xlsx)
"""

import io
import datetime
import pandas as pd
import streamlit as st

from toll_processor.processor import process_daily_run
from toll_processor.sample_generator import generate_sample_datasets
from toll_processor.master_store import (
    has_stored_master,
    get_master_info,
    list_archive,
    delete_stored_master,
    load_stored_master,
)
from toll_processor.schemas import (
    DEFAULT_ADMIN_FEE,
    DEFAULT_START_DATE_STR,
    SHEET_MASTER_DATA,
    SHEET_NEW_TOLLS,
    SHEET_DUPLICATES,
    SHEET_UNMATCHED,
    SHEET_ALREADY_PAID,
    SHEET_SUMMARY,
    CURRENT_MASTER_PATH,
)

# Set page config
st.set_page_config(
    page_title="Toll Processing Automation",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom CSS for clean, modern look
st.markdown("""
<style>
    .main-title {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1F4E78;
        margin-bottom: 0.2rem;
    }
    .sub-title {
        font-size: 1.05rem;
        color: #555555;
        margin-bottom: 1.5rem;
    }
    .financial-box {
        background-color: #EBF1F5;
        border: 1px solid #B8CCE4;
        border-radius: 8px;
        padding: 16px;
        margin-top: 10px;
        margin-bottom: 20px;
    }
    .formula-line {
        font-size: 1.15rem;
        font-weight: 600;
        color: #1F4E78;
    }
    .master-info-card {
        background-color: #E8F5E9;
        border: 1px solid #81C784;
        border-radius: 8px;
        padding: 16px;
        margin-top: 8px;
        margin-bottom: 8px;
    }
    .setup-card {
        background-color: #FFF3E0;
        border: 1px solid #FFB74D;
        border-radius: 8px;
        padding: 16px;
        margin-top: 8px;
        margin-bottom: 8px;
    }
</style>
""", unsafe_allow_html=True)

# ─── Detect Mode ───────────────────────────────────────────────────
is_daily_mode = has_stored_master()
master_info = get_master_info() if is_daily_mode else None

# App Title & Description
mode_label = "DAILY RUN" if is_daily_mode else "INITIAL SETUP"
st.markdown(f"<div class='main-title'>🚗 Toll Processing Automation System — {mode_label}</div>", unsafe_allow_html=True)
st.markdown(
    f"<div class='sub-title'>Vehicle Rental Operations — Automated LINKT reformatting, 365 driver lookup, duplicate protection & master generation. (Start Date: <b>{DEFAULT_START_DATE_STR}</b>)</div>",
    unsafe_allow_html=True,
)

# ─── Sidebar ───────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Processing Settings")
    st.markdown("---")

    # Master Dataset Status
    st.markdown("### 📁 Master Dataset Status")
    if is_daily_mode and master_info:
        st.success(f"✅ Stored Master Active")
        st.markdown(
            f"- **Rows:** {master_info['row_count']:,}\n"
            f"- **Size:** {master_info['file_size_kb']} KB\n"
            f"- **Updated:** {master_info['last_modified']}"
        )
    else:
        st.warning("⚠️ No stored Master — Initial Setup required")

    st.markdown("---")

    # Archive browser
    st.markdown("### 🗄️ Archive History")
    archives = list_archive()
    if archives:
        for arch in archives[:5]:  # Show latest 5
            st.caption(f"📄 {arch['filename']} ({arch['file_size_kb']} KB)")
        if len(archives) > 5:
            st.caption(f"... and {len(archives) - 5} more")
    else:
        st.caption("No archives yet")

    st.markdown("---")

    # Reset Master
    st.markdown("### ⚠️ Reset Master")
    if st.button("🔄 Reset to Initial Setup Mode", use_container_width=True, type="secondary"):
        st.session_state["confirm_reset"] = True

    if st.session_state.get("confirm_reset", False):
        st.warning("This will archive the current Master and switch back to Initial Setup mode.")
        col_yes, col_no = st.columns(2)
        with col_yes:
            if st.button("Yes, Reset", type="primary", use_container_width=True):
                delete_stored_master()
                st.session_state["confirm_reset"] = False
                st.rerun()
        with col_no:
            if st.button("Cancel", use_container_width=True):
                st.session_state["confirm_reset"] = False
                st.rerun()

    st.markdown("---")

    st.markdown("### 📌 Important Rules")
    st.info(
        f"• **Start Date**: 08/09/2026 onward\n"
        f"• Historical records in `Paste_eToll_Data` are **never modified**.\n"
        f"• Original Master file is **never overwritten**.\n"
        f"• **Admin Fee** ($5.55) is applied to every new matched toll ONLY.\n"
        f"• Duplicates and Already-Paid tolls do not incur extra charges.\n"
        f"• Unmatched tolls are NOT charged — manual review required.\n"
        f"• Ongoing rentals (`Return not set`) are matched from Start Date onward.\n"
        f"• Unmatched & overlapping tolls are routed for review."
    )

    st.markdown("---")
    st.markdown("### 🧪 Quick Demo Mode")
    if st.button("Generate & Load Sample Demo Files", use_container_width=True, type="secondary"):
        with st.spinner("Generating realistic test files..."):
            l_path, b_path, m_path = generate_sample_datasets("sample_data")
            with open(l_path, "rb") as f:
                st.session_state["demo_linkt"] = f.read()
            with open(b_path, "rb") as f:
                st.session_state["demo_bookings"] = f.read()
            with open(m_path, "rb") as f:
                st.session_state["demo_master"] = f.read()
            st.session_state["use_demo_files"] = True
            st.success("Demo files loaded! Click 'Process Toll Data' below.")


# ─── Main Upload Section ──────────────────────────────────────────
st.markdown("---")

if is_daily_mode:
    # DAILY RUN MODE — 2-column upload + master info card
    col_upload1, col_upload2 = st.columns(2)

    with col_upload1:
        st.subheader("Step 1: LINKT Report")
        linkt_file = st.file_uploader(
            "Upload LINKT Toll Excel (.xls / .xlsx)",
            type=["xlsx", "xls", "csv"],
            key="linkt_uploader",
            help="Downloaded toll statement from Australian LINKT system.",
        )
        if "use_demo_files" in st.session_state and not linkt_file:
            linkt_file = io.BytesIO(st.session_state["demo_linkt"])
            linkt_file.name = "sample_linkt_report.xlsx"
            st.caption("✅ Using Demo LINKT Report")

    with col_upload2:
        st.subheader("Step 2: 365 Bookings")
        bookings_file = st.file_uploader(
            "Upload 365 Booking Export (.xls / .xlsx)",
            type=["xlsx", "xls", "csv"],
            key="bookings_uploader",
            help="Export from 365 rental system containing Rego No, Client, Start & Finish dates.",
        )
        if "use_demo_files" in st.session_state and not bookings_file:
            bookings_file = io.BytesIO(st.session_state["demo_bookings"])
            bookings_file.name = "sample_365_bookings.xlsx"
            st.caption("✅ Using Demo 365 Booking Export")

    # Master Info Card
    st.markdown(
        f"""
        <div class="master-info-card">
            <div style="font-size: 1.1rem; font-weight: 600; color: #2E7D32; margin-bottom: 8px;">
                ✅ Step 3: Master Dataset — Auto-loaded
            </div>
            <div style="font-size: 0.95rem; color: #333;">
                <b>Rows:</b> {master_info['row_count']:,} &nbsp;|&nbsp;
                <b>Size:</b> {master_info['file_size_kb']} KB &nbsp;|&nbsp;
                <b>Last Updated:</b> {master_info['last_modified']}
            </div>
            <div style="font-size: 0.82rem; color: #666; margin-top: 4px;">
                Source: <code>data/current_master.xlsx</code> — The system will automatically use this Master. No upload needed.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Optional Master Override
    master_file = None  # Will auto-load from stored master
    with st.expander("🔧 Override Master (optional — for exceptional situations only)"):
        st.warning("Only use this if you need to replace the current Master with a different file.")
        override_file = st.file_uploader(
            "Upload Override Master Dataset (.xlsx)",
            type=["xlsx", "xls", "csv"],
            key="master_override_uploader",
            help="This will replace the auto-loaded Master for this run only.",
        )
        if override_file:
            master_file = override_file
            st.info("⚡ Override Master will be used for this processing run.")

else:
    # INITIAL SETUP MODE — 3-column upload
    st.markdown(
        """
        <div class="setup-card">
            <div style="font-size: 1.1rem; font-weight: 600; color: #E65100;">
                🔧 Initial Setup — Upload all three files to create the first Master Dataset
            </div>
            <div style="font-size: 0.85rem; color: #555; margin-top: 4px;">
                After successful processing, the generated Master will be saved as <code>data/current_master.xlsx</code>
                and used automatically for all future Daily Runs.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col_upload1, col_upload2, col_upload3 = st.columns(3)

    with col_upload1:
        st.subheader("Step 1: LINKT Report")
        linkt_file = st.file_uploader(
            "Upload LINKT Toll Excel (.xls / .xlsx)",
            type=["xlsx", "xls", "csv"],
            key="linkt_uploader",
            help="Downloaded toll statement from Australian LINKT system.",
        )
        if "use_demo_files" in st.session_state and not linkt_file:
            linkt_file = io.BytesIO(st.session_state["demo_linkt"])
            linkt_file.name = "sample_linkt_report.xlsx"
            st.caption("✅ Using Demo LINKT Report")

    with col_upload2:
        st.subheader("Step 2: 365 Bookings")
        bookings_file = st.file_uploader(
            "Upload 365 Booking Export (.xls / .xlsx)",
            type=["xlsx", "xls", "csv"],
            key="bookings_uploader",
            help="Export from 365 rental system containing Rego No, Client, Start & Finish dates.",
        )
        if "use_demo_files" in st.session_state and not bookings_file:
            bookings_file = io.BytesIO(st.session_state["demo_bookings"])
            bookings_file.name = "sample_365_bookings.xlsx"
            st.caption("✅ Using Demo 365 Booking Export")

    with col_upload3:
        st.subheader("Step 3: Master Dataset")
        master_file = st.file_uploader(
            "Upload Existing Master Dataset (.xlsx)",
            type=["xlsx", "xls", "csv"],
            key="master_uploader",
            help="Current Master Dataset with Paste_eToll_Data sheet.",
        )
        if "use_demo_files" in st.session_state and not master_file:
            master_file = io.BytesIO(st.session_state["demo_master"])
            master_file.name = "sample_master_dataset.xlsx"
            st.caption("✅ Using Demo Master Dataset")


# ─── Processing Configuration ─────────────────────────────────────
col_cfg1, col_cfg2, col_cfg3 = st.columns([1.5, 1.5, 2])

with col_cfg1:
    st.subheader("Processing Date")
    date_received_input = st.date_input(
        "Data Received (Import Period):",
        value=datetime.date(2026, 9, 8),
        help="Data Received date assigned to newly imported records.",
    )

with col_cfg2:
    st.subheader("Admin Fee")
    admin_fee_input = st.number_input(
        "Admin Fee ($ per toll):",
        min_value=0.0,
        max_value=100.0,
        value=float(DEFAULT_ADMIN_FEE),
        step=0.05,
        format="%.2f",
        help="Configurable company administrative fee added to each new matched toll.",
    )

with col_cfg3:
    st.write("")
    st.write("")
    process_btn = st.button("🚀 Process Toll Data", type="primary", use_container_width=True)


# ─── Processing Execution ─────────────────────────────────────────
if process_btn:
    # Validation
    if not linkt_file or not bookings_file:
        st.error("⚠️ Please upload both LINKT Report and 365 Booking Export to proceed.")
    elif not is_daily_mode and master_file is None:
        st.error("⚠️ Initial Setup requires all three files: LINKT, 365 Bookings, and Existing Master.")
    else:
        with st.spinner("Processing toll transactions, matching drivers, checking duplicates, and updating Master..."):
            try:
                if hasattr(linkt_file, "seek"):
                    linkt_file.seek(0)
                if hasattr(bookings_file, "seek"):
                    bookings_file.seek(0)
                if master_file and hasattr(master_file, "seek"):
                    master_file.seek(0)

                date_received_str = date_received_input.strftime("%Y-%m-%d")

                result, save_info = process_daily_run(
                    linkt_file=linkt_file,
                    bookings_365_file=bookings_file,
                    master_file=master_file,
                    date_received=date_received_str,
                    admin_fee=admin_fee_input,
                    is_initial_setup=not is_daily_mode,
                )

                st.session_state["process_result"] = result
                st.session_state["save_info"] = save_info
                st.session_state["processed_at"] = datetime.datetime.now()

                # Success messages
                st.success("🎉 Toll processing completed successfully!")

                # Master update confirmation
                prev_rows = save_info.get("previous_row_count", 0)
                new_rows = save_info.get("new_row_count", 0)
                archived_as = save_info.get("archived_as")

                if archived_as:
                    st.info(
                        f"📦 Master updated: **{prev_rows:,} → {new_rows:,} rows** "
                        f"(+{result.summary_stats['new_tolls_count']} new tolls)\n\n"
                        f"📁 Previous Master archived as: `{archived_as}`"
                    )
                else:
                    st.info(
                        f"📦 Master created: **{new_rows:,} rows** "
                        f"({result.summary_stats['new_tolls_count']} new tolls added to {prev_rows:,} historical rows)\n\n"
                        f"💾 Saved to: `data/current_master.xlsx`"
                    )

                # Validation confirmation
                st.caption(f"✅ Validation: {save_info.get('validation', 'Passed')}")

            except Exception as e:
                st.error(f"❌ Error during processing: {str(e)}")
                st.exception(e)


# ─── Display Dashboard & Results ──────────────────────────────────
if "process_result" in st.session_state:
    result = st.session_state["process_result"]
    stats = result.summary_stats
    file_date_suffix = stats["file_date_suffix"]
    updated_master_filename = f"Master_Toll_Updated_{file_date_suffix}.xlsx"

    st.markdown("---")
    st.header("📊 Processing Dashboard")

    # Metric Cards
    m_col1, m_col2, m_col3, m_col4, m_col5 = st.columns(5)
    with m_col1:
        st.metric(label="LINKT Vehicle Trips", value=stats["total_records"])
    with m_col2:
        st.metric(label="New Matched Tolls", value=stats["new_tolls_count"])
    with m_col3:
        st.metric(label="Duplicate Tolls", value=stats["duplicates_count"])
    with m_col4:
        st.metric(label="Already Paid Tolls", value=stats["already_paid_count"])
    with m_col5:
        st.metric(label="Unmatched / Review", value=stats["total_unmatched_all"])

    # Financial Breakdown Box
    st.markdown(
        f"""
        <div class="financial-box">
            <div style="font-size: 0.95rem; color: #555; text-transform: uppercase; font-weight: 600; margin-bottom: 6px;">
                💰 Financial Breakdown (New Matched Chargeable Tolls Only)
            </div>
            <div class="formula-line">
                Toll Amount (${stats['total_toll_amount']:,.2f}) + Admin Fees (${stats['total_admin_fee']:,.2f}) = 
                <span style="color: #2E75B6; font-size: 1.35rem; font-weight: 700;">TOTAL PAYABLE: ${stats['total_customer_amount']:,.2f}</span>
            </div>
            <div style="font-size: 0.85rem; color: #666; margin-top: 4px;">
                * Admin Fee of <b>${stats['admin_fee']:.2f}</b> applied per NEW matched toll only. Duplicate, already-paid, and unmatched records were NOT charged.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Downloads Bar
    st.subheader("📥 Download Generated Files")
    d_col1, d_col2, d_col3, d_col4, d_col5 = st.columns(5)

    with d_col1:
        st.download_button(
            label="📦 Complete Master Workbook",
            data=result.excel_bytes,
            file_name=updated_master_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
            help="Full updated workbook containing all 6 sheets: Master Data, New Tolls, Duplicates, Unmatched, Already Paid, Summary.",
        )
    with d_col2:
        new_buf = io.BytesIO()
        result.new_tolls_df.to_excel(new_buf, index=False)
        st.download_button(
            label="🚗 New Tolls Only",
            data=new_buf.getvalue(),
            file_name=f"New_Tolls_{file_date_suffix}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with d_col3:
        cust_buf = io.BytesIO()
        result.customer_summary_df.to_excel(cust_buf, index=False)
        st.download_button(
            label="👥 Customer Invoicing Totals",
            data=cust_buf.getvalue(),
            file_name=f"Customer_Invoicing_Summary_{file_date_suffix}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with d_col4:
        unmatched_buf = io.BytesIO()
        result.unmatched_df.to_excel(unmatched_buf, index=False)
        st.download_button(
            label="⚠️ Unmatched Records",
            data=unmatched_buf.getvalue(),
            file_name=f"Unmatched_Tolls_{file_date_suffix}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with d_col5:
        dup_buf = io.BytesIO()
        result.duplicates_df.to_excel(dup_buf, index=False)
        st.download_button(
            label="🔁 Duplicates",
            data=dup_buf.getvalue(),
            file_name=f"Duplicate_Tolls_{file_date_suffix}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    # Detailed Interactive Tables
    st.markdown("---")
    tab_master, tab_new, tab_cust, tab_unmatched, tab_dup, tab_paid, tab_summary = st.tabs([
        "📋 Master Dataset",
        "🚗 New Tolls",
        "👥 Customer Invoicing Totals",
        "⚠️ Unmatched / Review",
        "🔁 Duplicates",
        "💳 Already Paid",
        "📊 Summary & Diagnostics",
    ])

    with tab_master:
        st.markdown(f"**Master Dataset (Historical Paste_eToll_Data + Current Batch)** — Total {len(result.master_df)} rows")
        st.dataframe(result.master_df, use_container_width=True, height=400)

    with tab_new:
        st.markdown(f"**Newly Matched & Chargeable Tolls** — Total {len(result.new_tolls_df)} records")
        if not result.new_tolls_df.empty:
            st.dataframe(
                result.new_tolls_df[[
                    "REGO", "Date", "Time", "Driver", "Hire Date", "Return Date",
                    "Toll Amount", "Admin Fee", "TOTAL AMOUNT ", "Details"
                ]],
                use_container_width=True,
                height=400,
            )
        else:
            st.info("No new valid tolls matched in this batch.")

    with tab_cust:
        st.markdown("**Customer Overall Invoicing Totals (For 365 Invoice Reference)**")
        st.caption("Each toll remains an individual transaction in Master Data, while this table calculates the total customer payable amount across all their tolls.")
        if not result.customer_summary_df.empty:
            st.dataframe(
                result.customer_summary_df,
                use_container_width=True,
                height=400,
            )
        else:
            st.info("No customer totals available.")

    with tab_unmatched:
        st.markdown(f"**Unmatched Tolls for Manual Review** — Total {len(result.unmatched_df)} records")
        st.caption("⚠️ These tolls have NOT been charged. No admin fee was applied. Manual review required before invoicing.")
        if not result.unmatched_df.empty:
            st.dataframe(result.unmatched_df, use_container_width=True, height=400)
        else:
            st.success("✨ All tolls were successfully matched with drivers!")

    with tab_dup:
        st.markdown(f"**Duplicate Tolls Detected** — Total {len(result.duplicates_df)} records")
        st.caption("These tolls already exist in the Master Dataset. No additional charge was applied.")
        if not result.duplicates_df.empty:
            st.dataframe(result.duplicates_df, use_container_width=True, height=400)
        else:
            st.info("No duplicate tolls detected.")

    with tab_paid:
        st.markdown(f"**Already Paid Tolls** — Total {len(result.already_paid_df)} records")
        st.caption("These tolls have already been paid/settled. No additional charge was applied.")
        if not result.already_paid_df.empty:
            st.dataframe(result.already_paid_df, use_container_width=True, height=400)
        else:
            st.info("No already-paid tolls encountered.")

    with tab_summary:
        st.markdown("### 📊 System Diagnostics & Detected Column Mappings")
        c_diag1, c_diag2 = st.columns(2)
        with c_diag1:
            st.markdown("**LINKT Column Detection:**")
            st.json(result.detected_linkt_cols)
        with c_diag2:
            st.markdown("**365 Booking Column Detection:**")
            st.json(result.detected_365_cols)
