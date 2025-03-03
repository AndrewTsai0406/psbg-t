import os
import base64
import subprocess
from pathlib import Path
import streamlit as st
from document_processor import DocumentProcessor
from rag_pipleline import RAG
from utils import xlsx_to_pdf

# Make sure *no* st.* calls come before this line.
st.set_page_config(page_title="台達文件規格轉換器", layout="wide")

# Increase the sidebar width and text size
st.markdown(
    """
    <style>
    /* Increase the width of the entire sidebar */
    [data-testid="stSidebar"] > div:first-child {
        width: 400px;  /* Adjust to your desired width */
    }

    /* Increase the font size within the sidebar */
    [data-testid="stSidebar"] * {
        font-size: 1.2rem;  /* Adjust to your desired font size */
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# --- Language Dictionaries ---
CHINESE_TEXT = {
    "page_title": "台達文件規格轉換器",
    "main_title": "台達文件規格轉換器",
    "pdf_comparison_title": "PDF 修改比較器",
    "subtitle": "將客戶規格轉換為電器規格的一站式工具",
    "about_title": "關於此工具",
    "about_body": """
**關於：**
此工具能自動將客戶 PDF/Excel 規格文件
轉換為台達標準電器規格（docx 格式）。

**使用方法：**
請在左側填寫必要資訊並上傳客戶規格 PDF 或 Excel 檔案，然後按下「Start」按鈕。

**注意：** Libreoffice 預覽可能導致文件表格顯示異常，
請將轉換後的 docx 下載並透過 Microsoft Word 開啟。

**開發單位：** Delta Research Center  
""",
    "pdf_about_body": """
**關於：**
此工具能自動將客戶兩份PDF文件進行比較，並生成比較結果圖片。

**使用方法：**
請在左側上傳兩份 PDF 文件，然後按下「Start」按鈕。

**開發單位：** Delta Research Center
""",
    "view_logs": "檢視日誌 (View Logs)",
    "log_output": "Log Output",
    "no_logs": "No logs found. Please run the process first or check if the log file is generated.",
    "sidebar_title": "電器規格預填資訊",
    "start_button_text": "Start",
    "start_button_help": "點擊這裡開始轉換",
    "upload_label": "上傳客戶規格檔案 (PDF / Excel)",
    "page_upload_label_1": "## 上傳舊客戶規格檔案",
    "page_upload_label_2": "## 上傳新客戶規格檔案",
    "spinner_text": "小機器人正在背景努力工作中，請稍候...",
    "done_text": "Done!",
    "customer_spec_title": "客戶規格",
    "no_file_warning": "請上傳客戶規格檔案並點擊「Start」。",
    "converted_title": "(轉換後)電器規格",
    "no_docx_file": "No modified DOCX file found to download.",
    "convert_failed": "轉換失敗: ",
    "converted_ready": "### 轉換後的文件已準備好下載\n請點擊下方按鈕下載修改後的 DOCX 文件：",
    "no_converted_docx_warning": "尚未找到可下載的修改後 DOCX 文件。",
    "no_generated_warning": "轉換後的 docx 文件尚未生成，請檢查 DocumentProcessor 輸出。",
    "no_pdf_warning": "無法顯示轉換後文件 (PDF 未產生)。",
    "info_not_converted": "尚未進行轉換，請上傳檔案並點擊「Start」。",
    "footer": "© 2024 台達文件規格轉換器",
}

ENGLISH_TEXT = {
    "page_title": "Delta Document Spec Converter",
    "main_title": "Delta Document Spec Converter",
    "pdf_comparison_title": "PDF Comparison Tool",
    "subtitle": "A one-stop tool to convert customer specs into electrical specs",
    "about_title": "About This Tool",
    "about_body": """
**About:**
This tool automatically converts customer PDF/Excel specification documents
into Delta’s standard electrical specifications (docx format).

**How to Use:**
Please fill out the required information on the left sidebar, upload the customer specification PDF or Excel file, then click the “Start” button.

**Note:** LibreOffice preview may cause table display issues.
After conversion, please download the docx and open with Microsoft Word.

**Developer:** Delta Research Center  
""",
    "pdf_about_body": """
**About:**
This tool automatically compares two PDF files and generates a comparison result image.

**How to Use:**
Please upload two PDF files on the left, then click the “Start” button.

**Developer:** Delta Research Center
""",
    "view_logs": "View Logs",
    "log_output": "Log Output",
    "no_logs": "No logs found. Please run the process first or check if the log file is generated.",
    "sidebar_title": "Pre-fill Info for Electrical Specs",
    "start_button_text": "Start",
    "start_button_help": "Click here to begin conversion",
    "upload_label": "Upload Customer Spec File (PDF / Excel)",
    "page_upload_label_1": "## Upload Old Customer Spec File",
    "page_upload_label_2": "## Upload New Customer Spec File",
    "spinner_text": "Please wait while the system processes your file...",
    "done_text": "Done!",
    "customer_spec_title": "Customer Specification",
    "no_file_warning": "Please upload a customer spec file and click 'Start'.",
    "converted_title": "(Converted) Electrical Specification",
    "no_docx_file": "No modified DOCX file found to download.",
    "convert_failed": "Conversion failed: ",
    "converted_ready": "### Your converted file is ready\nPlease click the button below to download the modified DOCX:",
    "no_converted_docx_warning": "No modified DOCX file found for download.",
    "no_generated_warning": "Converted docx file not generated. Check DocumentProcessor output.",
    "no_pdf_warning": "Cannot display converted file (PDF not generated).",
    "info_not_converted": "No conversion has been done yet. Please upload and click 'Start'.",
    "footer": "© 2024 Delta Document Spec Converter",
}


@st.cache_resource
def init_rag(log_path):
    rag_pipe = RAG(log_path)
    return rag_pipe


def new_page_pdf_comparision():
    # Button to go back to main page
    if st.button("Go Back", icon="🔙"):
        st.session_state["show_new_page"] = False
        st.rerun()
    st.markdown("Go back to the main page by clicking the 'Go Back' button.")

    # --- Language Selection in Sidebar ---
    language_choice = st.sidebar.selectbox("Language 語言", ["中文", "English"])
    if language_choice == "中文":
        TEXT = CHINESE_TEXT
    else:
        TEXT = ENGLISH_TEXT

    # Main Title
    st.title(TEXT["pdf_comparison_title"])

    # Subtitle and Instructions
    st.write("")

    with st.expander(TEXT["about_title"]):
        st.write(TEXT["pdf_about_body"])

    # Sidebar
    uploaded_file_1 = st.sidebar.file_uploader(
        TEXT["page_upload_label_1"], type=["pdf"], key="file_1"
    )
    uploaded_file_2 = st.sidebar.file_uploader(
        TEXT["page_upload_label_2"], type=["pdf"], key="file_2"
    )

    # Start button
    start_button = st.sidebar.button(
        TEXT["start_button_text"],
        key="start-button",
        help=TEXT["start_button_help"],
        icon="▶️",
    )

    # Ensure the upload directory exists
    upload_dir = "./uploaded_file/"
    pdf_output_path = os.path.join(upload_dir, "diff.pdf")

    if not os.path.exists(upload_dir):
        os.makedirs(upload_dir)

    if start_button and uploaded_file_1 is not None and uploaded_file_2 is not None:
        if uploaded_file_1.name.endswith(".pdf"):
            with open(os.path.join(upload_dir, "舊客戶規格.pdf"), "wb") as f:
                f.write(uploaded_file_1.getbuffer())
        if uploaded_file_2.name.endswith(".pdf"):
            with open(os.path.join(upload_dir, "新客戶規格.pdf"), "wb") as f:
                f.write(uploaded_file_2.getbuffer())

        # Run pdf-diff before.pdf after.pdf > comparison_output.png and display the image
        with st.spinner(TEXT["spinner_text"]):
            # with open(
            #     os.path.join(upload_dir, "comparison_output.png"), "wb"
            # ) as out_file:
            #     subprocess.run(
            #         [
            #             "pdf-diff",
            #             os.path.join(upload_dir, "舊客戶規格.pdf"),
            #             os.path.join(upload_dir, "新客戶規格.pdf"),
            #             "-f",
            #             "png",
            #         ],
            #         stdout=out_file,
            #     )
            # st.image(
            #     os.path.join(upload_dir, "comparison_output.png"),
            #     caption="Comparison Output",
            #     use_container_width=True,
            # )
            subprocess.run(
                [
                    "diff-pdf",
                    "-m",
                    "-g",
                    f"--output-diff={pdf_output_path}",
                    os.path.join(upload_dir, "舊客戶規格.pdf"),
                    os.path.join(upload_dir, "新客戶規格.pdf"),
                ],
            )
            # Provide a download button for the PDF
            if os.path.isfile(pdf_output_path):
                with open(pdf_output_path, "rb") as f:
                    pdf_bytes = f.read()

                st.write(TEXT["converted_ready"])

                st.download_button(
                    label="📥 " + TEXT["start_button_text"] + " PDF",
                    data=pdf_bytes,
                    file_name=os.path.basename(pdf_output_path),
                    mime="application/pdf",
                    help=TEXT["start_button_help"],
                )

        st.success(TEXT["done_text"])

    st.markdown("---")


def main_page():
    """
    This is your original main page code,
    extracted into a separate function for clarity.
    """
    # Center the Delta logo
    st.sidebar.image("./data/delta_logo.png", width=300)
    # --- Language Selection in Sidebar ---
    language_choice = st.sidebar.selectbox("Language 語言", ["中文", "English"])
    if language_choice == "中文":
        TEXT = CHINESE_TEXT
    else:
        TEXT = ENGLISH_TEXT

    # Main Title
    st.title(TEXT["main_title"])

    # Subtitle and Instructions
    st.subheader(TEXT["subtitle"])
    st.write("")

    with st.expander(TEXT["about_title"]):
        st.write(TEXT["about_body"])

    with st.expander(TEXT["view_logs"]):
        log_file_path = "rag_logs.txt"
        try:
            with open(log_file_path, "r", encoding="utf8") as f:
                logs = f.read()
            st.text_area(TEXT["log_output"], logs, height=300)
        except FileNotFoundError:
            st.warning(TEXT["no_logs"])

    # Sidebar

    st.sidebar.markdown(f"## {TEXT['sidebar_title']}")

    # Input fields on sidebar
    date = st.sidebar.date_input("Date:")
    model_no = st.sidebar.text_input("Model no:", "AVE-MUJICA SERIES")
    drawn = st.sidebar.text_input("Drawn by:", "王小明")
    design_ee = st.sidebar.text_input("Design EE:", "劉小華")
    design_me = st.sidebar.text_input("Design ME:", "張小玲")
    document_name = st.sidebar.text_input("Document Name:", "ES-130GBC SERIES")
    rev = st.sidebar.text_input("Revision:", "13")

    file_map = {
        "Asus/65JW Y2A": "./data/NBBU/Asus/65JW Y2A --- 單port多輸出電壓 (Type-C)/ADP65JW X2X SERIES-ES06.docx",
        "HP/100BH": "./data/NBBU/HP/100BH/ADP100BH-SERIES-ES06.docx",
        "SIE/160FR": "./data/NBBU/SIE/160FR --- 雙port單輸出電壓/ADP160FR SERIES-ES.docx",
        "DELL/130GB BA": "./data/NBBU/DELL/130GB BA --- 單port多輸出電壓 (type-c)/ADP130GB B SERIES-ES09.docx",
    }

    # Template options
    drawn_options = list(file_map.keys())
    selected_template = st.sidebar.selectbox("使用參照公版", drawn_options)

    st.sidebar.markdown(f"## {TEXT['upload_label']}")
    uploaded_file = st.sidebar.file_uploader("", type=["pdf", "xlsx"])

    # Custom styling for the "Start" button in the sidebar
    st.sidebar.markdown(
        """
        <style>
        [data-testid="stSidebar"] button {
            background-color: #1976d2 !important;
            color: #ffffff !important;
            font-size: 1.1rem !important;
            padding: 0.6em 1em !important;
            border-radius: 0.5em !important;
            border: none !important;
            cursor: pointer !important;
            margin-top: 1em !important;
            margin-left: auto !important;
            margin-right: auto !important;
        }
        [data-testid="stSidebar"] button:hover {
            background-color: #1565c0 !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    # Start button
    start_button = st.sidebar.button(
        TEXT["start_button_text"],
        key="start-button",
        help=TEXT["start_button_help"],
        icon="▶️",
    )

    # Ensure the upload directory exists
    upload_dir = "./uploaded_file/"
    if not os.path.exists(upload_dir):
        os.makedirs(upload_dir)

    transfer_done = False

    if start_button and uploaded_file is not None:
        # Process upload
        if uploaded_file.name.endswith(".pdf"):
            with open(os.path.join(upload_dir, "客戶規格.pdf"), "wb") as f:
                f.write(uploaded_file.getbuffer())
        elif uploaded_file.name.endswith(".xlsx"):
            with open(os.path.join(upload_dir, "客戶規格.xlsx"), "wb") as f:
                f.write(uploaded_file.getbuffer())
            excel_file = os.path.join(upload_dir, "客戶規格.xlsx")
            pdf_file = os.path.join(upload_dir, "客戶規格.pdf")
            xlsx_to_pdf(excel_file, pdf_file)

        # Clear logs
        with open("rag_logs.txt", "w", encoding="utf8") as f:
            f.write("")

        # Notify the user the system is running
        with st.spinner(TEXT["spinner_text"]):
            rag_pipe = init_rag("rag_logs.txt")
            rag_pipe.init_retriever("./uploaded_file/客戶規格.pdf")

            processor = DocumentProcessor(
                ("./uploaded_file/客戶規格.pdf", file_map[selected_template]),
                date=date.strftime("%Y-%m-%d"),
                model_no=model_no,
                drawn=drawn,
                design_ee=design_ee,
                design_me=design_me,
                document_name=document_name,
                rev=rev,
            )
            processor.RAG = rag_pipe
            processor.process_document()
        st.success(TEXT["done_text"])
        transfer_done = True

    st.markdown("---")

    col1, col2 = st.columns(2)

    with col1:
        st.write(f"### {TEXT['customer_spec_title']}")
        if transfer_done:
            with open("./uploaded_file/客戶規格.pdf", "rb") as f:
                pdf_bytes = f.read()
            base64_pdf_1 = base64.b64encode(pdf_bytes).decode("utf-8")
            pdf_display_1 = f"""
                <iframe 
                    src="data:application/pdf;base64,{base64_pdf_1}" 
                    width="100%" 
                    height="800px" 
                    style="border:none;">
                </iframe>
            """
            st.markdown(pdf_display_1, unsafe_allow_html=True)
        else:
            st.info(TEXT["no_file_warning"])

    with col2:
        st.write(f"### {TEXT['converted_title']}")
        if transfer_done:
            old_path = Path(file_map[selected_template])
            new_root = Path("./modified_data")
            new_path = new_root / old_path.relative_to(old_path.parts[0]).parent
            target_file_path = new_path / (
                old_path.stem + "_modified" + old_path.suffix
            )
            target_file_path = str(target_file_path)

            if os.path.isfile(target_file_path):
                with open(target_file_path, "rb") as docx_file:
                    docx_data = docx_file.read()
            else:
                st.warning(TEXT["no_docx_file"])

            if os.path.isfile(target_file_path):
                pdf_output_path = target_file_path.replace("docx", "pdf")

                def convert_to_pdf_with_libreoffice(docx_path, pdf_path):
                    try:
                        subprocess.run(
                            [
                                "libreoffice",
                                "--headless",
                                "--convert-to",
                                "pdf",
                                docx_path,
                                "--outdir",
                                os.path.dirname(pdf_path),
                            ],
                            check=True,
                        )
                    except subprocess.CalledProcessError as e:
                        st.error(f"{TEXT['convert_failed']}{e}")

                convert_to_pdf_with_libreoffice(target_file_path, pdf_output_path)

                if os.path.isfile(pdf_output_path):
                    with open(pdf_output_path, "rb") as f:
                        pdf_bytes_converted = f.read()
                    base64_pdf_converted = base64.b64encode(pdf_bytes_converted).decode(
                        "utf-8"
                    )
                    pdf_display_converted = f"""
                        <iframe 
                            src="data:application/pdf;base64,{base64_pdf_converted}" 
                            width="100%" 
                            height="800px" 
                            style="border:none;">
                        </iframe>
                    """
                    st.markdown(pdf_display_converted, unsafe_allow_html=True)

                    # Provide a download button for the DOCX
                    if os.path.isfile(target_file_path):
                        with open(target_file_path, "rb") as docx_file:
                            docx_data = docx_file.read()

                        st.write(TEXT["converted_ready"])

                        st.download_button(
                            label="📥 " + TEXT["start_button_text"] + " DOCX",
                            data=docx_data,
                            file_name=os.path.basename(target_file_path),
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            help=TEXT["start_button_help"],
                        )
                    else:
                        st.warning(TEXT["no_converted_docx_warning"])
                else:
                    st.warning(TEXT["no_pdf_warning"])
            else:
                st.warning(TEXT["no_generated_warning"])
        else:
            st.info(TEXT["info_not_converted"])

    st.markdown("---")
    st.write(TEXT["footer"])
    st.image("./data/delta_footer_img.png", width=1600)


def run_app():
    """
    This function checks the session state and decides whether
    to show the main page or the new page.
    """
    # Initialize session state boolean
    if "show_new_page" not in st.session_state:
        st.session_state["show_new_page"] = False

    # Sidebar button to navigate to the new page
    # (Use any icon/emoji you like, e.g. "📝", "➡️", "👀", etc.)
    if st.sidebar.button("Comparing PDFs", icon="📝"):
        st.session_state["show_new_page"] = True

    # Conditionally display either the new page or the main page
    if st.session_state["show_new_page"]:
        new_page_pdf_comparision()
    else:
        main_page()


if __name__ == "__main__":
    run_app()
