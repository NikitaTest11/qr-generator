import streamlit as st
import pandas as pd
import qrcode
from io import BytesIO
from qrcode.image.pil import PilImage
from PIL import Image
from zipfile import ZipFile

APP_VERSION = "v2.1"

st.set_page_config(page_title=f"QR Code Generator from Excel ({APP_VERSION})", layout="wide")
st.title(f"📌 QR Code Generator from Excel ({APP_VERSION})")

def generate_vcard(name, phone, email):
    return f"""BEGIN:VCARD
VERSION:3.0
FN:{name}
TEL:{phone}
EMAIL:{email}
END:VCARD"""

def generate_qr_image(text, size):
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10,
        border=1,
    )
    qr.add_data(text)
    qr.make(fit=True)
    img = qr.make_image(image_factory=PilImage, fill_color="black", back_color="white").convert("RGB")
    return img.resize((size, size), resample=Image.NEAREST)

def safe_filename(text):
    name = text.replace(" ", "_").replace("|", "_").replace(":", "_")
    return "".join(c for c in name if c.isalnum() or c in "_-")[:50] + ".png"

# UI layout
col1, col2 = st.columns([2, 1])

with col1:
    uploaded_file = st.file_uploader("📁 Upload Excel file (.xlsx)", type=["xlsx"])

    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        st.write(f"✅ Loaded rows: {len(df)}")
        st.dataframe(df.head())

        columns = df.columns.tolist()
        qr_format = st.radio("📦 QR content format", ["Plain text (TXT)", "Contact card (vCard)"])

        if qr_format == "Plain text (TXT)":
            source_columns = st.multiselect("🧩 Select columns to generate QR from", columns)
        else:
            name_col = st.selectbox("🧑 Name", columns)
            phone_col = st.selectbox("📞 Phone", columns)
            email_col = st.selectbox("📧 Email", columns)

        target_column = st.selectbox("🎯 Target column for QR (will be overwritten)", columns)
        tooltip_columns = st.multiselect("💬 Tooltip text (on hover)", columns)
        qr_size = st.slider("📐 QR size (px)", 100, 600, 200, step=10)

        export_format = st.multiselect(
            "📤 What do you want to download?",
            ["Excel file with QR codes", "ZIP archive with QR images"],
            default=["Excel file with QR codes"]
        )

        if st.button("🚀 Generate QR Codes"):
            if qr_format == "Plain text (TXT)" and not source_columns:
                st.error("Please select at least one column to generate QR content.")
            else:
                df[target_column] = None
                qr_images = []
                qr_filenames = []
                tooltips = []

                for _, row in df.iterrows():
                    # QR content
                    if qr_format == "Plain text (TXT)":
                        qr_text = " | ".join(str(row[col]) for col in source_columns if pd.notna(row[col]))
                    else:
                        qr_text = generate_vcard(row.get(name_col, ""), row.get(phone_col, ""), row.get(email_col, ""))

                    tooltip_raw = " | ".join(str(row[col]) for col in tooltip_columns if pd.notna(row[col]))
                    tooltip_text = tooltip_raw.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ').strip()[:250] or "QR Code"

                    img = generate_qr_image(qr_text, qr_size)
                    buffer = BytesIO()
                    img.save(buffer, format="PNG")
                    qr_images.append(buffer.getvalue())
                    tooltips.append(tooltip_text)
                    qr_filenames.append(safe_filename(qr_text))

                output_excel = BytesIO()
                if "Excel file with QR codes" in export_format:
                    with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
                        df.to_excel(writer, index=False, sheet_name="QR Data")
                        workbook = writer.book
                        worksheet = writer.sheets["QR Data"]
                        col_index = df.columns.get_loc(target_column)

                        col_width = qr_size * 0.142857
                        row_height = qr_size * 0.75
                        worksheet.set_column(col_index, col_index, col_width)

                        for row_num, (img_bytes, tooltip_text) in enumerate(zip(qr_images, tooltips)):
                            image_stream = BytesIO(img_bytes)
                            worksheet.set_row(row_num + 1, row_height)
                            worksheet.write(row_num + 1, col_index, "")
                            worksheet.insert_image(row_num + 1, col_index, "qr.png", {
                                'image_data': image_stream,
                                'x_offset': 0,
                                'y_offset': 0,
                                'x_scale': 1,
                                'y_scale': 1,
                                'description': tooltip_text
                            })

                output_zip = BytesIO()
                if "ZIP archive with QR images" in export_format:
                    with ZipFile(output_zip, "w") as zipf:
                        for filename, img_bytes in zip(qr_filenames, qr_images):
                            zipf.writestr(filename, img_bytes)

                st.success(f"✅ QR codes generated successfully ({qr_format}, {qr_size}px)")

                if "Excel file with QR codes" in export_format:
                    st.download_button("📥 Download Excel with QR codes", data=output_excel.getvalue(), file_name="qr_output.xlsx")

                if "ZIP archive with QR images" in export_format:
                    st.download_button("📦 Download ZIP with QR images", data=output_zip.getvalue(), file_name="qr_images.zip")

with col2:
    st.markdown("### 🔍 QR Preview")

    preview_name = "AZN_Support"
    preview_email = "AZNSupport@aznresearch.com"
    preview_phone = "123123123"

    if 'qr_format' in locals():
        if qr_format == "Plain text (TXT)":
            preview_text = f"{preview_name} | {preview_phone} | {preview_email}"
        else:
            preview_text = generate_vcard(preview_name, preview_phone, preview_email)

        preview_img = generate_qr_image(preview_text, qr_size)
        st.image(preview_img, caption="QR Preview", use_container_width=True)
        st.code(preview_text, language="text")
