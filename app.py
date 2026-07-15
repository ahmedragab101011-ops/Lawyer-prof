import streamlit as st
from io import BytesIO
import qrcode
import docx
from docx.opc.exceptions import PackageNotFoundError

st.set_page_config(page_title="⚖️ QR Generator", page_icon="⚖️", layout="centered")

st.title("⚖️ نظام توليد QR للمكتب")
st.caption("ارفع ملف Word (.docx) وسيتم إنشاء QR Code بجودة عالية.")

def extract_clean_text(file):
    file.seek(0)
    doc = docx.Document(file)
    texts = []
    seen = set()
    for p in doc.paragraphs:
        t = p.text.strip()
        if t:
            texts.append(t)
            seen.add(t)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                t = cell.text.strip()
                if t and t not in seen:
                    seen.add(t)
                    texts.append(t)
    return "\n".join(texts)

uploaded = st.file_uploader("اختر ملف Word", type=["docx"])

if uploaded:
    try:
        with st.spinner("جاري قراءة الملف..."):
            text = extract_clean_text(uploaded)

        if not text:
            st.warning("الملف لا يحتوي على نص.")
            st.stop()

        limit = 1800
        qr_text = text[:limit]
        if len(text) > limit:
            st.info("تم اختصار النص لضمان سهولة قراءة QR.")

        with st.expander("معاينة النص"):
            st.text(text)

        qr = qrcode.QRCode(
            version=None,
            error_correction=qrcode.constants.ERROR_CORRECT_H,
            box_size=14,
            border=6,
        )
        qr.add_data(qr_text)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")

        buf = BytesIO()
        img.save(buf, format="PNG")
        data = buf.getvalue()

        st.success("تم إنشاء QR بنجاح.")
        st.image(data, use_container_width=True)
        st.download_button(
            "تحميل QR",
            data=data,
            file_name=uploaded.name.replace(".docx","")+"_QR.png",
            mime="image/png",
            use_container_width=True,
        )
    except PackageNotFoundError:
        st.error("الملف ليس مستند DOCX صالح.")
    except PermissionError:
        st.error("الملف محمي أو لا يمكن قراءته.")
    except Exception as e:
        st.error(f"حدث خطأ: {e}")
