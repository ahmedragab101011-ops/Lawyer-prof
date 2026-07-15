import streamlit as st
import docx
import qrcode
from io import BytesIO

# 1. إعدادات الصفحة
st.set_page_config(
    page_title="مكتب المستشار / احمد رجب",
    page_icon="⚖️",
    layout="centered"
)

# التنسيق النظيف بدون صور
st.markdown("""
    <style>
    .header-container {
        text-align: center;
        margin-bottom: 25px;
        padding: 20px;
        background-color: #F8FAFC;
        border-radius: 12px;
        border: 1px solid #E2E8F0;
    }
    .eng-title {
        color: #4B5563;
        font-size: 22px;
        font-family: 'Arial', sans-serif;
        font-weight: bold;
        margin: 0;
        padding: 0;
        direction: ltr;
    }
    .arb-title {
        color: #1E3A8A;
        font-size: 26px;
        font-family: 'Arial', sans-serif;
        font-weight: bold;
        margin-top: 10px;
        margin-bottom: 10px;
        white-space: nowrap;
        direction: rtl;
    }
    .sub-title {
        color: #6B7280;
        font-size: 15px;
        direction: rtl;
    }
    </style>
""", unsafe_allow_html=True)

# عرض الواجهة
st.markdown("""
    <div class="header-container">
        <div class="eng-title">Qrcode : lawyer-prof</div>
        <div class="arb-title">⚖️ مكتب المستشار / احمد رجب ⚖️</div>
        <div class="sub-title">نظام توليد الـ QR كود الذكي للقضايا والعرائض</div>
    </div>
""", unsafe_allow_html=True)

st.markdown("---")

# دالة قراءة ملف الوورد
def extract_clean_text(file):
    try:
        doc = docx.Document(file)
        extracted_elements = []
        for paragraph in doc.paragraphs:
            cleaned_text = paragraph.text.strip()
            if cleaned_text:
                extracted_elements.append(cleaned_text)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    cleaned_cell = cell.text.strip()
                    if cleaned_cell and cleaned_cell not in extracted_elements:
                        extracted_elements.append(cleaned_cell)
        return "\n".join(extracted_elements)
    except Exception as e:
        raise ValueError(f"فشل في قراءة ملف الـ Word: {str(e)}")

# واجهة رفع الملف
uploaded_file = st.file_uploader("اختر ملف الوورد الخاص بالقضية (.docx)", type=["docx"])

if uploaded_file is not None:
    st.markdown("<br>", unsafe_allow_html=True)
    generate_btn = st.button("⚡ توليد الـ QR كود الآن", use_container_width=True)
    
    if generate_btn:
        try:
            with st.spinner("جاري المعالجة..."):
                raw_text = extract_clean_text(uploaded_file)
                
            if not raw_text.strip():
                st.error("⚠️ الملف المرفوع فارغ.")
            else:
                max_limit = 1000
                qr_content = raw_text[:max_limit] + ("\n... [تم اختصار النص]" if len(raw_text) > max_limit else "")

                # توليد الـ QR
                qr_generator = qrcode.QRCode(
                    version=None,
                    error_correction=qrcode.constants.ERROR_CORRECT_L,
                    box_size=10,
                    border=4
                )
                qr_generator.add_data(qr_content)
                qr_generator.make(fit=True)
                
                qr_image = qr_generator.make_image(fill_color="black", back_color="white")
                
                image_buffer = BytesIO()
                qr_image.save(image_buffer, format="PNG")
                binary_image = image_buffer.getvalue()

                st.markdown("---")
                st.success("✅ تم بنجاح!")
                
                col_c = st.columns([1, 2, 1])[1]
                with col_c:
                    st.image(binary_image, use_container_width=True)
                    st.download_button("💾 تحميل الصورة", data=binary_image, file_name="QR_Code.png", mime="image/png", use_container_width=True)
                    
        except Exception as error:
            st.error(f"❌ حدث خطأ: {str(error)}")
