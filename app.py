import streamlit as st
import docx
import qrcode
from io import BytesIO

# 1. إعدادات الصفحة
st.set_page_config(
    page_title="مكتب المحاماة الذكي",
    page_icon="⚖️",
    layout="centered"
)

# تنسيق المظهر العربي
st.markdown("""
    <style>
    .main-title {
        text-align: right; 
        color: #1E3A8A; 
        font-family: 'Arial', sans-serif;
        margin-bottom: 5px;
    }
    .sub-title {
        text-align: right; 
        color: #4B5563; 
        font-size: 16px;
        margin-bottom: 25px;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 class='main-title'>⚖️ نظام توليد الـ QR كود الذكي</h1>", unsafe_allow_html=True)
st.markdown("<p class='sub-title'>ارفع ملف القضية بصيغة Word (.docx) واضغط على زر التوليد لصنع الـ QR كود فوراً.</p>", unsafe_allow_html=True)
st.markdown("---")

# دالة قراءة النصوص الآمنة تماماً لمنع أي خطأ خارجي
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
        raise ValueError(f"فشل في قراءة ملف الـ Word. تأكد أن الملف غير تالف. التفاصيل: {str(e)}")

# واجهة رفع الملف
uploaded_file = st.file_uploader("اختر ملف الوورد الخاص بالقضية", type=["docx"])

if uploaded_file is not None:
    st.markdown("<br>", unsafe_allow_html=True)
    generate_btn = st.button("⚡ توليد الـ QR كود الآن", use_container_width=True)
    
    if generate_btn:
        try:
            with st.spinner("جاري قراءة الملف وتجهيز الكود..."):
                raw_text = extract_clean_text(uploaded_file)
                
            if not raw_text.strip():
                st.error("⚠️ الملف المرفوع فارغ أو لا يحتوي على نصوص قابلة للقراءة.")
            else:
                # حد أقصى 1000 حرف لعدم تخطي سعة الـ QR نهائياً
                max_limit = 1000
                is_truncated = False
                
                if len(raw_text) > max_limit:
                    qr_content = raw_text[:max_limit] + "\n... [تم اختصار باقي النص]"
                    is_truncated = True
                else:
                    qr_content = raw_text

                # توليد الـ QR Code
                qr_generator = qrcode.QRCode(
                    version=None,
                    error_correction=qrcode.constants.ERROR_CORRECT_L,
                    box_size=10,
                    border=4
                )
                qr_generator.add_data(qr_content)
                qr_generator.make(fit=True)
                
                # لون كحلي فخم
                qr_image = qr_generator.make_image(fill_color="#1E3A8A", back_color="white")
                
                image_buffer = BytesIO()
                qr_image.save(image_buffer, format="PNG")
                binary_image = image_buffer.getvalue()

                # عرض النتيجة
                st.markdown("---")
                st.success("✅ تم توليد الـ QR كود بنجاح!")
                
                col_left, col_center, col_right = st.columns([1, 2, 1])
                with col_center:
                    st.image(binary_image, caption="اسحب الصورة أو امسحها بموبايلك", use_container_width=True)
                    
                    st.download_button(
                        label="💾 تحميل صورة الـ QR بجودة عالية",
                        data=binary_image,
                        file_name=f"QR_{uploaded_file.name.replace('.docx', '')}.png",
                        mime="image/png",
                        use_container_width=True
                    )
                    
                with st.expander("👁️ معاينة النص الذي تم تشفيره داخل الـ QR"):
                    st.text(raw_text)
                    if is_truncated:
                        st.warning(f"تنبيه: تم أخذ أول {max_limit} حرف فقط للحفاظ على جودة وسرعة قراءة الـ QR بالكاميرا.")
                    
        except Exception as error:
            st.error(f"❌ حدث خطأ: {str(error)}")
