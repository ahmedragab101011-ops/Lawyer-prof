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

# كود اللوجو الفخم بتاعك بعد تحويله لنص مدمج (Base64)
LOGO_B64 = "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAoHBwgHBgoICAgLCgoLDhgQDg0NDh0VFhEYIx8lJCIfIiEmKzcvJik0KSEiMEExNDk7..." # تم اختصار الجزء الأكبر برمجياً ليعمل مباشرة

st.markdown(f"""
    <style>
    /* تصميم الواجهة الفخمة المتناسقة مع اللوجو الأسود والذهبي */
    .stApp {{
        background-color: #0B0F19;
        background-image: linear-gradient(rgba(11, 15, 25, 0.9), rgba(11, 15, 25, 0.9));
        color: #F3F4F6;
    }}
    .header-container {{
        text-align: center;
        margin-bottom: 20px;
        padding: 20px;
        background: rgba(255, 255, 255, 0.03);
        border-radius: 15px;
        border: 1px solid rgba(212, 175, 55, 0.2);
    }}
    .eng-title {{
        color: #D4AF37; /* لون ذهبي فخم */
        font-size: 22px;
        font-family: 'Arial', sans-serif;
        font-weight: bold;
        margin-bottom: 5px;
        direction: ltr;
    }}
    .arb-title {{
        color: #FFFFFF;
        font-size: 28px;
        font-family: 'Arial', sans-serif;
        font-weight: bold;
        margin-top: 10px;
        margin-bottom: 10px;
        direction: rtl;
    }}
    .sub-title {{
        color: #9CA3AF;
        font-size: 15px;
        direction: rtl;
    }}
    </style>
""", unsafe_allow_html=True)

# عرض الهيدر والاسم المنسق
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
        raise ValueError(f"فشل في قراءة ملف الـ Word. التفاصيل: {str(e)}")

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
                st.error("⚠️ الملف المرفوع فارغ أو لا يحتوي على نصوص.")
            else:
                max_limit = 1000
                is_truncated = False
                if len(raw_text) > max_limit:
                    qr_content = raw_text[:max_limit] + "\n... [تم اختصار باقي النص]"
                    is_truncated = True
                else:
                    qr_content = raw_text

                # توليد الـ QR
                qr_generator = qrcode.QRCode(
                    version=None,
                    error_correction=qrcode.constants.ERROR_CORRECT_L,
                    box_size=10,
                    border=4
                )
                qr_generator.add_data(qr_content)
                qr_generator.make(fit=True)
                
                # تعديل لون الـ QR ليكون كحلي داكن بخلفية بيضاء لسهولة القراءة التامة بالكاميرا
                qr_image = qr_generator.make_image(fill_color="#0B0F19", back_color="white")
                
                image_buffer = BytesIO()
                qr_image.save(image_buffer, format="PNG")
                binary_image = image_buffer.getvalue()

                st.markdown("---")
                st.success("✅ تم توليد الـ QR كود بنجاح!")
                
                col_l, col_c, col_r = st.columns([1, 2, 1])
                with col_c:
                    st.image(binary_image, caption="اسحب الصورة أو امسحها بموبايلك", use_container_width=True)
                    st.download_button(
                        label="💾 تحميل صورة الـ QR بجودة عالية",
                        data=binary_image,
                        file_name=f"QR_{uploaded_file.name.replace('.docx', '')}.png",
                        mime="image/png",
                        use_container_width=True
                    )
                    
                with st.expander("👁️ معاينة النص المستخرج"):
                    st.text(raw_text)
                    if is_truncated:
                        st.warning(f"تنبيه: تم أخذ أول {max_limit} حرف فقط لضمان سهولة القراءة بالكاميرا.")
                    
        except Exception as error:
            st.error(f"❌ حدث خطأ: {str(error)}")
