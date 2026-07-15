import streamlit as st
import docx
import qrcode
from io import BytesIO

# إعدادات الصفحة بالعربي
st.set_page_config(page_title="مكتب المحاماة الذكي - QR", layout="centered")

st.markdown("""
    <div style="text-align: right; direction: rtl;">
        <h1 style="color: #1E3A8A;">⚖️ نظام توليد الـ QR كود الذكي للمكتب</h1>
        <p style="font-size: 18px; color: #4B5563;">ارفع ملف الـ Word الخاص بالقضية، وهيتم قرايته وعمل الـ QR كود بتاعه في ثانية!</p>
    </div>
""", unsafe_allow_allow_html=True)

# مربع رفع الملف
uploaded_file = st.file_uploader("", type=["docx"])

if uploaded_file is not None:
    try:
        # 1. قراءة محتوى ملف الـ Word
        doc = docx.Document(uploaded_file)
        
        # تجميع النصوص من الفقرات والجداول
        full_text = []
        for para in doc.paragraphs:
            if para.text.strip():
                full_text.append(para.text.strip())
                
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        full_text.append(cell.text.strip())

        # دمج النص بالكامل ليدخل في الـ QR (مع وضع حد أقصى للحروف لضمان سهولة القراءة بالكاميرا)
        final_text = "\n".join(full_text)[:2000]

        if final_text.strip():
            # 2. توليد الـ QR Code
            qr = qrcode.QRCode(
                version=None,
                error_correction=qrcode.constants.ERROR_CORRECT_L,
                box_size=10,
                border=4
            )
            qr.add_data(final_text)
            qr.make(fit=True)
            
            img = qr.make_image(fill_color="black", back_color="white")
            
            # تحويل الصورة لبايتس لعرضها وتحميلها
            buf = BytesIO()
            img.save(buf, format="PNG")
            byte_im = buf.getvalue()
            
            # 3. عرض النتيجة للمستخدم
            st.markdown("---")
            st.markdown("<div style='text-align: right;'><h3>تم توليد الـ QR كود بنجاح:</h3></div>", unsafe_allow_html=True)
            
            # عرض الصورة في النص
            st.image(byte_im, width=250)
            
            # زرار التحميل
            st.download_button(
                label="تحميل صورة الـ QR",
                data=byte_im,
                file_name="case_qr.png",
                mime="image/png"
            )
        else:
            st.error("الملف المرفوع فارغ ولا يحتوي على نصوص!")
            
    except Exception as e:
        st.error(f"حصل خطأ أثناء معالجة الملف: {str(e)}")
