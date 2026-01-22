import streamlit as st
import os
import site

# 1. إعداد الصفحة
st.set_page_config(page_title="Apryse PDF Converter", layout="centered")
st.title("🚀 محول العقود (Apryse Engine)")
st.caption("تحويل احترافي يحافظ على التنسيق والجداول")

# 2. استدعاء المكتبة
try:
    from PDFNetPython3 import PDFNet, Convert, WordOutputOptions
except ImportError:
    site.main()
    try:
        from PDFNetPython3 import PDFNet, Convert, WordOutputOptions
    except ImportError:
        st.error("❌ المكتبة غير موجودة. تأكد من requirements.txt")
        st.stop()

# مفتاح الديمو
LICENSE_KEY = "demo:1769086181672:60be7658030000000080d95114798c23373c11c26b9b2d0022d81ff14e"

def init_apryse():
    try:
        PDFNet.Initialize(LICENSE_KEY)
        return True
    except Exception as e:
        st.error(f"خطأ في الترخيص: {e}")
        return False

# 3. الواجهة
uploaded_file = st.file_uploader("ارفع ملف PDF هنا", type=['pdf'])

if uploaded_file and st.button("تحويل إلى Word"):
    if not init_apryse():
        st.stop()

    with st.spinner('⏳ جاري التحويل...'):
        input_filename = "input.pdf"
        output_filename = "converted.docx"
        
        # حفظ الملف
        with open(input_filename, "wb") as f:
            f.write(uploaded_file.getbuffer())

        try:
            # --- التعديل: استخدام أبسط أمر تحويل بدون إعدادات معقدة ---
            word_options = WordOutputOptions()
            
            # التحويل
            Convert.ToWord(input_filename, output_filename, word_options)
            
            st.success("✅ تم التحويل بنجاح!")
            
            # التحميل
            with open(output_filename, "rb") as f:
                st.download_button(
                    label="⬇️ تحميل ملف الوورد",
                    data=f,
                    file_name="Final_Document.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
        except Exception as e:
            st.error(f"حدث خطأ أثناء التحويل: {e}")
            st.warning("قد يكون الملف محمياً أو يحتوي على خطوط غير مدعومة.")

# تنظيف الملفات
if os.path.exists("input.pdf"): os.remove("input.pdf")
if os.path.exists("converted.docx"): os.remove("converted.docx")
