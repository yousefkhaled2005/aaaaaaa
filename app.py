import streamlit as st
import os
import sys
import site

# 1. إعداد الصفحة
st.set_page_config(page_title="Apryse PDF Converter", layout="centered")
st.title("🚀 محول العقود والملفات المعقدة")
st.caption("Using Apryse Solid Documents Technology")

# 2. محاولة استدعاء مكتبة Apryse (مع معالجة مشاكل المسارات)
try:
    from apryse_sdk import PDFNet, Convert, WordOutputOptions
except ImportError:
    # إعادة تحميل المسارات في بيئة السيرفر
    site.main() 
    try:
        from apryse_sdk import PDFNet, Convert, WordOutputOptions
    except ImportError:
        st.error("جاري تثبيت المحرك... يرجى تحديث الصفحة.")
        st.stop()

# مفتاح الديمو الخاص بك (سيظهر علامة مائية للتجربة)
LICENSE_KEY = "demo:1769086181672:60be7658030000000080d95114798c23373c11c26b9b2d0022d81ff14e"

def init_apryse():
    try:
        PDFNet.Initialize(LICENSE_KEY)
        return True
    except Exception as e:
        st.error(f"خطأ في الترخيص: {e}")
        return False

# 3. واجهة المستخدم
uploaded_file = st.file_uploader("ارفع ملف PDF هنا", type=['pdf'])

if uploaded_file and st.button("تحويل إلى Word"):
    if not init_apryse():
        st.stop()

    with st.spinner('جاري إعادة بناء التنسيق والجداول...'):
        # حفظ الملف المرفوع مؤقتاً
        input_filename = "temp_input.pdf"
        output_filename = "converted_contract.docx"
        
        with open(input_filename, "wb") as f:
            f.write(uploaded_file.getbuffer())

        try:
            # التحقق من وجود أدوات التحويل
            if not Convert.IsToWordPackagePresent():
                st.info("جاري تحميل حزمة التحويل لأول مرة...")
            
            # --- عملية التحويل السحرية ---
            word_options = WordOutputOptions()
            # إعدادات لزيادة الدقة
            word_options.SetSetPaperSize(True) 
            
            Convert.ToWord(input_filename, output_filename, word_options)
            
            st.success("✅ تم التحويل بنجاح!")
            
            # زر التحميل
            with open(output_filename, "rb") as f:
                st.download_button(
                    label="⬇️ تحميل ملف الوورد",
                    data=f,
                    file_name="Converted_Document.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
        except Exception as e:
            st.error(f"حدث خطأ: {e}")
            st.warning("تأكد أن الملف غير محمي بكلمة مرور.")