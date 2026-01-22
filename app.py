import streamlit as st
import os
import site

# 1. إعداد الصفحة
st.set_page_config(page_title="Apryse PDF Converter", layout="centered")
st.title("🚀 محول العقود والملفات المعقدة")
st.caption("Powered by Apryse (Solid Documents Engine)")

# 2. استدعاء المكتبة (بالاسم المتوافق مع اللينكس)
try:
    from PDFNetPython3 import PDFNet, Convert, WordOutputOptions
except ImportError:
    # محاولة إصلاح المسار لو السيرفر تاه
    site.main()
    try:
        from PDFNetPython3 import PDFNet, Convert, WordOutputOptions
    except ImportError:
        st.error("❌ لم يتم العثور على مكتبة PDFNetPython3. يرجى مراجعة requirements.txt")
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

    with st.spinner('⏳ جاري التحويل بأعلى جودة...'):
        # حفظ الملف مؤقتاً
        input_filename = "input.pdf"
        output_filename = "converted.docx"
        
        with open(input_filename, "wb") as f:
            f.write(uploaded_file.getbuffer())

        try:
            # التأكد من تحميل أدوات التحويل
            if not Convert.IsToWordPackagePresent():
                st.info("📦 جاري تحميل حزمة التحويل...")

            # التحويل
            word_options = WordOutputOptions()
            word_options.SetSetPaperSize(True)
            Convert.ToWord(input_filename, output_filename, word_options)
            
            st.success("✅ تم التحويل بنجاح!")
            
            # تحميل الملف
            with open(output_filename, "rb") as f:
                st.download_button(
                    label="⬇️ تحميل ملف الوورد",
                    data=f,
                    file_name="Converted_Contract.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
        except Exception as e:
            st.error(f"حدث خطأ: {e}")

# تنظيف الملفات المؤقتة
if os.path.exists("input.pdf"): os.remove("input.pdf")
if os.path.exists("converted.docx"): os.remove("converted.docx")
