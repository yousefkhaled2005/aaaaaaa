import streamlit as st
import os
import site
import requests
import tarfile
import platform

# 1. إعداد الصفحة
st.set_page_config(page_title="Apryse PDF Pro", layout="centered")
st.title("🚀 محول العقود والملفات المعقدة")
st.caption("Powered by Apryse (Solid Documents Engine)")

# 2. استدعاء المكتبة
try:
    from PDFNetPython3 import PDFNet, Convert, WordOutputOptions
except ImportError:
    site.main()
    try:
        from PDFNetPython3 import PDFNet, Convert, WordOutputOptions
    except ImportError:
        st.error("❌ المكتبة غير موجودة.")
        st.stop()

# --- دالة إعداد المحرك (تحميل ملفات اللينكس الناقصة) ---
def setup_apryse_module():
    if platform.system() == 'Linux':
        module_path = "Lib"
        # التحقق من وجود المجلد
        if not os.path.exists(module_path):
            st.info("⚙️ جاري إعداد محرك التحويل (لأول مرة فقط)...")
            url = "https://www.pdftron.com/downloads/StructuredOutputModuleLinux.tar.gz"
            file_name = "module.tar.gz"
            try:
                # تحميل الملف
                response = requests.get(url, stream=True)
                with open(file_name, "wb") as f:
                    for chunk in response.iter_content(chunk_size=1024):
                        if chunk:
                            f.write(chunk)
                # فك الضغط
                with tarfile.open(file_name) as tar:
                    tar.extractall(".")
                st.success("✅ المحرك جاهز!")
            except Exception as e:
                st.error(f"فشل إعداد المحرك: {e}")
                return False
        
        # إضافة المسارات للمكتبة
        try:
            PDFNet.AddResourceSearchPath(".")
            PDFNet.AddResourceSearchPath("./Lib")
        except:
            pass
    return True

# --- دالة التفعيل (بدون مفتاح لتجنب تعارض الإصدارات) ---
def init_apryse():
    try:
        # القوس الفارغ يجبر المكتبة على العمل مهما كان إصدارها
        PDFNet.Initialize()
        return True
    except Exception as e:
        st.error(f"خطأ في التشغيل: {e}")
        return False

# 3. الواجهة
uploaded_file = st.file_uploader("ارفع ملف PDF هنا", type=['pdf'])

if uploaded_file and st.button("تحويل إلى Word"):
    
    # 1. إعداد ملفات النظام
    if not setup_apryse_module():
        st.stop()
        
    # 2. تشغيل المكتبة
    if not init_apryse():
        st.stop()

    with st.spinner('⏳ جاري التحويل (Apryse Engine)...'):
        input_filename = "input.pdf"
        output_filename = "converted.docx"
        
        with open(input_filename, "wb") as f:
            f.write(uploaded_file.getbuffer())

        try:
            word_options = WordOutputOptions()
            Convert.ToWord(input_filename, output_filename, word_options)
            
            st.success("✅ تم التحويل بنجاح!")
            
            with open(output_filename, "rb") as f:
                st.download_button(
                    label="⬇️ تحميل ملف الوورد",
                    data=f,
                    file_name="Converted_Contract.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
        except Exception as e:
            st.error(f"حدث خطأ أثناء التحويل: {e}")

# تنظيف
if os.path.exists("input.pdf"): os.remove("input.pdf")
if os.path.exists("converted.docx"): os.remove("converted.docx")
if os.path.exists("module.tar.gz"): os.remove("module.tar.gz")
