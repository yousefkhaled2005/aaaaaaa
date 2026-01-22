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

# مفتاح الديمو (صحيح وسليم)
LICENSE_KEY = "demo:1769086181672:60be7658030000000080d95114798c23373c11c26b9b2d0022d81ff14e"

# --- دالة الإصلاح السحري: تحميل ملفات التحويل الناقصة ---
def setup_apryse_module():
    # التحقق هل نحن على نظام لينكس (سيرفر ستريم ليت)
    if platform.system() == 'Linux':
        module_path = "Lib"
        
        # لو الملفات مش موجودة، نحملها
        if not os.path.exists(module_path):
            st.info("⚙️ جاري إعداد محرك التحويل لأول مرة (قد يستغرق دقيقة)...")
            
            # رابط الملف الناقص من رسالة الخطأ
            url = "https://www.pdftron.com/downloads/StructuredOutputModuleLinux.tar.gz"
            file_name = "module.tar.gz"
            
            try:
                # 1. التحميل
                response = requests.get(url, stream=True)
                with open(file_name, "wb") as f:
                    for chunk in response.iter_content(chunk_size=1024):
                        if chunk:
                            f.write(chunk)
                
                # 2. فك الضغط
                with tarfile.open(file_name) as tar:
                    tar.extractall(".") # سيفك الضغط وينتج مجلد اسمه Lib
                
                st.success("✅ تم تثبيت المحرك بنجاح!")
                
            except Exception as e:
                st.error(f"فشل تحميل المحرك: {e}")
                return False
        
        # 3. إخبار المكتبة بمكان الملفات
        try:
            # إضافة المسار الحالي والمجلد Lib للبحث
            PDFNet.AddResourceSearchPath(".")
            PDFNet.AddResourceSearchPath("./Lib")
        except Exception as e:
            st.warning(f"تنبيه مسار: {e}")

    return True

# --- الدالة الأساسية لتفعيل المفتاح ---
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
    
    # أولاً: تشغيل دالة الإصلاح وتفعيل المكتبة
    if not setup_apryse_module():
        st.stop()
        
    if not init_apryse():
        st.stop()

    with st.spinner('⏳ جاري التحويل (يتم الآن معالجة التنسيقات المعقدة)...'):
        input_filename = "input.pdf"
        output_filename = "converted.docx"
        
        with open(input_filename, "wb") as f:
            f.write(uploaded_file.getbuffer())

        try:
            # التأكد من تحميل الحزمة (بعد ما حملناها يدوياً فوق)
            if not Convert.IsToWordPackagePresent():
                st.warning("⚠️ حزمة التحويل غير ظاهرة، جاري المحاولة...")

            word_options = WordOutputOptions()
            
            # أمر التحويل
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
            st.error("تفاصيل: تأكد أن الرامات كافية وأن الملف ليس تالفاً.")

# تنظيف
if os.path.exists("input.pdf"): os.remove("input.pdf")
if os.path.exists("converted.docx"): os.remove("converted.docx")
