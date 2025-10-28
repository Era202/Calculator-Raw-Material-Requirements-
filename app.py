import streamlit as st
import subprocess

st.set_page_config(page_title="تشغيل سكربتات متعددة", page_icon="⚙️", layout="centered")

st.title("⚙️ لوحة تشغيل سكربتات متعددة")

st.markdown("اختر السكربتات التي تريد تشغيلها:")

# ✅ اختيارات السكربتات
script1 = st.checkbox("🔹 script1.py")
script2 = st.checkbox("🔹 script2.py")
script3 = st.checkbox("🔹 script3.py")

# 🔄 اختيار طريقة التشغيل
mode = st.radio(
    "طريقة التشغيل:",
    ["🕐 بالتتابع", "🚀 بالتوازي"]
)

# زر التشغيل
if st.button("▶️ تشغيل المحدد"):
    selected_scripts = []
    if script1:
        selected_scripts.append("script1.py")
    if script2:
        selected_scripts.append("script2.py")
    if script3:
        selected_scripts.append("script3.py")

    if not selected_scripts:
        st.warning("⚠️ من فضلك اختر سكربت واحد على الأقل.")
    else:
        st.write(f"سيتم تشغيل {len(selected_scripts)} سكربت:")
        for s in selected_scripts:
            st.write(f"✅ {s}")

        with st.spinner("جارٍ التشغيل..."):
            if mode == "🕐 بالتتابع":
                # تشغيل واحد بعد التاني
                for script in selected_scripts:
                    st.write(f"🔸 جاري تشغيل {script} ...")
                    result = subprocess.run(["python", script], capture_output=True, text=True)
                    st.text(result.stdout)
                st.success("✅ تم تشغيل جميع السكربتات بالتتابع!")

            else:
                # تشغيل كلهم في وقت واحد
                processes = []
                for script in selected_scripts:
                    st.write(f"🚀 تشغيل {script} في الخلفية...")
                    p = subprocess.Popen(["python", script])
                    processes.append(p)
                st.success("✅ تم تشغيل جميع السكربتات بالتوازي!")
