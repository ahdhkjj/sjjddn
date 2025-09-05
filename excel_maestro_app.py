# -*- coding: utf-8 -*-
"""
Excel Data Maestro: An AI-powered Streamlit web application for analyzing and 
editing Excel files using natural language commands, with version control.
Ready for deployment on Streamlit Community Cloud.
"""

import streamlit as st
import pandas as pd
import io
import re
import google.generativeai as genai
import json
import os

# --- Page Configuration ---
st.set_page_config(
    layout="wide",
    page_title="استاد داده اکسل (نسخه هوشمند)",
    page_icon="🤖",
)

# --- Helper Functions ---

def dataframe_to_excel_bytes(df):
    """
    Converts a pandas DataFrame into an in-memory Excel file (bytes).
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

def auto_clean_dataframe(df):
    """
    Performs a standard set of cleaning operations on a DataFrame.
    """
    cleaned_df = df.copy()
    for col in cleaned_df.select_dtypes(include=['object']).columns:
        cleaned_df[col] = cleaned_df[col].str.strip()
    cleaned_df.drop_duplicates(inplace=True)
    return cleaned_df

def get_ai_response(api_key, df, command, proxy_url=None):
    """
    Sends the user's command and dataframe schema to a generative AI model
    to get executable pandas code, an explanation, and the user's intent.
    Uses a proxy if provided.
    """
    try:
        if proxy_url:
            os.environ['https_proxy'] = proxy_url
            os.environ['http_proxy'] = proxy_url
            
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-1.5-flash')
    except Exception as e:
        raise ValueError(f"خطا در تنظیم کلید API: {e}")

    schema = ", ".join(f"'{col}'" for col in df.columns)
    
    prompt = f"""
        شما یک متخصص تحلیل داده با پایتون و کتابخانه pandas هستید. وظیفه شما این است که دستور زبان طبیعی کاربر را تحلیل کرده و یک پاسخ JSON تولید کنید.

        دیتافریم با نام `df_copy` در دسترس است. نام ستون‌ها: [{schema}]
        دستور کاربر: "{command}"

        ابتدا نیت کاربر را تشخیص دهید: آیا او قصد تغییر و ویرایش دیتافریم را دارد (modification) یا قصد پرسیدن سوال و تحلیل داده را دارد (analysis)؟

        سپس یک پاسخ در قالب JSON ارائه دهید که شامل سه کلید باشد:
        1. "intent": نیت کاربر، که باید یکی از این دو مقدار باشد: "modification" یا "analysis".
        2. "code": یک قطعه کد پایتون (pandas) که دستور کاربر را اجرا می‌کند.
           - اگر intent برابر "modification" است، کد باید دیتافریم `df_copy` را مستقیماً تغییر دهد (با `inplace=True` یا `df_copy = ...`).
           - اگر intent برابر "analysis" است، کد باید نتیجه تحلیل را در متغیری به نام `result` ذخیره کند.
        3. "explanation": یک توضیح کوتاه و روان به زبان فارسی در مورد کاری که کد انجام می‌دهد.
        
        اکنون، برای دستور کاربر بالا، پاسخ JSON را تولید کنید.
        """

    try:
        response = model.generate_content(prompt)
        json_response_cleaned = re.search(r'```json\n({.*?})\n```', response.text, re.DOTALL)
        if json_response_cleaned:
            return json.loads(json_response_cleaned.group(1))
        else:
            return json.loads(response.text)
    except Exception as e:
        raise ConnectionError(f"خطا در ارتباط با مدل هوش مصنوعی: {e}. لطفاً کلید API، پراکسی و اتصال اینترنت خود را بررسی کنید.")
    finally:
        if proxy_url and 'https_proxy' in os.environ:
            del os.environ['https_proxy']
        if proxy_url and 'http_proxy' in os.environ:
            del os.environ['http_proxy']


def execute_ai_command(api_key, df, command, proxy_url=None):
    """
    Gets the AI-generated code, determines intent, and executes it safely.
    """
    original_rows = len(df)
    df_copy = df.copy()
    ai_response = get_ai_response(api_key, df, command, proxy_url)
    intent = ai_response.get("intent")
    generated_code = ai_response.get("code")
    explanation = ai_response.get("explanation", "توضیحی ارائه نشد.")

    if not generated_code or not intent:
        raise ValueError("هوش مصنوعی پاسخ معتبری تولید نکرد.")

    local_vars = {'df_copy': df_copy, 'pd': pd, 'result': None}
    try:
        exec(generated_code, globals(), local_vars)
    except Exception as e:
        raise SyntaxError(f"کد تولید شده توسط هوش مصنوعی با خطا مواجه شد: {e}")

    if intent == "modification":
        df_copy = local_vars['df_copy']
        rows_affected = original_rows - len(df_copy)
        answer = f"عملیات با موفقیت انجام شد. {abs(rows_affected)} سطر تغییر کرد. مجموعه داده اکنون {len(df_copy)} سطر دارد."
        return intent, df_copy, explanation, answer
    
    elif intent == "analysis":
        analysis_result = local_vars.get('result')
        if isinstance(analysis_result, (pd.Series, pd.DataFrame)):
            analysis_result = analysis_result.to_string()
        elif isinstance(analysis_result, float):
             analysis_result = f"{analysis_result:,.2f}"
        return intent, str(analysis_result), explanation, None
    else:
        raise ValueError(f"نیت نامشخصی توسط هوش مصنوعی تشخیص داده شد: {intent}")


# --- Session State Initialization ---
if 'history' not in st.session_state:
    st.session_state.history = []
if 'current_index' not in st.session_state:
    st.session_state.current_index = -1

# --- UI Layout ---
st.title("🧙‍♂️ استاد داده اکسل (نسخه هوشمند 🤖)")
st.write("فایل اکسل خود را آپلود کنید و با زبان طبیعی (فارسی یا انگلیسی) دستورات خود را برای تحلیل و ویرایش داده‌ها وارد کنید.")

# --- Sidebar ---
with st.sidebar:
    st.header("۱. تنظیمات")

    # Securely get API key for deployed app, with fallback for local use
    try:
        api_key = st.secrets.get("GOOGLE_API_KEY")
        if not api_key:
             api_key = st.text_input("🔑 کلید Google AI API", type="password", help="کلید API خود را اینجا وارد کنید.")
        else:
            st.success("✅ کلید API با موفقیت از Secrets بارگذاری شد.")
    except Exception:
        api_key = st.text_input("🔑 کلید Google AI API", type="password", help="کلید API خود را اینجا وارد کنید.")

    # Securely get Proxy for deployed app, with fallback for local use
    try:
        proxy_url = st.secrets.get("PROXY_URL")
        if not proxy_url:
            proxy_url = st.text_input("🌐 پراکسی (اختیاری)", placeholder="http://user:pass@host:port")
    except Exception:
        proxy_url = st.text_input("🌐 پراکسی (اختیاری)", placeholder="http://user:pass@host:port")
    
    st.markdown("[دریافت کلید API از Google AI Studio](https://aistudio.google.com/app/apikey)")
    st.markdown("---")
    
    st.header("۲. بارگذاری داده")
    uploaded_file = st.file_uploader("یک فایل اکسل (xlsx. یا xls.) انتخاب کنید", type=['xlsx', 'xls'])

    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
            st.session_state.history = [df.copy()]
            st.session_state.current_index = 0
            st.sidebar.success("فایل با موفقیت بارگذاری شد!")
        except Exception as e:
            st.sidebar.error(f"خطا در بارگذاری فایل: {e}")
            st.session_state.history = []
    
    st.markdown("---")
    st.header("نمونه دستورات")
    st.info("""**برای ویرایش:**\n- `فقط کشور ایران رو نشون بده`\n- `سطرهای خالی رو حذف کن`\n\n**برای تحلیل:**\n- `میانگین فروش چقدره؟`\n- `گران‌ترین محصول کدام است؟`""")

# --- Main Application Logic ---
if st.session_state.history:
    current_df = st.session_state.history[st.session_state.current_index]

    st.header("۳. تحلیل و ویرایش با زبان طبیعی")
    prompt = st.text_area("دستور خود را اینجا وارد کنید:", placeholder="مثلاً: میانگین ستون 'فروش' چقدر است؟", height=100)

    if st.button("🚀 اجرای دستور"):
        if not api_key:
            st.error("لطفاً کلید Google AI API خود را در نوار کناری وارد کنید یا در Secrets تنظیم نمایید.")
        elif not prompt:
            st.warning("لطفاً یک دستور وارد کنید.")
        else:
            with st.spinner("در حال ارتباط با هوش مصنوعی و پردازش دستور شما..."):
                try:
                    intent, result_data, explanation, summary_message = execute_ai_command(api_key, current_df, prompt, proxy_url)
                    with st.container(border=True):
                        st.markdown(f"**💡 توضیح هوش مصنوعی:** {explanation}")
                        if intent == "modification":
                            st.markdown(f"**📈 نتیجه:** {summary_message}")
                            st.session_state.history = st.session_state.history[:st.session_state.current_index + 1]
                            st.session_state.history.append(result_data)
                            st.session_state.current_index += 1
                            st.rerun()
                        elif intent == "analysis":
                            st.markdown(f"**📊 پاسخ تحلیل:**")
                            st.code(result_data, language=None)
                except Exception as e:
                    st.error(f"یک خطا رخ داد: {e}")

    st.header("۴. کنترل‌ها و دانلود")
    cols = st.columns([1.5, 2, 2.5, 2.5])
    if cols[0].button("↩️ بازگشت", use_container_width=True, disabled=(st.session_state.current_index <= 0)):
        st.session_state.current_index -= 1
        st.rerun()
    if cols[1].button("↪️ جلو بردن", use_container_width=True, disabled=(st.session_state.current_index >= len(st.session_state.history) - 1)):
        st.session_state.current_index += 1
        st.rerun()
    cols[2].download_button("💾 دانلود اکسل ویرایش شده", dataframe_to_excel_bytes(current_df), "edited_data.xlsx", use_container_width=True)
    cols[3].download_button("✨ دانلود اکسل پاکسازی شده", dataframe_to_excel_bytes(auto_clean_dataframe(current_df.copy())), "cleaned_data.xlsx", use_container_width=True)

    st.header("۵. پیش‌نمایش داده")
    st.info(f"نمایش نسخه {st.session_state.current_index + 1} از {len(st.session_state.history)}. تعداد سطرها: {len(current_df)}")
    st.dataframe(current_df, height=400, use_container_width=True)
else:
    st.info("برای شروع، لطفاً کلید API خود را وارد کرده و یک فایل اکسل از نوار کناری آپلود کنید!")

