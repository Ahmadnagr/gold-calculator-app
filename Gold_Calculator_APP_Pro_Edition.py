import json
import os
import urllib.parse
from datetime import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font
import streamlit as st

# --- إعدادات الصفحة العامة ---
st.set_page_config(
    page_title="Jewellery Price Calculator © 2026",
    page_icon="👑",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- إدارة حالة التطبيق (Session State) ---
if "language" not in st.session_state:
    st.session_state.language = "en"
if "gold_results" not in st.session_state:
    st.session_state.gold_results = None
if "diamond_results" not in st.session_state:
    st.session_state.diamond_results = None

# --- الثوابت ---
AVAILABLE_THEMES = ["litera", "minty", "pulse", "flatly", "journal", "azure", "forest"]
COUNTRY_CODES = ["+971", "+20", "+966", "+965", "+974", "+973", "+968", "+1", "+44"]
CARAT_OPTIONS = ["21", "18"]
DIVISOR_OPTIONS = [1.5, 1.75, 2.0, 2.25, 2.5, 3.0, 4.0]

# --- قاموس الترجمة الكامل ---
def tr(text):
    translations = {
        "title": "Jewellery Price Calculator",
        "company": "Jewellery Price Calculator for Staff",
        "gold_price_tab": "Gold Calculator",
        "diamond_price_tab": "Diamond Calculator",
        "gold_price": "Gold Price",
        "carat": "Carat:",
        "piece_data": "Piece Data",
        "weight": "Weight (grams):",
        "manufacturing": "Manufacturing (AED/gram):",
        "discount": "Discount (%):",
        "suggested_price": "Suggested Price (AED):", 
        "calculate": "Calculate Price",
        "clear": "Clear Fields",
        "save": "Download JSON",
        "export_excel": "Export to Excel",
        "customer_results": "Customer Results",
        "company_results": "Company Results",
        "piece_cost": "Piece Cost:",
        "price_before_discount": "Price Before Discount:",
        "discount_amount": "Discount Amount:",
        "price_after_discount": "Price After Discount:",
        "vat": "VAT (4.7619%):",
        "final_price": "Final Price:",
        "net_price": "Net Price:",
        "net_profit": "Net Profit:",
        "gross_margin": "Gross Margin:",
        "manufacturing_cost_label": "Manufacturing Cost per Gram:",
        "language": "Language / اللغة",
        "arabic": "العربية",
        "english": "English",
        "error": "Error",
        "success": "Success",
        "positive_values_error": "Values must be positive numbers",
        "calculation_success": "Calculation successful!",
        "input_error": "Input error, please check mathematical expression",
        "no_results_yet": "No calculations made yet",
        "theme": "Theme",
        "tag_price": "Barcode Price (AED):",
        "cost_tag_divided_by": "Cost (Tag divided by):",
        "calculated_cost": "Calculated Cost:",
        "diamond_customer_price": "Customer Price (AED):",
        "proposed_discount": "Proposed Discount (%):",
        "item_cost": "Item Cost:",
        "diamond_tax": "VAT (4.7619%):",
        "diamond_final_price": "Final Price (Customer):",
        "diamond_net_price": "Net Price (Company):",
        "diamond_gross_profit": "Gross Profit (Company):",
        "diamond_gross_margin": "Gross Margin:",
        "calc_tip": "💡 Note: You can type direct expressions like: 100*2.5 or 50+30 then hit Enter"
    }
    
    # قاموس الترجمة للعربية
    translations_ar = {
        "title": "حاسبة أسعار الذهب والمجوهرات",
        "company": "حاسبة أسعار الذهب والمجوهرات للموظفين",
        "gold_price_tab": "حاسبة الذهب",
        "diamond_price_tab": "حاسبة الألماس",
        "gold_price": "سعر الذهب",
        "carat": "العيار:",
        "piece_data": "بيانات القطعة",
        "weight": "الوزن (جرام):",
        "manufacturing": "المصنعية (درهم/جرام):",
        "discount": "الخصم (%):",
        "suggested_price": "السعر المقترح (درهم):", 
        "calculate": "حساب السعر",
        "clear": "مسح الحقول",
        "save": "تحميل ملف JSON",
        "export_excel": "تصدير لإكسل",
        "customer_results": "نتائج العميل",
        "company_results": "نتائج الشركة",
        "piece_cost": "تكلفة القطعة:",
        "price_before_discount": "السعر قبل الخصم:",
        "discount_amount": "قيمة الخصم:",
        "price_after_discount": "السعر بعد الخصم:",
        "vat": "الضريبة (4.7619%):",
        "final_price": "السعر النهائي:",
        "net_price": "صافي السعر:",
        "net_profit": "صافي الربح:",
        "gross_margin": "هامش الربح:",
        "manufacturing_cost_label": "المصنعية لكل جرام:",
        "language": "اللغة / Language",
        "arabic": "العربية",
        "english": "الإنجليزية",
        "error": "خطأ",
        "success": "نجاح",
        "positive_values_error": "القيم يجب أن تكون أرقاماً موجبة",
        "calculation_success": "تم الحساب بنجاح!",
        "input_error": "خطأ في المدخلات، يرجى مراجعة العمليات الحسابية المكتوبة",
        "no_results_yet": "لا توجد نتائج بعد",
        "theme": "السمة",
        "tag_price": "سعر الباركود (درهم):",
        "cost_tag_divided_by": "تقسيم التكلفة (Tag divided by):",
        "calculated_cost": "التكلفة المحسوبة:",
        "diamond_customer_price": "سعر العميل (درهم):",
        "proposed_discount": "الخصم المقترح (%):",
        "item_cost": "تكلفة القطعة:",
        "diamond_tax": "الضريبة (4.7619%):",
        "diamond_final_price": "السعر النهائي (العميل):",
        "diamond_net_price": "صافي السعر (الشركة):",
        "diamond_gross_profit": "إجمالي الربح (الشركة):",
        "diamond_gross_margin": "هامش الربح الإجمالي:",
        "calc_tip": "💡 ملحوظة: يمكنك كتابة عمليات رياضية مباشرة مثل: 100*2.5 أو 50+30 ثم اضغط Enter"
    }
    
    if st.session_state.language == "ar":
        return translations_ar.get(text, text)
    return translations.get(text, text)

# --- دالة الذكاء الرياضي لمعالجة النصوص الحسابية ---
def evaluate_expression(val_str):
    if not val_str:
        return 0.0
    val_str = str(val_str).strip()
    if val_str.startswith('='):
        val_str = val_str[1:]
    allowed_chars = set('0123456789.-+*/() ')
    if not all(c in allowed_chars for c in val_str):
        try:
            return float(val_str)
        except:
            return None
    try:
        return float(eval(val_str))
    except:
        return None

# --- تصميم القائمة الجانبية (Sidebar) ---
with st.sidebar:
    st.title("🛡️ Core Settings")
    st.markdown("---")
    
    # اختيار الأقسام الأساسية (حاسبة الذهب أو الألماس)
    active_tab = st.radio("Select Calculator Tab", [tr("gold_price_tab"), tr("diamond_price_tab")])
    
    st.markdown("---")
    # إعدادات اللغة
    lang_choice = st.selectbox(tr("language"), ["English", "العربية"], 
                               index=0 if st.session_state.language == "en" else 1)
    new_lang = "ar" if lang_choice == "العربية" else "en"
    if new_lang != st.session_state.language:
        st.session_state.language = new_lang
        st.rerun()

# --- عنوان التطبيق الرئيسي ---
st.title(tr("company"))
st.caption(tr("calc_tip"))
st.markdown("---")

# ==========================================
# --- شاشة حاسبة الذهب (Gold Calculator) ---
# ==========================================
if active_tab == tr("gold_price_tab"):
    st.header(f"✨ {tr('gold_price_tab')}")
    
    # المدخلات في أعمدة متناسقة
    col_in1, col_in2, col_in3 = st.columns(3)
    with col_in1:
        carat_input = st.selectbox(tr("carat"), CARAT_OPTIONS, index=1) # الافتراضي عيار 18
        gold_price_raw = st.text_input(f"{tr('gold_price')} (AED/gram)", value="0.0")
    with col_in2:
        weight_raw = st.text_input(tr("weight"), value="0.0")
        manufacturing_raw = st.text_input(tr("manufacturing"), value="0.0")
    with col_in3:
        suggested_price_raw = st.text_input(tr("suggested_price"), value="")
        discount_raw = st.text_input(tr("discount"), value="0.0", 
                                      disabled=True if suggested_price_raw.strip() and evaluate_expression(suggested_price_raw) > 0 else False)

    # زر الحساب ومسح البيانات
    btn_c1, btn_c2 = st.columns([1, 5])
    with btn_c1:
        calc_gold = st.button(tr("calculate"), type="primary", use_container_width=True)
    with btn_c2:
        if st.button(tr("clear"), use_container_width=False):
            st.session_state.gold_results = None
            st.rerun()

    if calc_gold:
        # تقييم وحساب المدخلات رياضياً
        g_price = evaluate_expression(gold_price_raw)
        w_weight = evaluate_expression(weight_raw)
        m_manuf = evaluate_expression(manufacturing_raw)
        s_price = evaluate_expression(suggested_price_raw) if suggested_price_raw.strip() else 0.0
        d_disc = evaluate_expression(discount_raw) if discount_raw else 0.0
        
        if g_price is None or w_weight is None or m_manuf is None or d_disc is None:
            st.error(tr("input_error"))
        elif g_price <= 0 or w_weight <= 0 or m_manuf < 0:
            st.error(tr("positive_values_error"))
        else:
            # بدء المعادلات المحاسبية المعتمدة للشركة
            piece_cost = ((m_manuf / 3) + g_price) * w_weight
            
            if m_manuf > 60:
                customer_price_before_discount = ((m_manuf + g_price) * w_weight) * 1.3
            else:
                customer_price_before_discount = ((m_manuf + g_price) * w_weight) * 1.5
                
            if s_price > 0:
                customer_price_after_discount_calculated = s_price
                calculated_discount_amount = max(0.0, customer_price_before_discount - s_price)
                calculated_discount_percentage = (calculated_discount_amount / customer_price_before_discount) * 100 if customer_price_before_discount > 0 else 0
                d_disc = calculated_discount_percentage
            else:
                discount_amount = customer_price_before_discount * (d_disc / 100)
                customer_price_after_discount_calculated = customer_price_before_discount - discount_amount
                
            discount_amount = customer_price_before_discount - customer_price_after_discount_calculated
            vat_rate = 0.047619
            vat_amount = customer_price_after_discount_calculated * vat_rate
            net_price_for_company = customer_price_after_discount_calculated - vat_amount
            net_profit = net_price_for_company - piece_cost
            gross_margin = (net_profit / net_price_for_company) * 100 if net_price_for_company > 0 else 0
            manufacturing_cost_per_gram = (net_price_for_company / w_weight) - g_price if w_weight > 0 else 0
            
            st.session_state.gold_results = {
                "gold_price": g_price, "weight": w_weight, "manufacturing": m_manuf,
                "discount": d_disc, "suggested_price": s_price, "carat": carat_input,
                "piece_cost": piece_cost, "customer_price_before_discount": customer_price_before_discount,
                "discount_amount": discount_amount, "customer_price_after_discount": customer_price_after_discount_calculated,
                "vat_amount": vat_amount, "net_price_for_company": net_price_for_company,
                "net_profit": net_profit, "gross_margin": gross_margin, "manufacturing_cost_per_gram": manufacturing_cost_per_gram,
                "calculation_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            st.success(tr("calculation_success"))

    # عرض النتائج المحسوبة
    st.markdown("---")
    if st.session_state.gold_results:
        res = st.session_state.gold_results
        col_res1, col_res2 = st.columns(2)
        
        with col_res1:
            st.subheader(f"👤 {tr('customer_results')}")
            st.metric(tr("piece_cost"), f"{res['piece_cost']:.2f} AED")
            st.metric(tr("price_before_discount"), f"{res['customer_price_before_discount']:.2f} AED")
            st.metric(tr("discount_amount"), f"{res['discount_amount']:.2f} AED ( {res['discount']:.2f}% )")
            st.metric(tr("price_after_discount"), f"{res['customer_price_after_discount']:.2f} AED")
            st.metric(tr("vat"), f"{res['vat_amount']:.2f} AED")
            st.subheader(f"💰 {tr('final_price')}: `{res['customer_price_after_discount']:.2f} AED`")
            
        with col_res2:
            st.subheader(f"🏢 {tr('company_results')}")
            st.metric(tr("net_price"), f"{res['net_price_for_company']:.2f} AED")
            st.metric(tr("net_profit"), f"{res['net_profit']:.2f} AED")
            st.metric(tr("gross_margin"), f"{res['gross_margin']:.2f}%")
            st.metric(tr("manufacturing_cost_label"), f"{res['manufacturing_cost_per_gram']:.2f} AED/gram")
            
        # أزرار تنزيل وتصدير التقارير أونلاين
        st.markdown("---")
        down_col1, down_col2 = st.columns(2)
        with down_col1:
            js_data = json.dumps(res, ensure_ascii=False, indent=4)
            st.download_button(tr("save"), data=js_data, file_name=f"gold_calc_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json", mime="application/json")
        with down_col2:
            # إنشاء ملف إكسل ديناميكي للتنزيل المباشر
            excel_data = [
                ["Gold Price", res['gold_price'], "AED/gram"], ["Weight", res['weight'], "grams"],
                ["Manufacturing", res['manufacturing'], "AED/gram"], ["Discount", res['discount'], "%"],
                ["Piece Cost", res['piece_cost'], "AED"], ["Price Before Discount", res['customer_price_before_discount'], "AED"],
                ["Final Price", res['customer_price_after_discount'], "AED"], ["Net Profit", res['net_profit'], "AED"],
                ["Gross Margin", f"{res['gross_margin']:.2f}", "%"]
            ]
            df = pd.DataFrame(excel_data, columns=["Parameter", "Value", "Unit"])
            df.to_excel("temp_gold.xlsx", index=False)
            with open("temp_gold.xlsx", "rb") as f:
                st.download_button(tr("export_excel"), data=f.read(), file_name=f"gold_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info(tr("no_results_yet"))

# =============================================
# --- شاشة حاسبة الألماس (Diamond Calculator) ---
# =============================================
elif active_tab == tr("diamond_price_tab"):
    st.header(f"💎 {tr('diamond_price_tab')}")
    
    col_dia1, col_dia2 = st.columns(2)
    with col_dia1:
        tag_price_raw = st.text_input(tr("tag_price"), value="0.0")
        cost_divisor_input = st.selectbox(tr("cost_tag_divided_by"), DIVISOR_OPTIONS, index=0) # الافتراضي 1.5
    with col_dia2:
        customer_price_diamond_raw = st.text_input(tr("diamond_customer_price"), value="")
        proposed_discount_diamond_raw = st.text_input(tr("proposed_discount"), value="")

    # إدارة العلاقات المتبادلة بين السعر والخصم أونلاين
    t_price = evaluate_expression(tag_price_raw) or 0.0
    c_price = evaluate_expression(customer_price_diamond_raw) if customer_price_diamond_raw.strip() else None
    p_disc = evaluate_expression(proposed_discount_diamond_raw) if proposed_discount_diamond_raw.strip() else None

    if t_price > 0:
        if c_price is not None and p_disc is None:
            p_disc = ((t_price - c_price) / t_price) * 100
            st.caption(f"💡 Calculated Discount from Customer Price: **{max(0.0, p_disc):.2f}%**")
        elif p_disc is not None and c_price is None:
            c_price = t_price * (1 - (p_disc / 100))
            st.caption(f"💡 Calculated Customer Price from Discount: **{max(0.0, c_price):.2f} AED**")

    # زر الحساب ومسح البيانات
    btn_d1, btn_d2 = st.columns([1, 5])
    with btn_d1:
        calc_diamond = st.button(tr("calculate"), type="primary", use_container_width=True)
    with btn_d2:
        if st.button(tr("clear"), use_container_width=False):
            st.session_state.diamond_results = None
            st.rerun()

    if calc_diamond:
        if t_price <= 0:
            st.error(tr("positive_values_error"))
        elif c_price is None and p_disc is None:
            st.error("Please enter either Customer Price or Proposed Discount!")
        else:
            item_cost = t_price / cost_divisor_input
            vat_rate = 0.047619
            tax_amount = c_price * vat_rate
            net_price_company = c_price - tax_amount
            gross_profit = net_price_company - item_cost
            gross_margin = (gross_profit / net_price_company) * 100 if net_price_company > 0 else 0
            
            st.session_state.diamond_results = {
                "tag_price": t_price, "cost_divisor": cost_divisor_input, "customer_price": c_price, 
                "proposed_discount": p_disc, "calculated_item_cost": item_cost, "diamond_tax_amount": tax_amount,
                "net_price_company": net_price_company, "gross_profit": gross_profit, "gross_margin": gross_margin,
                "calculation_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            st.success(tr("calculation_success"))

    # عرض نتائج الألماس
    st.markdown("---")
    if st.session_state.diamond_results:
        res_d = st.session_state.diamond_results
        col_res_d1, col_res_d2 = st.columns(2)
        
        with col_res_d1:
            st.subheader(f"👤 {tr('customer_results')}")
            st.metric(tr("item_cost"), f"{res_d['calculated_item_cost']:.2f} AED")
            st.metric(tr("diamond_tax"), f"{res_d['diamond_tax_amount']:.2f} AED")
            st.subheader(f"💎 {tr('diamond_final_price')}: `{res_d['customer_price']:.2f} AED`")
            
        with col_res_d2:
            st.subheader(f"🏢 {tr('company_results')}")
            st.metric(tr("diamond_net_price"), f"{res_d['net_price_company']:.2f} AED")
            st.metric(tr("diamond_gross_profit"), f"{res_d['gross_profit']:.2f} AED")
            st.metric(tr("diamond_gross_margin"), f"{res_d['gross_margin']:.2f}%")
            
        # أزرار تحميل التقارير للألماس أونلاين
        st.markdown("---")
        down_dia_col1, down_dia_col2 = st.columns(2)
        with down_dia_col1:
            js_data_d = json.dumps(res_d, ensure_ascii=False, indent=4)
            st.download_button(tr("save"), data=js_data_d, file_name=f"diamond_calc_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json", mime="application/json")
        with down_dia_col2:
            excel_data_d = [
                ["Barcode Price", res_d['tag_price'], "AED"], ["Cost Divisor", res_d['cost_divisor'], ""],
                ["Calculated Cost", res_d['calculated_item_cost'], "AED"], ["Customer Price", res_d['customer_price'], "AED"],
                ["Proposed Discount", f"{res_d['proposed_discount']:.2f}", "%"], ["Gross Profit", res_d['gross_profit'], "AED"],
                ["Gross Margin", f"{res_d['gross_margin']:.2f}", "%"]
            ]
            df_d = pd.DataFrame(excel_data_d, columns=["Parameter", "Value", "Unit"])
            df_d.to_excel("temp_diamond.xlsx", index=False)
            with open("temp_diamond.xlsx", "rb") as f:
                st.download_button(tr("export_excel"), data=f.read(), file_name=f"diamond_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info(tr("no_results_yet"))
