# -------------------------------
# 1. استدعاء المكتبات اللازمة
# -------------------------------
import streamlit as st
import pandas as pd
import datetime
from io import BytesIO
import calendar
import plotly.express as px
import time
import numpy as np


st.set_page_config(page_title="💪🔥 MRP Tool", page_icon="👍", layout="wide")




if "loaded" not in st.session_state:
    print("Loading once")
    st.session_state.loaded = True

def optimize_and_sanitize_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    دالة هندسية لتطهير الـ DataFrame وتجهيزه لطبقة معالجة PyArrow في Streamlit.
    تحافظ على الأداء العالي وتمنع الـ Mixed Types وانهيار الأعمدة.
    """
    # نسخ خفيف للـ DataFrame للحفاظ على البيانات الأصلية
    cleaned_df = df.copy()

    # 1. تسوية أسماء الأعمدة (Flattening) ومنع كوارث الـ MultiIndex/Tuples
    if isinstance(cleaned_df.columns, pd.MultiIndex):
        cleaned_df.columns = ['_'.join(str(c).strip() for c in col if c).strip() for col in cleaned_df.columns.values]
    else:
        cleaned_df.columns = [str(col).strip() for col in cleaned_df.columns]

    # 2. تطهير ومعالجة أنواع البيانات كـ Vectors (تجنب الحلقات التكرارية البطيئة)
    for col in cleaned_df.columns:
        # فحص الأنواع المختلطة أو الأعمدة التي تبدو كـ Objects
        if cleaned_df[col].dtype == 'object':
            # استبدال النصوص الفارغة أو المسافات بقيم نال حقيقية
            # هذا يمنع خطأ "tried to convert to double"
            if cleaned_df[col].astype(str).str.contains(r'^\s*$', regex=True).any():
                cleaned_df[col] = cleaned_df[col].replace(r'^\s*$', np.nan, regex=True)

            # محاولة تحويل الأعمدة الرقمية التي تلوثت بنصوص بشكل ذكي
            # errors='ignore' تضمن أنه لو كان العمود نصياً حقيقياً (كـ اسم المنتج) لا يفسد
            try:
                converted_numeric = pd.to_numeric(cleaned_df[col], errors='coerce')
                # إذا كانت نسبة القيم الرقمية المحولة عالية جداً، نعتمد التحويل الرقمي
                if converted_numeric.notna().sum() > (len(cleaned_df) * 0.5):
                    cleaned_df[col] = converted_numeric
            except Exception:
                pass

    # 3. التحويل الصارم لـ Strings الصافية لمنع PyArrow من التخمين الخاطئ
    # الأفضل تحويل الـ object النصي إلى نوع 'string' المخصص في Pandas الحديثة
    for col in cleaned_df.select_dtypes(include=['object']).columns:
        cleaned_df[col] = cleaned_df[col].astype(str).replace('nan', '')

    return cleaned_df

# --- الاستخدام المباشر في خط الأنابيب ---
# processed_df = execute_mrp_explosion_or_costing()
# sanitized_df = optimize_and_sanitize_df(processed_df)
# st.dataframe(sanitized_df)



try:
    from statsmodels.tsa.holtwinters import ExponentialSmoothing
except ImportError:
    ExponentialSmoothing = None

# ==============================================================================
# 1a. إعدادات الواجهة والعمق الافتراضي (V4)
# ==============================================================================
DEFAULT_MAX_BOM_LEVEL = 10
DEFAULT_DARK_MODE = False

if "dark_mode" not in st.session_state:
    st.session_state.dark_mode = DEFAULT_DARK_MODE
if "max_bom_level" not in st.session_state:
    st.session_state.max_bom_level = DEFAULT_MAX_BOM_LEVEL

def get_max_bom_depth() -> int:
    try:
        return max(1, int(st.session_state.get("max_bom_level", DEFAULT_MAX_BOM_LEVEL)))
    except Exception:
        return DEFAULT_MAX_BOM_LEVEL

def apply_app_theme():
    dark_mode = bool(st.session_state.get("dark_mode", DEFAULT_DARK_MODE))
    bg = "#0E1117" if dark_mode else "#F7F9FC"

    surface = "#161B22" if dark_mode else "#FFFFFF"

    text = "#E6EDF3" if dark_mode else "#162447"

    border = "#30363D" if dark_mode else "#D9E2EC"

    accent = "#00ADB5" if dark_mode else "#0F4C75"

    accent_soft = "rgba(56, 189, 248, 0.15)" if dark_mode else "rgba(14, 165, 233, 0.10)"

    shadow = "0 10px 25px rgba(0,0,0,0.35)" if dark_mode else "0 10px 25px rgba(56, 189, 248, 0.12)"

    st.markdown(f"""
    <style>
    :root {{ --app-bg: {bg}; --app-surface: {surface}; --app-text: {text}; --app-border: {border}; --app-accent: {accent}; --app-accent-soft: {accent_soft}; --app-shadow: {shadow}; }}
    .stApp, [data-testid=\"stAppViewContainer\"], [data-testid=\"stHeader\"], [data-testid=\"stToolbar\"] {{ background: var(--app-bg) !important; color: var(--app-text) !important; }}
    [data-testid=\"stSidebar\"] {{ background: var(--app-surface) !important; border-left: 1px solid var(--app-border) !important; }}
    html, body, .stMarkdown, p, h1, h2, h3, h4, h5, h6, label, span, div {{ color: var(--app-text); }}

    .stMetric, div[data-testid=\"metric-container\"], .stDataFrame, .stTable, .stExpander, [data-testid=\"stFileUploader\"], .stAlert, .stSelectbox, .stMultiSelect, .stTextInput, .stTextArea, .stDateInput {{ background: var(--app-surface); border-color: var(--app-border) !important; }}
    .stMetric, div[data-testid=\"metric-container\"], .stExpander, [data-testid=\"stFileUploader\"], .stAlert {{ border: 1px solid var(--app-border) !important; border-radius: 14px !important; box-shadow: var(--app-shadow) !important; }}

    .stButton > button, .stDownloadButton > button {{ background: linear-gradient(135deg, var(--app-accent), #60A5FA) !important; color: #ffffff !important; border: none !important; border-radius: 12px !important; min-height: 42px !important; font-weight: 700 !important; font-size: 16px !important; box-shadow: 0 4px 12px rgba(56,189,248,0.25);transition: all 0.25s ease !important; }}

    .stButton > button:hover, .stDownloadButton > button:hover {{ transform: translateY(-4px); box-shadow: 0 8px 18px rgba(56,189,248,0.35); filter: brightness(1.05); }}

    .stExpander summary p {{ color: var(--app-text) !important; font-size: 20px !important; }}
    div[data-baseweb=\"select\"] > div, input, textarea {{ background: var(--app-surface) !important; color: var(--app-text) !important; border: 1px solid var(--app-border) !important; }}

    .stTabs [data-baseweb=\"tab\"] {{ color: var(--app-text) !important; border-radius: 10px 10px 0 0 !important; padding: 0.6rem 1rem !important; }}

    .stTabs [aria-selected=\"true\"] {{ background: var(--app-accent-soft) !important; color: var(--app-accent) !important; font-weight: 700 !important; }}

    @media (max-width: 768px) {{ .block-container {{ padding-top: 0.9rem !important; padding-left: 0.8rem !important; padding-right: 0.8rem !important; }} h1 {{ font-size: 1.55rem !important; line-height: 1.35 !important; }} h2 {{ font-size: 1.25rem !important; }} h3 {{ font-size: 1.08rem !important; }} .stExpander summary p {{ font-size: 16px !important; }} div[data-testid=\"metric-container\"] {{ padding: 0.65rem !important; }} .stButton > button, .stDownloadButton > button {{ width: 100% !important; min-height: 44px !important; }} [data-testid=\"column\"] {{ min-width: 100% !important; flex: 1 1 100% !important; }} .stTabs [data-baseweb=\"tab\"] {{ font-size: 0.9rem !important; padding: 0.55rem 0.7rem !important; }} }}

    </style>
    """, unsafe_allow_html=True)

# ==============================================================================
# 1b. دوال التنسيق الموحدة (V3)
# ==============================================================================
def fmt_n(v) -> str:
    """تنسيق الأعداد الصحيحة الكبيرة:  1,234,567"""
    try:
        return f"{float(v):,.0f}"
    except (TypeError, ValueError):
        return str(v) if v is not None else "—"

def fmt_f(v, decimals: int = 2) -> str:
    """تنسيق الأعداد العشرية: 1,234,567.89"""
    try:
        return f"{float(v):,.{decimals}f}"
    except (TypeError, ValueError):
        return str(v) if v is not None else "—"

def fmt_pct(v, decimals: int = 1) -> str:
    """تنسيق النسب المئوية: 12.3%"""
    try:
        return f"{float(v):.{decimals}f}%"
    except (TypeError, ValueError):
        return "—"

# ==============================================================================
# 2. إعداد التكوين والأعمدة
# ==============================================================================
COLUMN_NAMES = {
    "material":             ["Material", "Item", "code", "Code", "المادة", "Product"],
    "material_desc":        ["Material Description", "Description", "وصف"],
    "order_type":           ["Order Type", "OT", "نوع الطلب", "Sales Org."],
    "component":            ["Component", "Comp", "المكون"],
    "component_desc":       ["Component Description", "Comp Desc", " المسمى", "وصف المكون"],
    "component_uom":        ["Component UoM", "UoM", "الوحدة"],
    "component_qty":        ["Component Quantity", "Qty", "كمية المكون"],
    "base_qty":             ["Base Quantity", "Base Qty", "الكمية الأساسية"],
    "mrp_controller":       ["MRP Controller", "مسؤول MRP"],
    "current_stock":        ["Current Stock", "Stock", "المخزون الحالي", "Unrestricted"],
    "component_order_type": ["Component Order Type", "Order Category", "نوع أمر المكون", "Procurement Type"],
    "hierarchy_level":      ["Hierarchy Level", "Level", "المستوى الهرمي"],
    "parent_material":      ["Parent Material", "Direct Parent", "الأب المباشر"],
    "price":                ["Price", "Unit Cost", "Last price", "Cost", "السعر", "سعر الوحدة", "تكلفة الوحدة"],
}

def col(name_key):
    """إرجاع اسم العمود الرئيسي"""
    return COLUMN_NAMES[name_key][0]

def normalize_columns(df, column_map):
    """توحيد أسماء الأعمدة إلى الاسم الرئيسي"""
    rename_dict = {}
    for key, aliases in column_map.items():
        for alias in aliases:
            if alias in df.columns and alias != aliases[0]:
                rename_dict[alias] = aliases[0]
    return df.rename(columns=rename_dict)

# ==============================================================================
# 3. دالة تحميل البيانات والتحقق منها
# ==============================================================================
@st.cache_data

def load_and_validate_data(uploaded_file):
    try:
        xls = pd.ExcelFile(uploaded_file, engine='openpyxl')

        # --- التحقق من الأوراق ---
        required_sheets = ["plan", "Component"]
        missing_sheets = [s for s in required_sheets if s not in xls.sheet_names]
        if missing_sheets:
            st.error(f"❌ الملف لا يحتوي على الأوراق المطلوبة: {', '.join(missing_sheets)}")
            st.stop()

        # --- تحميل البيانات (مرة واحدة فقط) ---
        raw_plan = xls.parse("plan")
        raw_component = xls.parse("Component")

        plan_df = normalize_columns(raw_plan, COLUMN_NAMES)
        component_df = normalize_columns(raw_component, COLUMN_NAMES)

        mrp_df = (
            normalize_columns(xls.parse("MRP Controller"), COLUMN_NAMES)
            if "MRP Controller" in xls.sheet_names
            else pd.DataFrame()
        )

        # ✅ إزالة الأعمدة الزائدة غير المعروفة من ورقة Component
        known_cols = [aliases[0] for aliases in COLUMN_NAMES.values()]
        extra_cols = [c for c in component_df.columns if c not in known_cols]
        if extra_cols:
            component_df.drop(columns=extra_cols, inplace=True)

        # --- التحقق من الأعمدة الأساسية ---
        required_plan_cols = [col("material"), col("material_desc"), col("order_type")]
        if not all(c in plan_df.columns for c in required_plan_cols):
            st.error(f"❌ جدول الخطة ناقص أعمدة: {required_plan_cols}")
            st.stop()

        required_comp_cols = [col("material"), col("component"), col("component_qty")]
        if not all(c in component_df.columns for c in required_comp_cols):
            st.error(f"❌ جدول المكونات ناقص أعمدة: {required_comp_cols}")
            st.stop()

        # --- تنظيف الأعمدة الرقمية في component_df ---
        comp_qty_col = col("component_qty")
        base_qty_col = col("base_qty")

        component_df[comp_qty_col] = pd.to_numeric(component_df[comp_qty_col], errors='coerce').fillna(0)

        # ✅ FIX: معالجة Base Quantity بدون إعادة قراءة Excel

        if base_qty_col in component_df.columns:

            base_series = pd.to_numeric(component_df[base_qty_col], errors='coerce')

            zero_base = (base_series == 0).sum()

            if zero_base > 0:
                st.warning(
                    f"⚠️ يوجد {zero_base} قيمة صفرية في عمود Base Quantity — تم استبدالها بـ 1 تلقائياً."
                )
            base_series = base_series.fillna(1).replace(0, 1)
            component_df[comp_qty_col] = component_df[comp_qty_col] / base_series
            component_df.drop(columns=[base_qty_col], inplace=True)


        # --- الأعمدة الاختيارية مع قيم افتراضية ---
        if col("current_stock") not in component_df.columns:
            component_df[col("current_stock")] = 0
        else:
            component_df[col("current_stock")] = pd.to_numeric(
                component_df[col("current_stock")], errors='coerce'
            ).fillna(0)

        if col("component_order_type") not in component_df.columns:
            component_df[col("component_order_type")] = "غير محدد"

        if col("hierarchy_level") not in component_df.columns:
            component_df[col("hierarchy_level")] = 1
        else:
            component_df[col("hierarchy_level")] = pd.to_numeric(
                component_df[col("hierarchy_level")], errors='coerce'
            ).fillna(1).astype(int)

        if col("component_desc") not in component_df.columns:
            component_df[col("component_desc")] = ""

        if col("component_uom") not in component_df.columns:
            component_df[col("component_uom")] = ""

        if col("mrp_controller") not in component_df.columns:
            component_df[col("mrp_controller")] = "غير محدد"
        component_df[col("mrp_controller")] = (
            component_df[col("mrp_controller")]
            .replace(r"^\s*$", pd.NA, regex=True)
            .fillna("غير محدد")
            .astype(str)
            .str.strip())

        # ✅ تنظيف عمود Parent Material إن وُجد
        # هذا العمود يحتوي على الأب المباشر الفعلي لكل مكون (من SAP CS12)
        if col("parent_material") in component_df.columns:
            component_df[col("parent_material")] = (
                component_df[col("parent_material")].astype(str).str.strip()
            )
        # إذا لم يكن موجوداً → نُنشئه من Material (fallback للتوافق مع ملفات قديمة)
        else:
            component_df[col("parent_material")] = component_df[col("material")]

        # ✅ توحيد وحدات الوزن إلى كيلوجرام
        # أي مكون وحدته G أو g أو GM أو gram → نقسم الكمية والرصيد على 1000 ونغير الوحدة إلى KG
        gram_variants = {"g", "gm", "gr", "gram", "grams", "جرام", "جم"}
        uom_col = col("component_uom")
        qty_col = col("component_qty")
        stk_col = col("current_stock")

        is_gram = component_df[uom_col].astype(str).str.strip().str.lower().isin(gram_variants)

        if is_gram.any():
            component_df.loc[is_gram, qty_col] = component_df.loc[is_gram, qty_col] / 1000
            component_df.loc[is_gram, stk_col] = component_df.loc[is_gram, stk_col] / 1000
            component_df.loc[is_gram, uom_col] = "KG"

        # ✅ NEW: توحيد وحدات المساحة من CM2 إلى M2
        cm2_variants = {"cm2", "cm^2", "cm²", "سم2", "سم²"}
        is_cm2 = component_df[uom_col].astype(str).str.strip().str.lower().isin(cm2_variants)

        if is_cm2.any():
            component_df.loc[is_cm2, qty_col] = component_df.loc[is_cm2, qty_col] / 10000
            component_df.loc[is_cm2, stk_col] = component_df.loc[is_cm2, stk_col] / 10000
            component_df.loc[is_cm2, uom_col] = "M2"

        return plan_df, component_df , mrp_df

    except Exception as e:
        st.error(f"❌ فشل تحميل الملف: {str(e)}")
        st.stop()

# ==============================================================================
# ✅ V3: دالة BOM Explosion متعددة المستويات — تكرارية (Iterative) بدلاً من التعاودية
# ==============================================================================

def bom_explosion(plan_melted, component_df, max_depth=None):
    """
    Multi-Level BOM Explosion — النهج الصحيح لـ SAP CS12  (V3 — Iterative)

    V3 Changes:
    - استُبدلت الدالة التعاودية explode() بحلقة BFS/stack تكرارية باستخدام deque
    - يُجنّب Python recursion limit مع BOMs عميقة (>950 مستوى)
    - أسرع بـ 10-50x على BOMs الكبيرة بسبب إزالة call-stack overhead
    - comp_info الآن dict مبني مسبقاً (O(1) lookup) بدلاً من merge في النهاية

    الخوارزمية:
    1. groupby(Material + Parent + Component) → bom_core فريد لكل منتج
    2. bom_dict per Material → tree[parent] = [(comp, qty), ...]
    3. Iterative explosion باستخدام deque لكل صف في الخطة مستقلاً
    """
    from collections import defaultdict, deque

    max_depth = int(max_depth or get_max_bom_depth())

    # ── تنظيف أولي (بدون .copy() غير ضروري — نعمل على slice) ──────────────
    component_df = component_df[[c for c in component_df.columns]].copy()
    component_df[col("component")] = component_df[col("component")].astype(str).str.strip()
    component_df[col("material")]  = component_df[col("material")].astype(str).str.strip()

    has_parent_col = col("parent_material") in component_df.columns
    parent_col = col("parent_material") if has_parent_col else col("material")
    if has_parent_col:
        component_df[parent_col] = component_df[parent_col].astype(str).str.strip()
        empty_parent_mask = component_df[parent_col].isin(["", "nan", "None", "NaN"])
        if empty_parent_mask.any():
            component_df.loc[empty_parent_mask, parent_col] = component_df.loc[empty_parent_mask, col("material")].astype(str).str.strip()

    # ── STEP 1: drop duplicates + groupby ──────────────────────────────────
    component_df = component_df.drop_duplicates(
        subset=[col("material"), parent_col, col("component"), col("component_qty")],
        keep="first"
    )
    bom_core = component_df.groupby(
        [col("material"), parent_col, col("component")],
        as_index=False
    )[col("component_qty")].sum()

    # ── STEP 2: bom_dict منفصل لكل Material (V3: zip أسرع من iterrows/itertuples مع أسماء بمسافات) ──
    bom_dict: dict[str, dict] = {}
    _mat_arr    = bom_core[col("material")].values
    _par_arr    = bom_core[parent_col].values
    _comp_arr   = bom_core[col("component")].values
    _qty_arr    = bom_core[col("component_qty")].values

    for mat, group in bom_core.groupby(col("material")):
        tree: dict = defaultdict(list)
        mask = _mat_arr == mat
        for par, comp, qty_val in zip(_par_arr[mask], _comp_arr[mask], _qty_arr[mask]):
            tree[par].append((comp, float(qty_val)))
        bom_dict[mat] = tree

    # ── STEP 2b: comp_info كـ dict (O(1) lookup, أسرع من merge) ────────────
    _info_cols = [col("component_desc"), col("component_uom"),
                  col("mrp_controller"), col("current_stock"),
                  col("component_order_type")]
    _info_cols = [c for c in _info_cols if c in component_df.columns]
    comp_info_dict: dict[str, dict] = (
        component_df
        .drop_duplicates(subset=[col("component")], keep="last")
        .set_index(col("component"))[_info_cols]
        .to_dict(orient="index")
    )

    # ── STEP 3: Iterative explosion باستخدام deque ──────────────────────────
    def explode_iterative(root_material: str, root_qty: float) -> list[dict]:
        """
        BFS/Stack تكراري — يستبدل الـ recursion تماماً.
        المكدس: (parent, qty, level, visited_frozenset)
        """
        row_buf = []
        # كل عنصر في المكدس: (parent_node, cumulative_qty, bom_level, visited_set)
        stack = deque()
        stack.append((root_material, root_qty, 1, frozenset()))

        while stack:
            parent, qty, level, visited = stack.pop()

            if parent in visited or level > max_depth:
                continue

            # البحث في شجرة الجذر أولاً، ثم في شجرة الأب (نصف مصنّع)
            tree = bom_dict.get(root_material, {})
            children = tree.get(parent, [])
            if not children:
                tree = bom_dict.get(parent, {})
                children = tree.get(parent, [])

            if not children:
                continue

            new_visited = visited | {parent}
            for comp, comp_qty in children:
                needed = qty * comp_qty
                row_buf.append({
                    "Parent":                      parent,
                    col("component"):              comp,
                    col("component_qty"):          comp_qty,
                    "Required Component Quantity": needed,
                    "BOM Level":                   level,
                })
                # أضف المستوى التالي للمكدس (deque.append = LIFO → DFS order مطابق للأصل)
                stack.append((comp, needed, level + 1, new_visited))

        return row_buf

    # ── STEP 4: تشغيل explosion لكل صف في الخطة ───────────────────────────
    plan_positive = plan_melted[plan_melted["Planned Quantity"] > 0]
    all_rows: list[dict] = []

    for _, plan_row in plan_positive.iterrows():
        mat      = str(plan_row[col("material")]).strip()
        qty      = plan_row["Planned Quantity"]
        ot       = plan_row[col("order_type")]
        date     = plan_row["Date"]
        mat_desc = str(plan_row.get(col("material_desc"), "")).strip()

        row_buf = explode_iterative(mat, qty)
        for r in row_buf:
            r[col("material")]      = mat
            r[col("material_desc")] = mat_desc
            r["Order Type"]         = ot
            r["Date"]               = date
        all_rows.extend(row_buf)

    if not all_rows:
        return pd.DataFrame()

    result = pd.DataFrame(all_rows)

    # ── STEP 5: إضافة الأعمدة الوصفية بالـ map (أسرع من merge) ──────────────
    for info_col in _info_cols:
        result[info_col] = result[col("component")].map(
            lambda c, ic=info_col: comp_info_dict.get(c, {}).get(ic, ""))

    return result

# =============================================================================
# Build BOM Paths (Recursive Multi-Level)
# =============================================================================
def build_bom_paths(component_df):

    parent_col = col("material")
    child_col = col("component")
    qty_col = col("component_qty")

    # -----------------------------------------------------------------------------
    # Copy & Clean
    # -----------------------------------------------------------------------------
    df = component_df.copy()

    df[parent_col] = df[parent_col].astype(str).str.strip()
    df[child_col] = df[child_col].astype(str).str.strip()

    df[qty_col] = pd.to_numeric(
        df[qty_col],
        errors="coerce"
    ).fillna(0)

    # -----------------------------------------------------------------------------
    # Build BOM Dictionary
    # -----------------------------------------------------------------------------
    bom_dict = {}

    for _, row in df.iterrows():
        parent = row[parent_col]
        child = row[child_col]
        qty = row[qty_col]

        if parent not in bom_dict:
            bom_dict[parent] = []

        bom_dict[parent].append((child, qty))

    # -----------------------------------------------------------------------------
    # Parent / Child Analysis
    # -----------------------------------------------------------------------------
    parents = set(df[parent_col].unique())
    children = set(df[child_col].unique())

    # المنتجات النهائية
    top_materials = parents - children

    all_paths = []

    # -----------------------------------------------------------------------------
    # Recursive DFS
    # -----------------------------------------------------------------------------
    def trace_path(current_material,
                   final_product,
                   cumulative_qty,
                   level,
                   path):

        # لو خامة نهائية بدون أبناء
        if current_material not in bom_dict:

            all_paths.append({
                "Leaf_Material": current_material,
                "Final_Product": final_product,
                "Cum_Qty": cumulative_qty,
                "Level": level,
                "Path": " -> ".join(path)
            })

            return

        # استكمال النزول داخل الـ BOM
        for child, qty in bom_dict[current_material]:

            new_qty = cumulative_qty * qty

            trace_path(
                current_material=child,
                final_product=final_product,
                cumulative_qty=new_qty,
                level=level + 1,
                path=path + [child]
            )

    # -----------------------------------------------------------------------------
    # Start Explosion
    # -----------------------------------------------------------------------------
    for material in top_materials:

        trace_path(
            current_material=material,
            final_product=material,
            cumulative_qty=1,
            level=0,
            path=[material]
        )

    # -----------------------------------------------------------------------------
    # Output
    # -----------------------------------------------------------------------------
    df_paths = pd.DataFrame(all_paths)

    return df_paths

    component_df = normalize_columns(component_df, COLUMN_NAMES)


# ==============================================================================
# 3b. دالة BOM Paths — المسارات الأفقية الكاملة لكل مكون
# ==============================================================================
@st.cache_data(show_spinner=False)   # ✅ V3: تخزين مؤقت — يُجنّب إعادة الحساب عند كل تفاعل
def generate_bom_paths(component_df, plan_df=None):
    from collections import defaultdict
    from decimal import Decimal

    component_df = component_df.copy()

    # ── 1. تحديد عمود الأب المباشر ─────────────────────────────────────────
    has_parent_col = col("parent_material") in component_df.columns
    parent_col = col("parent_material") if has_parent_col else col("material")

    component_df[col("component")] = component_df[col("component")].astype(str).str.strip()
    component_df[col("material")]  = component_df[col("material")].astype(str).str.strip()
    component_df[parent_col]       = component_df[parent_col].astype(str).str.strip()

    # V4: إصلاح parent bug — أي Parent فارغ/NaN يُعاد ملؤه من الـ Material
    empty_parent_mask = component_df[parent_col].isin(["", "nan", "None", "NaN"])
    if empty_parent_mask.any():
        component_df.loc[empty_parent_mask, parent_col] = component_df.loc[empty_parent_mask, col("material")].astype(str).str.strip()

    # ── 2. بناء قاموس العلاقات ───────────────────────────────────────────────
    uom_col  = col("component_uom")
    qty_col  = col("component_qty")
    desc_col = col("component_desc")

    bom_core = (
        component_df
        .drop_duplicates(subset=[parent_col, col("component")], keep="first")
        [[parent_col, col("component"), desc_col, qty_col, uom_col]]
    )

    bom_dict = defaultdict(list)
    for _, row in bom_core.iterrows():
        parent     = row[parent_col]
        child      = row[col("component")]
        child_name = str(row.get(desc_col, "")).strip()
        try:
            child_qty = Decimal(str(row.get(qty_col, 1) or 1))
        except Exception:
            child_qty = Decimal("1")
        child_uom  = str(row.get(uom_col, "")).strip()
        bom_dict[parent].append((child, child_name, child_qty, child_uom))

    # ── 3. قاموس أسماء الجذور (من Component و Plan) ─────────────────────────
    root_name_dict = {}
    for _, row in component_df.drop_duplicates(subset=[col("component")]).iterrows():
        code = str(row[col("component")]).strip()
        name = str(row.get(col("component_desc"), "")).strip()
        if code and name:
            root_name_dict[code] = name

    if plan_df is not None:
        for _, row in plan_df.drop_duplicates(subset=[col("material")]).iterrows():
            code = str(row[col("material")]).strip()
            name = str(row.get(col("material_desc"), "")).strip()
            if code and name:
                root_name_dict[code] = name

    # ── 4. تحديد الجذور (كل من يظهر كـ Parent وليس Component في أي مكان آخر) ──
    all_parents = set(component_df[parent_col].unique())
    all_children = set(component_df[col("component")].unique())
    roots = list(all_parents - all_children)   # الجذور الحقيقية

    # إذا لم يظهر أي جذر (حالة نادرة)، استخدم كل الـ Parents الفريدة
    if not roots:
        roots = list(all_parents)

    # ── 5. الدالة التكرارية لبناء المسار ────────────────────────────────────
    def build_paths(node, node_label, current_path, visited, cumulative):
        current_path = current_path + [(node, node_label)]
        if node not in bom_dict:
            return [current_path]

        all_paths = []
        for child, child_name, child_qty, child_uom in bom_dict[node]:
            if child in visited:
                continue
            new_cumulative = cumulative * child_qty
            # تنسيق الكمية
            if new_cumulative == new_cumulative.to_integral_value():
                qty_str = str(new_cumulative.to_integral_value())
            else:
                qty_str = format(new_cumulative.normalize(), 'f')
                if '.' in qty_str:
                    qty_str = qty_str.rstrip('0').rstrip('.')
            uom_str    = f" {child_uom}" if child_uom else ""
            child_label = f"{child_name} , {qty_str}{uom_str}"
            child_paths = build_paths(
                child, child_label, current_path,
                visited | {node}, new_cumulative
            )
            all_paths.extend(child_paths)
        return all_paths if all_paths else [current_path]

    # ── 6. جمع كل المسارات ───────────────────────────────────────────────────
    all_paths = []
    for root in roots:
        root_label = root_name_dict.get(root, "") or ""
        paths = build_paths(root, root_label, [], set(), Decimal("1"))
        all_paths.extend(paths)

    if not all_paths:
        return pd.DataFrame()

    # ── 7. تحويل إلى DataFrame أفقي ─────────────────────────────────────────
    max_depth = max(len(p) for p in all_paths)
    columns   = []
    for i in range(1, max_depth + 1):
        columns.append(f"Level_{i}")
        columns.append(f"Name_{i}")

    rows = []
    for path in all_paths:
        row = {}
        for i, (code, label) in enumerate(path, start=1):
            row[f"Level_{i}"] = code
            row[f"Name_{i}"]  = label
        rows.append(row)

    df_paths = pd.DataFrame(rows, columns=columns)

    df_paths = df_paths.fillna("")
    df_paths = df_paths.astype(str)
    df_paths = df_paths.replace(["nan", "None", "NaN"], "")

    df_paths = df_paths.drop_duplicates().reset_index(drop=True)

    # ── 8. عدد الآباء المباشرين لكل مكون ────────────────────────────────────
    parent_count = (
        bom_core
        .groupby(col("component"))[parent_col]
        .nunique()
        .to_dict()
    )
    if "Level_2" in df_paths.columns:
        df_paths.insert(
            0, "عدد آباء المكون المباشر",
            df_paths["Level_2"].astype(str).str.strip()
            .map(parent_count).fillna(1).astype(int)
        )
    else:
        df_paths.insert(0, "عدد آباء المكون المباشر", 1)

    return df_paths

# =============================================================================
# ── اضافة مسارات عكسية───────────────────────────────────────────
# =============================================================================

def generate_reverse_bom_paths(df_forward_paths, debug_costing_df=None):
    if df_forward_paths is None or df_forward_paths.empty:
        return pd.DataFrame()
        
    level_cols = [c for c in df_forward_paths.columns if c.startswith("Level_")]
    name_cols = [c for c in df_forward_paths.columns if c.startswith("Name_")]
    
    reverse_rows = []
    
    for _, row in df_forward_paths.iterrows():
        final_product_code = str(row.get("Level_1", "")).strip()
        final_product_name = str(row.get("Name_1", "")).strip()
        if " , " in final_product_name:
            final_product_name = final_product_name.split(" , ", 1)[0].strip()
            
        actual_levels = []
        actual_names = []
        
        for l_col, n_col in zip(level_cols, name_cols):
            l_val = str(row.get(l_col, "")).strip()
            n_val = str(row.get(n_col, "")).strip()
            if l_val and l_val != "nan" and l_val != "":
                actual_levels.append(l_val)
                actual_names.append(n_val)
                
        if not actual_levels:
            continue
            
        raw_material_code = actual_levels[-1]

        # ── Level-1 SF Logic (متوافق تماماً مع Costing_Debug) ──────────────────
        # FG → B → C → D → X  ⟹  Parent_Material = B  (actual_levels[1])
        # FG → X               ⟹  Parent_Material = FG (actual_levels[0])  ← مباشر

        if len(actual_levels) >= 3:
            level1_sf_code = actual_levels[1]   # أول مكون وسيط تحت المنتج التام
            level1_sf_name = actual_names[1].split(" , ", 1)[0].strip() if " , " in actual_names[1] else actual_names[1]
        elif len(actual_levels) == 2:
            level1_sf_code = actual_levels[0]   # خامة مرتبطة بالمنتج التام مباشرة
            level1_sf_name = final_product_name
        else:
            level1_sf_code = ""
            level1_sf_name = ""
        
        _raw_name = actual_names[-1]
        cum_qty_f = 1.0
        _uom = ""
        if " , " in _raw_name:
            _qty_part = _raw_name.split(" , ", 1)[1].strip()
            _parts = _qty_part.split()
            try:
                cum_qty_f = float(_parts[0])
                _uom = _parts[1] if len(_parts) > 1 else ""
            except (ValueError, IndexError):
                pass

        _uom_upper = _uom.upper()
        if _uom_upper in ("G", "GM", "GRAM"):
            cum_qty_f /= 1000
            _uom = "KG"
        elif _uom_upper in ("CM2", "CM²"):
            cum_qty_f /= 10000
            _uom = "M2"
            
        reversed_levels = list(reversed(actual_levels))
        reversed_names = list(reversed(actual_names))
        
        new_row = {
            "Raw_Material_Code":  raw_material_code,
            "Path_Cum_Qty":       cum_qty_f,
            "UoM":                _uom,
            "Unit_Price_EGP":     0.0,
            "Path_Extended_Cost": 0.0,
            "Final_Product_Code": final_product_code,
            "Final_Product_Name": final_product_name,
            "Parent_Material":    level1_sf_code,
            "Parent_Material_Name": level1_sf_name,
        }
        
        for i, (l_v, n_v) in enumerate(zip(reversed_levels, reversed_names), start=1):
            clean_n_v = n_v.split(" , ", 1)[0].strip() if " , " in n_v else n_v
            new_row[f"Rev_Level_{i}"] = l_v
            new_row[f"Rev_Name_{i}"] = clean_n_v
            
        reverse_rows.append(new_row)
        
    df_rev = pd.DataFrame(reverse_rows)
    if df_rev.empty:
        return df_rev

    if debug_costing_df is not None and not debug_costing_df.empty:
        _db_df = debug_costing_df.copy()
        if "Product_Code"    in _db_df.columns: _db_df.rename(columns={"Product_Code":    "Final_Product"}, inplace=True)
        if "RM_Code"         in _db_df.columns: _db_df.rename(columns={"RM_Code":         "Raw_Material"},  inplace=True)
        if "Raw_Material"    not in _db_df.columns and "Raw_Material_Code" in _db_df.columns:
            _db_df.rename(columns={"Raw_Material_Code": "Raw_Material"}, inplace=True)
        if "Final_Product"   not in _db_df.columns and "Final_Product_Code" in _db_df.columns:
            _db_df.rename(columns={"Final_Product_Code": "Final_Product"}, inplace=True)
        
        _db_df["Final_Product"] = _db_df["Final_Product"].astype(str).str.strip()
        _db_df["Raw_Material"]  = _db_df["Raw_Material"].astype(str).str.strip()
        
        df_rev["Final_Product_Code"] = df_rev["Final_Product_Code"].astype(str).str.strip()
        df_rev["Raw_Material_Code"]  = df_rev["Raw_Material_Code"].astype(str).str.strip()
        df_rev["Parent_Material"]    = df_rev["Parent_Material"].astype(str).str.strip()

        # ── تجميع debug_costing مع Parent_Material لدقة أعلى عند تكرار الخامة في SFs مختلفة ──

        _has_pm = "Parent_Material" in _db_df.columns
        if _has_pm:
            _db_df["Parent_Material"] = _db_df["Parent_Material"].astype(str).str.strip()
            _group_keys  = ["Final_Product", "Parent_Material", "Raw_Material"]
            _left_keys   = ["Final_Product_Code", "Parent_Material", "Raw_Material_Code"]
            _right_keys  = ["Final_Product",      "Parent_Material", "Raw_Material"]
        else:
            _group_keys = ["Final_Product", "Raw_Material"]
            _left_keys  = ["Final_Product_Code", "Raw_Material_Code"]
            _right_keys = ["Final_Product",      "Raw_Material"]

        _db_grouped = _db_df.groupby(_group_keys, as_index=False).agg({
            "Unit_Price_EGP": "first",
            "Extended_Cost":  "sum",
            "Cum_Qty":        "sum",
            "UoM":            "first",
        })
        
        df_rev = df_rev.merge(
            _db_grouped,
            left_on=_left_keys,
            right_on=_right_keys,
            how="left",
            suffixes=("_old", "")
        )
        
        if "Cum_Qty"       in df_rev.columns: df_rev["Path_Cum_Qty"]       = df_rev["Cum_Qty"].fillna(df_rev["Path_Cum_Qty"])
        if "Extended_Cost" in df_rev.columns: df_rev["Path_Extended_Cost"] = df_rev["Extended_Cost"].fillna(0.0)
        if "UoM"           in df_rev.columns: df_rev["UoM"]                = df_rev["UoM"].fillna(df_rev.get("UoM_old", df_rev["UoM"]))
            
        drop_cols = ["Final_Product", "Raw_Material", "Cum_Qty", "Extended_Cost", "UoM_old"]

        # تجنّب حذف Parent_Material إذا أتى من الدمج — فقط الأعمدة المكررة الزائدة

        if _has_pm and "Parent_Material_old" in df_rev.columns:
            drop_cols.append("Parent_Material_old")
        df_rev.drop(columns=[c for c in drop_cols if c in df_rev.columns], inplace=True)

    # ── الأعمدة الثابتة بالترتيب المنطقي — Parent_Material بعد بيانات المنتج التام ──
    fixed_cols = [
        "Raw_Material_Code", "Path_Cum_Qty", "UoM", "Unit_Price_EGP",
        "Path_Extended_Cost", "Final_Product_Code", "Final_Product_Name",
        "Parent_Material", "Parent_Material_Name",   # ✅ Level-1 SF
    ]
    
    # تحديد الأعمدة الديناميكية (أعمدة المستويات)
    dynamic_cols = [c for c in df_rev.columns if c not in fixed_cols]
    
    # دالة ترتيب آمنة ومحمية بالكامل من الحروف والنصوص الزائدة
    def safe_sort_key(col_name):
        parts = col_name.split("_")
        last_part = parts[-1]
        # التحقق إذا كان الجزء الأخير رقماً فعلياً مثل (1, 2, 3)
        if last_part.isdigit():
            level_num = int(last_part)
        else:
            level_num = 999 # لو نصي نضعه في النهاية لمنع انهيار البرنامج
            
        is_name = 1 if "Name" in col_name else 0
        return (level_num, is_name)
        
    dynamic_cols.sort(key=safe_sort_key)
    
    return df_rev[fixed_cols + dynamic_cols].drop_duplicates().reset_index(drop=True)

# ==============================================================================
# 4. واجهة المستخدم (الممتدة لليمين)
# ==========================================================

# 1️⃣ إعدادات الصفحة الرئيسية للبرنامج (جعل الواجهة عريضة)
st.set_page_config(page_title="💪🔥 MRP Tool", page_icon="👍", layout="wide")

# 2️⃣ حقن أكواد التنسيق CSS المتطورة والشاملة لضبط الألوان والمحاذاة والاتجاهات
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;700&display=swap');
    

    /* ✅ إجبار كل العناصر الداخلية تختفي */
    section[data-testid="stSidebar"][aria-expanded="false"] * {
        display: none !important;
    }
    
    /* 🔀 1. القلاب العام: نقل جسم التطبيق والشريط الجانبي بالكامل إلى اليمين فوراً */
    .stApp, [data-testid="stSidebar"] {
        direction: rtl !important;
    }

    /* 📝 2. تطبيق خط القاهرة وتوجيه النصوص والعناوين لليمين في كل الشاشات */
    html, body, .stMarkdown, p, h1, h2, h3, h4, h5, h6, label {
        font-family: 'Cairo', sans-serif !important;
        text-align: right !important;
        direction: rtl !important;}
    }


    /* 📐 3. تحسين عرض الجداول وإزالة الهوامش الزائدة لتوفير المساحة */
    .custom-container {
        width: 100%;
        margin: 0 auto;
    }
    .main > div {
        padding-top: 1rem;
    }

    /* 🗂️ 4. تنسيق الـ Tabs المتقدم ليمتد بكامل عرض الشاشة (100%) مطابق للصورة */
    div[data-testid="stTabs"] {
        width: 100% !important;
    }
    .stTabs [data-baseweb="tab-list"] {
        direction: rtl !important;
        justify-content: flex-start !important;

        padding: 5px 10px 0px 10px !important;
        border-radius: 8px 8px 0px 0px !important;
        border-bottom: 3px solid #1f4068 !important; /* الخط السفلي الأزرق الممتد */
        gap: 6px !important;
        width: 100% !important;
    }
    .stTabs [data-baseweb="tab"] {
        font-family: 'Cairo', sans-serif !important;
        font-weight: bold !important;
        color: #5f6368 !important;
        background-color: #e8eaed !important; /* لون التبويب غير النشط */
        padding: 10px 25px !important;
        border-radius: 6px 6px 0px 0px !important;
        border: none !important;
        margin-bottom: -3px !important;
        transition: all 0.2s ease-in-out !important;
    }
    .stTabs [data-baseweb="tab"]:hover {
        background-color: #d2d6dc !important;
        color: #1f4068 !important;
    }
    .stTabs [data-baseweb="tab"][aria-selected="true"] {
        font-family: 'Cairo', sans-serif !important;
        font-weight: bold !important;
        background-color: #1f4068 !important; /* لون التبويب النشط متناسق مع صندوق العنوان */
        color: white !important;
        border-bottom: none !important;
        box-shadow: 0px -2px 5px rgba(0,0,0,0.05) !important;
    }
    .stTabs [data-baseweb="tab-border"] {
        background-color: transparent !important;
        width: 100% !important;
    }
    .stTabs [data-baseweb="tab-panel"] {
        direction: rtl !important;
        text-align: right !important;
        padding-top: 15px !important;
        width: 100% !important;
    }

    /* ⚙️ 5. محاذاة أدوات الإدخال الرقمية والكتابة داخل الشريط الجانبي (Sidebar) */
    section[data-testid="stSidebar"] input {
        text-align: left  !important;
        direction: ltr !important;
    }
    section[data-testid="stSidebar"] div[data-testid="stNumberInput"] input {
        text-align: left  !important;
        direction: ltr !important;
    }

    /* label عربي */
    [data-testid="stSlider"] label {
        direction: rtl !important;
        text-align: right !important;
    }    

    /* السلايدر نفسه */
    [data-testid="stSlider"] {
        direction: ltr !important;
    }

    /* 📘 6. تنسيق صندوق الـ Expander المخصص لدليل الاستخدام */
    .stExpander {
        direction: rtl !important;
        text-align: right !important;
    }
    
    /* 🌌 7. تنسيق الصندوق المتدرج الفخم للعنوان الرئيسي (إصدار v3 المتطور) */
    .main-title {
        background: linear-gradient(135deg, #1f4068, #162447);
        color: white;
        padding: 20px;
        border-radius: 10px;
        text-align: center !important;
        margin-bottom: 25px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }

    /* 🧹 حماية الفلاتر من الألوان الغلط */
    [data-baseweb="select"],
    [data-baseweb="tag"],
    div[role="option"] {
        color: var(--app-text) !important;
    }

    /* tags داخل multiselect */
    span[data-baseweb="tag"] {
        background-color: var(--app-accent) !important;
        color: white !important;
        border-radius: 6px !important;
    }

    /* 🔥 لون الحظ داخل الزر */
    .stButton button *,
    .stDownloadButton button * {
        color: #ffffff !important;
        font-weight: 700 !important;
        opacity: 1 !important;
        font-family: 'Cairo', sans-serif !important;
    }

    /* fallback إضافي قوي */
    .stButton button p,
    .stButton button span {
        color: #ffffff !important;
    }

    /* النص داخل الفلتر */
    [data-baseweb="select"] span {color: white !important;font-weight: bold !important;}


    .main-title h1,.main-title h2,.main-title h3,.main-title p,.main-title span {color: white !important;}
    </style>
    """,
    unsafe_allow_html=True)

# 3️⃣ استدعاء الصندوق المتدرج وعرض العنوان الرئيسي الفخم بمنتصف الشاشة
st.markdown(
    """
    <div class="main-title">
        <h2 style="margin:0; padding:0; color:white !important; font-family:'Cairo', sans-serif; font-weight: 700; font-size: 26px;">
            🔥💪 نظام تخطيط الاحتياجات والتحليلات المالية 📂 
            <span style="font-size: 14px; font-weight: 400; color: #ffd700; float: left; padding-top: 8px; font-family:'Cairo', sans-serif;">
                (تصميم م. رضا رشدي)
            </span>
        </h2>
    </div>
    """, 
    unsafe_allow_html=True)


# 4️⃣ دليل الاستخدام المتجاور والمريح للعين (يفتح من اليمين - نسخة الدعم الفني)
st.markdown(
    """
    <style>
    /* استهداف نص الـ expander وتكبيره إلى 24px واختيار الخط المعتمد */
    .stExpander summary p {
        font-size: 24px !important;
        font-weight: bold !important;
        font-family: 'Cairo', sans-serif !important;
        color: #162447 !important;
    }

    </style>
    """,
    unsafe_allow_html=True)

# V4: تطبيق الثيم الديناميكي + تحسين الاستجابة للموبايل
apply_app_theme()

# 👇 هنا مكان إضافة CSS الخاص بالـ Tabs
st.markdown(
    """
    <style>
    /* لون أسماء التبويبات */
    button[data-baseweb="tab"] {
        color: white !important;
        font-weight: bold !important;
    }

    /* التبويب النشط */
    button[data-baseweb="tab"][aria-selected="true"] {
        color: white !important;
        background-color: rgba(255,255,255,0.10) !important;
        border-radius: 6px 6px 0 0 !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)


with st.expander("📘 دليل استخدام النظام وهيكل الملفات", expanded=False):

    guide_tab1, guide_tab2, guide_tab3 = st.tabs(["📋 شروط الملفات", "💰 التحليل المالي", "📞 الدعم والتواصل"])
    with guide_tab1:
        st.markdown(
            """
            <div style="direction: rtl; text-align: right; font-size: 20px; line-height: 1.7; display: grid; grid-template-columns: 1fr 1fr; gap: 15px; font-family:'Cairo', sans-serif;">
                <div style="background-color: #f8f9fa; padding: 15px; border-radius: 6px; border-right: 5px solid #1976d2;">
                    <b style="font-size: 22px; display: block; margin-bottom: 8px;">🗓️ ورقة plan (الخطة):</b>
                    • <code style="color:#1976d2; font-weight:bold; font-size: 18px; direction: ltr; display: inline-block;">Material</code>: كود المنتج النهائي.<br>
                    • <code style="color:#1976d2; font-weight:bold; font-size: 18px; direction: ltr; display: inline-block;">Order Type</code>: نوع الأمر (E تصدير / L محلي).<br>
                    • <code style="color:#1976d2; font-weight:bold; font-size: 18px; direction: ltr; display: inline-block;">Date/Qty</code>:اعمدة التواريخ (mm/dd/yyyy)-تحتوى على الكميات المخططة..<br>
                    • <code style="color:#1976d2; font-weight:bold; font-size: 18px; direction: ltr; display: inline-block;">**********</code>:ملف تجريبى :</b><a href="https://github.com/Era202/Final_MRP-Product_Costing/raw/main/assets/MRP_Template.xlsx"
                       target="_blank"style="color:#1976d2; text-decoration:none; font-weight:bold;">📥 تحميل نموذج Excel للتجربة
                    </a><br>

            </div>
                <div style="background-color: #f8f9fa; padding: 15px; border-radius: 6px; border-right: 5px solid #2e7d32;">
                    <b style="font-size: 22px; display: block; margin-bottom: 8px;">🔩 Component ورقة("ZMAT_BOM_LEVEL" TCode in SAP):</b>
                    • <code style="color:#2e7d32; font-weight:bold; font-size: 18px; direction: ltr; display: inline-block;">Material</code>: كود الجذر | <code style="color:#2e7d32; font-weight:bold; font-size: 18px; direction: ltr; display: inline-block;">Parent</code>: الأب الفعلي.<br>
                    • <code style="color:#2e7d32; font-weight:bold; font-size: 18px; direction: ltr; display: inline-block;">Component</code>: كود M المكون أو الخامة.<br>
                    • <code style="color:#2e7d32; font-weight:bold; font-size: 18px; direction: ltr; display: inline-block;">Hierarchy Level</code>: المستوى الهرمي.<br>
                    • <code style="color:#2e7d32; font-weight:bold; font-size: 18px; direction: ltr; display: inline-block;">Order Type</code>: توفير (F شراء / E تصنيع).
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )
        
    with guide_tab2:
        st.markdown(
            """
            <div style="direction: rtl; text-align: right; font-size: 20px; line-height: 1.7; display: grid; grid-template-columns: 1fr 1fr; gap: 15px; font-family:'Cairo', sans-serif;">
                <div style="background-color: #f0f4f8; padding: 15px; border-radius: 6px; border-right: 5px solid #0288d1;">
                    <b style="font-size: 22px; display: block; margin-bottom: 8px;">1️⃣ نظام التسعير:</b>
                    • <b>السعر التلقائي:</b> يعتمد افتراضياً على عمود <code style="color:#1976d2; font-weight:bold; font-size: 18px; direction: ltr; display: inline-block;">Last Price</code> في المكونات.<br>
                    • <b>ملف جديد:</b> تفعيل خيار "ملف أسعار جديد" لرفع ملف خارجي.
                </div>
                <div style="background-color: #f0f4f8; padding: 15px; border-radius: 6px; border-right: 5px solid #f57c00;">
                    <b style="font-size: 22px; display: block; margin-bottom: 8px;">2️⃣ أدوات المحاكاة:</b>
                    • <b>المحاكاة المالية:</b> تعديل سعر الصرف ونسبة الزيادة لمحاكاة التضخم.<br>
                    • <b>الفلترة السريعة:</b> الفلتر العام يعزل منتجاً واحداً لتحديث رسوماته فوراً.
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with guide_tab3:
        st.markdown(
            """
            <div style="direction: rtl; text-align: right; font-size: 20px; line-height: 1.8; display: grid; grid-template-columns: 1fr; gap: 15px; font-family:'Cairo', sans-serif;">
                <div style="background-color: #f9f6f0; padding: 20px; border-radius: 8px; border-right: 5px solid #d4af37; box-shadow: 0 2px 4px rgba(0,0,0,0.02);">
                    <b style="font-size: 24px; color: #162447; display: block; margin-bottom: 12px;">1️⃣ إدارة تطوير النظام:</b>
                    • 💻 <b> المصمم والمطور مهندس /  رضا رشدى </b> <br>
                    • ✉️ <b>البريد الإلكتروني:</b> <a href="mailto:reda.roshdy@fresh.com.eg" style="color:#1976d2; text-decoration:none; font-weight:bold; direction:ltr; display:inline-block;">reda.roshdy@fresh.com.eg</a><br>
                    • 📞 <b>رقم الهاتف / الواتساب:</b> <span style="direction:ltr; display:inline-block; font-weight:bold; color:#2e7d32;">+20 01271702820 </span><br>
                    • 🏢 <b>ملف تجريبى:</b><a href="https://github.com/Era202/Final_MRP-Product_Costing/raw/main/assets/MRP_Template.xlsx"
                       target="_blank"style="color:#1976d2; text-decoration:none; font-weight:bold;">📥 تحميل نموذج Excel للتجربة
                    </a><br>
                    <span style="color: white !important; font-weight: bold !important; font-size: 16px; display:block; margin-top: 10px; font-style: italic;">💡 يسعدنا تواصلكم لأي استفسارات تقنية، تحديثات في كود الـ MRP، أو تخصيص تحليلات مالية إضافية.</span>
                </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

# ==============================================================================
# Sidebar — رفع الملفات واستقرار جهة اليمين كاملاً
# ==============================================================================
with st.sidebar:
    st.markdown("### إعدادات العرض والحساب", unsafe_allow_html=True)
    st.checkbox("تفعيل الوضع الليلي", key="dark_mode")
    st.number_input("أقصى عمق BOM", min_value=5, max_value=20, step=5, key="max_bom_level", help="استخدم قيمة أعلى إذا كانت شجرة BOM لديك متعددة المستويات بشكل عميق.")
    st.markdown("---")

    st.markdown("<h3 style='text-align:right; direction:rtl;'>📂 اختر ملف الخطة الشهرية Excel</h3>", unsafe_allow_html=True)
    
    # 2. أداة رفع الملف النظيفة بدون تسمية مكررة مشوهة
    uploaded_file = st.file_uploader(
        "ملف الخطة", 
        type=["xlsx"],
        label_visibility="collapsed"
    )
    
    # حارس الأخطاء: يوقف التنفيذ حتى يتم رفع الملف
    if not uploaded_file:
        st.stop()

    # 📌 وضع هذا البلوك في نهاية الـ sidebar تماماً ليعمل كـ Footer ثابت

# ==============================================================================
# Sidebar — ملف الأسعار + محاكي السيناريوهات
# ==============================================================================
    # 3. خط الإعدادات المالي الحالي الخاص بك
  #  st.markdown("<h3 style='text-align:right; direction:rtl;'>⚙️ إعدادات التحليل المالي</h3>", unsafe_allow_html=True)

    #st.markdown("---")

# ── رفع ملف الأسعار ─────────────────────────────────────────────────────
    
    # حقن تنسيق خاص بـ st.radio ليصبح اتجاهه من اليمين إلى اليسار ومحاذاة النصوص لليمين
    st.markdown(
        """
        <style>

        /* ✅ إجبار كل العناصر الداخلية تختفي */
        section[data-testid="stSidebar"][aria-expanded="false"] * {
            display: none !important;
        }

        /* ✅ استهداف sidebar فقط */
        [data-testid="stSidebarContent"] [data-baseweb="select"],
        [data-testid="stSidebarContent"] [data-baseweb="input"] {
            overflow: visible !important;
        }

        /* ضبط اتجاه عناصر الراديو والنصوص التابعة لها */
        div[data-testid="stRadio"] {
            direction: rtl !important;
            text-align: right !important;
        }
        div[data-testid="stRadio"] label {
            text-align: right !important;
            direction: rtl !important;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    # عنوان القسم بنفس خط وحجم الخطة الشهرية والإعدادات
    st.markdown("<h3 style='text-align:right; direction:rtl;'>💰 ملف الأسعار (Price Master)</h3>", unsafe_allow_html=True)

    # أداة الاختيار (الراديو) مع جعل النص التوضيحي متوافقاً مع اليمين
    price_source = st.radio(
        "مصدر الأسعار الحالي:",
        options=["last_price", "price_file"],
        format_func=lambda x: "📋 عمود Last Price من ملف المكونات" if x == "last_price" else "📂 ملف أسعار جديد",
        key="price_source_radio"
    )

    price_file = None
    
    # الحالة الأولى (الافتراضية الحالية): استخدام السعر من ملف المكونات مباشرة
    if price_source == "last_price":
        financial_active = True
        st.success("✅ تم تفعيل التحليل المالي تلقائياً باستخدام عمود Last Price من ملف المكونات.")
        
    # الحالة الثانية: إذا اختار المستخدم رفع ملف جديد، ينتظر البرنامج رفع الملف
    else:
        # عنوان مخصص ومحاذٍ لليمين لأداة رفع الأسعار متناسق مع بقية العناوين
        st.markdown("<p style='text-align:right; direction:rtl; font-weight:bold; font-size:15px; margin-bottom:5px;'>📂 ارفع ملف الأسعار الجديد</p>", unsafe_allow_html=True)
        
        price_file = st.file_uploader(
            "ملف الأسعار", # نص داخلي مخفي
            type=["xlsx"],
            key="price_file",
            help="يجب أن يحتوي على عمود Component/كود + عمود Price/السعر",
            label_visibility="collapsed" # إخفاء لتجنب كلمة None واللخبطة اليسارية
        )
        
        financial_active = price_file is not None
        if financial_active:
            st.success("✅ ملف الأسعار الجديد محمّل — التحليل المالي مفعّل بناءً عليه.")
        else:
            st.info("⏳ في انتظار رفع ملف الأسعار الجديد... التبويب المالي معلق لحين الرفع.")
            st.stop()  #    الحارس لمنع ظهور أخطاء أسفل الصفحة لجين رفع  ملف السعر والاستكمال  

    st.markdown("---")

    # ── محاكي السيناريوهات (What-if) ─────────────────────────────────────────
    st.markdown("### 📊 محاكي السيناريوهات")
    exchange_rate = st.number_input(
        "💱 سعر الصرف (جنيه / دولار)",
        min_value=1.0, max_value=1000.0,
        value=1.0, step=0.5,
        help="يُستخدم لتحويل التكاليف إلى الجنيه المصري"
    )

    price_increase_pct = st.slider(
        "📈 نسبة زيادة أسعار الخامات (%)",
        min_value=0, max_value=100,
        value=0, step=5,
        help="تطبيق نسبة زيادة افتراضية على جميع الأسعار"
    )
    price_multiplier = 1 + (price_increase_pct / 100)

    if price_increase_pct > 0:
        st.warning(f"⚠️ الأسعار مضروبة × {price_multiplier:.2f} (زيادة {price_increase_pct}%)")


    st.markdown("---")
    if st.button("🗑️ أمسح الذاكرة المؤقتة  (Clear Cache)", use_container_width=True):
        st.cache_data.clear()
        st.success("تم مسح الكاش، سيتم إعادة حساب كل شيء عند الحاجة.")

    st.markdown(
        """
        <hr>
        <div style="text-align:center;padding:10px;border-radius:10px;background-color:#f0f2f6;margin-top:10px;">
            <b> تنفيذ مهندس/ رضا رشدي ✨</b><br>
            <span style="color:gray;">© 2026 جميع الحقوق محفوظة</span></div>""",unsafe_allow_html=True)

# ==============================================================================
# --- تحميل البيانات ---
# ==============================================================================
with st.spinner("جاري تحميل وتطهير البيانات..."):

    plan_df, component_df, mrp_df = load_and_validate_data(uploaded_file)
    plan_df_orig      = plan_df.copy()
    component_df_orig = component_df.copy()
    mrp_df_orig       = mrp_df.copy()

# --- استخراج أعمدة التواريخ ---
date_cols = [c for c in plan_df.columns if isinstance(c, (datetime.datetime, pd.Timestamp))]

if not date_cols:
    st.error("❌ لم يتم العثور على أعمدة تواريخ في ورقة الخطة.")
    st.stop()

# ==============================================================================
# ✅ معالجة Order Type وتطهير الفراغات قبل الـ Melt (حماية من تشوه إحصائيات الداتا)
# ==============================================================================
missing_ot = 0
if "warnings_list" not in st.session_state:
    st.session_state.warnings_list = []

# 1. إذا كان العمود غير موجود نهائياً في البيانات
if col("order_type") not in plan_df.columns:
    plan_df[col("order_type")] = "L"
    missing_ot = len(plan_df)
    st.session_state.warnings_list.append("⚠️ لم يتم العثور على عمود Order Type — تم إضافة القيمة الافتراضية 'L' لكافة السطور")

# 2. العمود موجود، نقوم بتطهيره وحساب الفراغات الفعلية فقط دون مساس بالقيم الأصلية
else:
    # تنظيف مبدئي للمسافات وتحويل نصوص الفراغات المضللة
    temp_series = plan_df[col("order_type")].astype(str).str.strip()
    is_missing = (
        plan_df[col("order_type")].isna() | 
        (temp_series == "") | 
        (temp_series.str.lower() == "nan") |  # 👈 إضافة .str هنا
        (temp_series.str.lower() == "none")   # 👈 إضافة .str هنا
    )
    
    # حساب العدد الحقيقي للسطور الفارغة فعلياً
    missing_ot = int(is_missing.sum())
    
    # استبدال الفراغات فقط بـ 'L' والحفاظ على الـ F والـ E السليمة
    plan_df[col("order_type")] = plan_df[col("order_type")].where(~is_missing, "L")
    plan_df[col("order_type")] = plan_df[col("order_type")].astype(str).str.strip()

    if missing_ot > 0:
        st.session_state.warnings_list.append(f"⚠️ تم تعويض {missing_ot} قيمة فارغة فقط في عمود Order Type - تم استبدالها بـ 'L'")
    # ==============================================================================
    # --- استخراج أعمدة التواريخ ---
    # ==============================================================================
date_cols = [c for c in plan_df.columns if isinstance(c, (datetime.datetime, pd.Timestamp))]

if not date_cols:
    st.error("❌ لم يتم العثور على أعمدة تواريخ في ورقة الخطة.")
    st.stop()

with st.spinner("جاري حساب الاحتياجات (MRP)⏳..."):

    # ==============================================================================
      # A. تجهيز الخطة (Melt)
    # ==============================================================================
    start_time = time.time()
    plan_melted = plan_df.melt(
        id_vars=[col("material"), col("material_desc"), col("order_type")],
        value_vars=date_cols,
        var_name="Date",
        value_name="Planned Quantity"
    )
    plan_melted["Date"] = pd.to_datetime(plan_melted["Date"], errors='coerce')
    plan_melted["Planned Quantity"] = pd.to_numeric(
        plan_melted["Planned Quantity"], errors="coerce"
    ).fillna(0)
    
    # إزالة الصفوف بكمية صفر أو تاريخ مجهول
    plan_melted = plan_melted[
        (plan_melted["Planned Quantity"] > 0) &
        (plan_melted["Date"].notna())
    ].copy()

with st.spinner("جاري حساب الاحتياج لكل منتج..."):

    # ==============================================================================
    # ✅ B. تشغيل Multi-Level BOM Explosion
    # ==============================================================================
#    st.markdown("---")
#    st.subheader("🔩 نتائج BOM Explosion — جميع المستويات الهرمية")

    result_df = bom_explosion(plan_melted, component_df)
with st.spinner("جاري حساب التكاليف المالية..."):

    # ==============================================================================
    # B2. المحرك المالي — Financial Engine (Bottom-Up Leaf-Only Costing)
    # ==============================================================================

    price_df         = pd.DataFrame()
    product_costing  = pd.DataFrame()
    ctrl_financial   = pd.DataFrame()
    sensitivity_df   = pd.DataFrame()
    debug_costing_df = pd.DataFrame()
    financial_ready = True

    # ✅ FIX: استخدام generate_bom_paths (تُنتج Level_N/Name_N) بدلاً من build_bom_paths
    try:
        df_bom_paths = generate_bom_paths(component_df, plan_df)
    except Exception as e:
        st.error(f"BOM Paths Error: {e}")
        df_bom_paths = pd.DataFrame()
    # ==================كشف الاخطاء فى البومات بسبب الواحدات ========
      # Component UOM Issues (Analytical – detect WRONG BOMs only)
    # Rule: Most frequent UOM = Standard / Correct
    # ==========================

    # 1) Count usage of each UOM per Component
    uom_usage = (
        component_df
        .groupby([col("component"), col("component_desc"), col("component_uom")])
        .size()
        .reset_index(name="uom_count")
    )

    # 2) Determine STANDARD UOM (highest frequency)
    standard_uom = (
        uom_usage
        .sort_values("uom_count", ascending=False)
        .groupby(col("component"))
        .first()
        .reset_index()
        .rename(columns={
            col("component_uom"): "Standard_UOM",
            "uom_count": "Standard_Count"
        })
    )

    # 3) Merge Standard UOM back to component master
    wrong_uom_rows = (
        component_df
        .merge(
            standard_uom[[col("component"), "Standard_UOM"]],
            on=col("component"),
            how="left"
        )
    )

    # 4) Identify ONLY wrong UOM rows
    wrong_uom_rows = wrong_uom_rows[
        wrong_uom_rows[col("component_uom")] != wrong_uom_rows["Standard_UOM"]
    ]

    # 5) Final analytical sheet (WRONG BOMs only)
    component_uom_issues_sheet = (
        wrong_uom_rows
        .groupby(
            [col("component"), col("component_desc"), "Standard_UOM", col("component_uom")]
        )
        .agg(
            Wrong_BOMs=(
                col("material"),
                lambda x: ", ".join(sorted(map(str, x.dropna().unique())))
            )
        )
        .reset_index()
        .rename(columns={
            col("component_uom"): "Wrong_UOM"
        })
    )

if financial_active and not df_bom_paths.empty:
    try:
        # ── 1. تحديد مصدر الأسعار ────────────────────────────────────────
        price_map = {}

        if price_source == "price_file" and price_file is not None:
            price_raw = pd.read_excel(price_file, sheet_name=0)
            price_raw = normalize_columns(price_raw, COLUMN_NAMES)
            _pcode_col = None
            for _alias in COLUMN_NAMES["component"] + COLUMN_NAMES["material"]:
                if _alias in price_raw.columns:
                    _pcode_col = _alias
                    break
            _pval_col = None
            for _alias in COLUMN_NAMES["price"]:
                if _alias in price_raw.columns:
                    _pval_col = _alias
                    break
            if not _pcode_col or not _pval_col:
                st.sidebar.error("❌ لم يُعثر على أعمدة الكود أو السعر في ملف الأسعار — تأكد من وجود عمود Component/كود وعمود Price/السعر")
                financial_ready = False
            else:
                _pdf = price_raw[[_pcode_col, _pval_col]].drop_duplicates(subset=[_pcode_col])
                _pdf[_pcode_col] = _pdf[_pcode_col].astype(str).str.strip()
                _pdf[_pval_col]  = pd.to_numeric(_pdf[_pval_col], errors="coerce").fillna(0)
                price_map = dict(zip(_pdf[_pcode_col],
                                     _pdf[_pval_col] * price_multiplier * exchange_rate))
        elif price_source == "last_price":
            _lp_col = None
            for _alias in COLUMN_NAMES["price"]:
                if _alias in component_df_orig.columns:
                    _lp_col = _alias
                    break
            if _lp_col:
                _lpdf = component_df_orig[[col("component"), _lp_col]].drop_duplicates(subset=[col("component")])
                _lpdf[col("component")] = _lpdf[col("component")].astype(str).str.strip()
                _lpdf[_lp_col] = pd.to_numeric(_lpdf[_lp_col], errors="coerce").fillna(0)
                price_map = dict(zip(_lpdf[col("component")],
                                     _lpdf[_lp_col] * price_multiplier * exchange_rate))
            else:
                st.sidebar.warning("⚠️ لم يُعثر على عمود Last Price في ملف المكونات")
                financial_active = False

        if financial_ready and price_map:

            # ── 2. تحديد Leaf Nodes بشكل صحيح ─────────────────────────────
            parent_col_leaf = col("parent_material") if col("parent_material") in component_df.columns else col("material")
            all_parents    = set(component_df[parent_col_leaf].astype(str).str.strip())
            all_components = set(component_df[col("component")].astype(str).str.strip())
            leaf_codes      = all_components - all_parents   # فقط المكونات الطرفية

            # ── 3. بناء قاموس شامل متكامل للأوصاف ومسؤولي الـ MRP لضمان عدم الفقد ──
            # نجمع البيانات من عمود المكونات وعمود المواد التامة معاً في قاعدة بيانات موحدة

            full_meta = pd.concat([
                component_df_orig[[col("component"), col("component_desc"), col("mrp_controller")]].rename(
                    columns={col("component"): "code", col("component_desc"): "desc", col("mrp_controller"): "mrp"}),
                component_df_orig[[col("material"), col("component_desc"), col("mrp_controller")]].rename(
                    columns={col("material"): "code", col("component_desc"): "desc", col("mrp_controller"): "mrp"})
            ]).drop_duplicates(subset=["code"], keep="first")
            
            full_meta["code"] = full_meta["code"].astype(str).str.strip()
            _desc_map = dict(zip(full_meta["code"], full_meta["desc"].fillna("")))
            _mrp_map = dict(zip(full_meta["code"], full_meta["mrp"].fillna("غير محدد")))

            # ── 4. بناء جدول الحساب من BOM_Paths ─────────────────────────
        with st.spinner(" BOM_Paths جاري بناء  مسارات BOM..."):

            # ✅ Unified Description Map (بدون تغيير أسماء)
            def _norm(x):
                return str(x).strip().upper()

            # استخدام نفس الاسم السابق لضمان عدم كسر الكود
            _desc_map = {}

            # من component
            for c, d in zip(
                component_df_orig[col("component")],
                component_df_orig[col("component_desc")].fillna("")
            ):
                key = _norm(c)
                if key and d:
                    _desc_map[key] = d

            # دعم إضافي من material
            for m, d in zip(
                component_df_orig[col("material")],
                component_df_orig[col("component_desc")].fillna("")
            ):
                key = _norm(m)
                if key and d and key not in _desc_map:
                    _desc_map[key] = d

            _level_cols = [c for c in df_bom_paths.columns if c.startswith("Level_")]
            _name_cols  = [c for c in df_bom_paths.columns if c.startswith("Name_")]

            _costing_rows = []

            for _, path_row in df_bom_paths.iterrows():
#*********************************************************************
           # for path_row in df_bom_paths.itertuples(index=False):
#*********************************************************************
                root = str(path_row.get("Level_1", "")).strip()
                if not root:
                    continue

                # ── جمع كل المستويات الممتلئة بالترتيب الأمامي ────────────────────────
                # (أكثر وضوحاً وأمناً من البحث المعكوس ويتيح تطبيق قاعدة Level-1 SF)
                actual_levels     = []   # كودات المستويات بالترتيب
                actual_name_cols  = []   # أعمدة الأسماء المقابلة

                for _lc, _nc in zip(_level_cols, _name_cols):
                    _v = str(path_row.get(_lc, "")).strip()
                    if _v and _v != "nan":
                        actual_levels.append(_v)
                        actual_name_cols.append(_nc)
                    else:
                        break   # المستويات متسلسلة — أول فراغ = نهاية المسار

                # مسار ناقص (أقل من 2 مستوى) → تجاهل
                if len(actual_levels) < 2:
                    continue

                leaf_code     = actual_levels[-1]        # آخر مستوى = الخامة
                leaf_name_col = actual_name_cols[-1]

                if leaf_code == root or leaf_code not in leaf_codes:
                    continue

                # ── قاعدة Level-1 SF لـ Parent_Material ────────────────────────────
                # المطلوب المحاسبي: أول مكون تحت المنتج النهائي مباشرة (Level_2)
                # وليس الأب المباشر للخامة — حتى لو كانت الخامة في عمق Level_6
                #  FG → B → C → D → X  =>  Parent_Material = B  (actual_levels[1])
                #  FG → X               =>  Parent_Material = FG (actual_levels[0])
                #                            الخامة مرتبطة بالمنتج التام مباشرة
                if len(actual_levels) == 2:
                    # خامة مباشرة تحت المنتج التام — لا يوجد SF بينهما
                    parent_material = actual_levels[0]          # = root / FG
                else:
                    # يوجد على الأقل مكون وسيط واحد → نأخذ الأول (Level_2)
                    parent_material = actual_levels[1]

                _raw_name = str(path_row.get(leaf_name_col, ""))
                cum_qty_f = 1.0
                _uom = ""
                if " , " in _raw_name:
                    _qty_part = _raw_name.split(" , ", 1)[1].strip()
                    _parts = _qty_part.split()
                    try:
                        cum_qty_f = float(_parts[0])
                        _uom = _parts[1] if len(_parts) > 1 else ""
                    except (ValueError, IndexError):
                        pass

                _uom_upper = _uom.upper()
                if _uom_upper in ("G", "GM", "GRAM"):
                    cum_qty_f /= 1000
                    _uom = "KG"
                elif _uom_upper in ("CM2", "CM²"):
                    cum_qty_f /= 10000
                    _uom = "M2"

                unit_price = price_map.get(leaf_code, 0.0)
                ext_cost   = cum_qty_f * unit_price

                # ✅ تحديد وصف Parent بشكل ذكي بناءً على الحالة جلب البيانات 

                if parent_material == root:
                    # حالة: الخامة مباشرة تحت المنتج التام
                    # نستخدم وصف من Name_2 (المكون المباشر)
                    if len(actual_name_cols) > 1:
                        parent_desc_val = str(path_row.get(actual_name_cols[1], "")).split(" , ", 1)[0].strip()
                    else:
                        parent_desc_val = _desc_map.get(_norm(parent_material), "")
                else:
                    # الحالة العادية
                    parent_desc_val = _desc_map.get(_norm(parent_material), "")

                # ✅ Raw Material Description (بدون تأثير على Parent)
                rm_desc_val = _desc_map.get(_norm(leaf_code), "")

                if not rm_desc_val and " , " in _raw_name:
                    rm_desc_val = _raw_name.split(" , ", 1)[0].strip()

                _costing_rows.append({
                    "Final_Product":   root,
                    "Parent_Material": parent_material,
                    "Parent_Material_Description": parent_desc_val,  # ✅ مهم جدًا
                    "Raw_Material":    leaf_code,
                    "RM_Description":  rm_desc_val,
                    "MRP_Controller":  _mrp_map.get(leaf_code, "غير محدد"),
                    "Cum_Qty":         cum_qty_f,
                    "UoM":             _uom,
                    "Unit_Price_EGP":  unit_price,
                    "Extended_Cost":   ext_cost,
                    })

            if not _costing_rows:
                st.sidebar.warning("⚠️ لا توجد خامات مطابقة بين Leaf Nodes وملف الأسعار")
                financial_active = False
            else:
                debug_costing_df = pd.DataFrame(_costing_rows)
                # تنظيف نصوص الأكواد لضمان الدمج الخالي من الفراغات
                debug_costing_df["Final_Product"]   = debug_costing_df["Final_Product"].astype(str).str.strip()
                debug_costing_df["Parent_Material"] = debug_costing_df["Parent_Material"].astype(str).str.strip()
                

                # ✅ ترتيب الأعمدة: المنتج التام → أول SF تحته → الخامة → البيانات المالية 1598 Direct parent for parent
                _desired_col_order = [
                    "Final_Product", "Parent_Material",  "Parent_Material_Description", "Raw_Material",
                    "RM_Description", "MRP_Controller",
                    "Cum_Qty", "UoM", "Unit_Price_EGP", "Extended_Cost",
                ]
                # الاحتفاظ بأي أعمدة إضافية غير متوقعة في النهاية
                _extra_debug_cols = [c for c in debug_costing_df.columns if c not in _desired_col_order]
                debug_costing_df  = debug_costing_df[_desired_col_order + _extra_debug_cols]

                # ── 5. ورقة تكلية المنتج النهائي Product_Unit_Costing ───────────────────────────────
                _prod_cost = (
                    debug_costing_df
                    .groupby("Final_Product", as_index=False)
                    .agg(
                        Raw_Materials_Count = ("Raw_Material", "nunique"),
                        Total_Cost_EGP = ("Extended_Cost", "sum"),
                    )
                )
                
                # تأمين خرائط أسماء ووصف المنتجات التامة من ورقة الخطة وورقة المكونات معاً لعدم السقوط
                plan_df_orig[col("material")] = plan_df_orig[col("material")].astype(str).str.strip()
                _mat_desc_map = dict(zip(plan_df_orig[col("material")], plan_df_orig[col("material_desc")].fillna("")))
                
                # دعم إضافي من ورقة المكونات لوصف المنتج التام إذا لم يوجد بالخطة
                for _m, _d in zip(component_df_orig[col("material")]
                .astype(str).str.strip(), component_df_orig[col("component_desc")].fillna("")):
                    if _m not in _mat_desc_map or not _mat_desc_map[_m]:
                        _mat_desc_map[_m] = _d

                _plan_qty_map = (
                    plan_df_orig
                    .set_index(col("material"))[date_cols]
                    .sum(axis=1)
                    .to_dict()
                )

                _prod_cost["Material_Desc"] = _prod_cost["Final_Product"].map(_mat_desc_map).fillna("")
                _prod_cost["Plan_Total_Qty"] = _prod_cost["Final_Product"].map(_plan_qty_map).fillna(0)
                _prod_cost["Total_Plan_Cost"] = _prod_cost["Total_Cost_EGP"] * _prod_cost["Plan_Total_Qty"]
                
                product_costing = _prod_cost[[
                    "Final_Product", "Material_Desc", "Raw_Materials_Count", "Total_Cost_EGP", "Plan_Total_Qty", "Total_Plan_Cost"
                ]].sort_values("Total_Plan_Cost", ascending=False).reset_index(drop=True)

                # ── 6. ورقة تحليل الحساسية Sensitivity Analysis ───────────────────────────────
                _grand_total = debug_costing_df["Extended_Cost"].sum()
                sensitivity_df = (
                    debug_costing_df
                    .groupby(["Raw_Material", "RM_Description", "MRP_Controller"], as_index=False)
                    .agg(
                        Products_Used_In = ("Final_Product", "nunique"),
                        Total_Cum_Qty = ("Cum_Qty", "sum"),
                        Unit_Price_EGP = ("Unit_Price_EGP", "first"),
                        Total_Cost = ("Extended_Cost", "sum"),
                    )
                )
                sensitivity_df["Impact_%"] = (sensitivity_df["Total_Cost"] / (_grand_total if _grand_total > 0 else 1) * 100).round(3)
                sensitivity_df["Critical"] = sensitivity_df["Impact_%"].apply(
                    lambda x: "🔥 حرج" if x >= 5 else ("⚠️ متوسط" if x >= 1 else "✅ منخفض")
                )
                sensitivity_df = sensitivity_df.sort_values("Impact_%", ascending=False).reset_index(drop=True)

                # ── 7. ورقة الأثر المالي للمخططين Controllers_Financial_Impact ───────────────────────
        ctrl_financial = (
            debug_costing_df
            .groupby("MRP_Controller", as_index=False)
            .agg(
                Raw_Materials_Count = ("Raw_Material", "nunique"),
                Total_Cost_EGP = ("Extended_Cost", "sum"),
            )
        )
        ctrl_financial["Cost_%"] = (ctrl_financial["Total_Cost_EGP"] / (_grand_total if _grand_total > 0 else 1) * 100).round(2)
        ctrl_financial = ctrl_financial.sort_values("Total_Cost_EGP", ascending=False).reset_index(drop=True)

        st.toast("✅ التحليل المالي (Bottom-Up) اكتمل بنجاح دون أي فقد في البيانات", icon="💰")


        # ==============================================================================
        # 🌟 بناء الـ Pivot Table الأفقي مباشرة وبدقة قبل تغيير أسماء الأعمدة
        # ==============================================================================

        if not debug_costing_df.empty:
            try:
                # 1. نسخ جدول التكلفة التفصيلي الخام لتأمين البيانات
                df_src = debug_costing_df.copy()
                
                # 2. تأمين وتوحيد مسمى عمود المنتج النهائي من الجدول الخام
                if "Final_Product" not in df_src.columns and "Level_1" in df_src.columns:
                    df_src.rename(columns={"Level_1": "Final_Product"}, inplace=True)
                elif "Final_Product" not in df_src.columns:
                    prod_col = [c for c in df_src.columns if "product" in c.lower() or "منتج" in c]
                    if prod_col:
                        df_src.rename(columns={prod_col[0]: "Final_Product"}, inplace=True)
                    else:
                        df_src["Final_Product"] = "غير محدد"

                # تنظيف أكواد المنتجات لضمان عدم وجود مسافات تعيق الـ Pivot
                df_src["Final_Product"] = df_src["Final_Product"].astype(str).str.strip()

                # 3. حساب إجمالي تكلفة كل منتج تام من البيانات التفصيلية
                with st.spinner("جاري توليد المسارات العكسية..."):

                    product_total_cost = df_src.groupby("Final_Product")["Extended_Cost"].sum().to_dict()
                
                # 4. بناء الـ Pivot Table (تحويل مسؤول الـ MRP إلى أعمدة أفقية)
                df_pivot = df_src.pivot_table(
                    index=["Final_Product"], 
                    columns="MRP_Controller", 
                    values="Extended_Cost", 
                    aggfunc="sum"
                ).fillna(0).reset_index()
                
                # 5. ربط أسماء وأوصاف المنتجات بدقة لمنع ظهور NaN
                name_map = {}
                if "Final_Product_Name" in df_src.columns:
                    name_map = dict(zip(df_src["Final_Product"], df_src["Final_Product_Name"]))
                elif "Name_1" in df_src.columns:
                    name_map = dict(zip(df_src["Final_Product"], df_src["Name_1"]))
                elif '_mat_desc_map' in locals() or '_mat_desc_map' in globals():
                    name_map = _mat_desc_map
                else:
                    name_map = {}
                
                # إدراج عمود الوصف وإجمالي التكلفة في بداية الجدول الأفقي الجديد
                if name_map:
                    df_pivot.insert(1, "Final_Product_Name", df_pivot["Final_Product"].map(name_map).fillna("وصف غير متوفر"))
                else:
                    df_pivot.insert(1, "Final_Product_Name", df_pivot["Final_Product"])
                    
                df_pivot.insert(2, "إجمالي_تكلفة_المنتج", df_pivot["Final_Product"].map(product_total_cost).fillna(0))
                
                # 6. حساب نسبة تأثير كل كنترول وترتيب الأعمدة بالتوالي (القيمة ثم النسبة)
    #            exclude_cols = ["Final_Product", "Final_Product_Name", "إجمالي_تكلفة_المنتج"]
    #            control_columns = [str(c) for c in df_pivot.columns if c not in exclude_cols]
     #           final_ordered_cols = ["Final_Product", "Final_Product_Name", "إجمالي_تكلفة_المنتج"]
      #          for ctrl in control_columns:
       #             df_pivot[f"Impact_%_{ctrl}"] = df_pivot.apply(
        #                lambda row: round((row[ctrl] / row["إجمالي_تكلفة_المنتج"] * 100), 2) if row["إجمالي_تكلفة_المنتج"] > 0 else 0.0, 
         #               axis=1
          #          )
           #         final_ordered_cols.extend([ctrl, f"Impact_%_{ctrl}"])
       #         Controllers_Financial_Impact_Pivot = df_pivot[final_ordered_cols]

                # 6. حساب نسبة تأثير كل كنترول وترتيب الأعمدة بالتوالي (القيمة ثم النسبة)
                exclude_cols = ["Final_Product", "Final_Product_Name", "إجمالي_تكلفة_المنتج"]
                control_columns = [str(c) for c in df_pivot.columns if c not in exclude_cols]
                
                # إنشاء قائمتين منفصلتين لتنظيم المخرجات
                value_cols = []
                percentage_cols = []
                
                for ctrl in control_columns:
                    # حساب النسبة المئوية لكل كنترولر بدقة
                    df_pivot[f"Impact_%_{ctrl}"] = df_pivot.apply(
                        lambda row: round((row[ctrl] / row["إجمالي_تكلفة_المنتج"] * 100), 2) if row["إجمالي_تكلفة_المنتج"] > 0 else 0.0, 
                        axis=1
                    )
                    # فصل الأعمدة: القيم المالية تذهب لقائمة، والنسب تذهب لقائمة أخرى
                    value_cols.append(ctrl)
                    percentage_cols.append(f"Impact_%_{ctrl}")
                
                # تجميع الترتيب النهائي: الأعمدة الثابتة -> ثم كل الكنترولرات مالياً -> ثم كل النسب المئوية في النهاية
                final_ordered_cols = ["Final_Product", "Final_Product_Name", "إجمالي_تكلفة_المنتج"] + value_cols + percentage_cols
                
                # حفظ الجدول النهائي في المتغير المطلق بالهيكل الجديد
                Controllers_Financial_Impact_Pivot = df_pivot[final_ordered_cols]

            except Exception as pivot_err:
                st.error(f"⚠️ حدث خطأ أثناء إعادة هيكلة شيت الكنترول المالي: {pivot_err}")
                Controllers_Financial_Impact_Pivot = pd.DataFrame()
        else:
            Controllers_Financial_Impact_Pivot = pd.DataFrame()

        # الآن نُعيد تسمية Raw_Material → RM_Code بأمان دون أن يؤثر على الـ Pivot
        # ملاحظة: Parent_Material يبقى بنفس اسمه ولا يتأثر بهذا الـ rename
        debug_costing_df.rename(columns={
            "Raw_Material": "RM_Code"
        }, inplace=True)

    except Exception as _fe:
        st.sidebar.error(f"❌ خطأ في المحرك المالي: {_fe}")
        financial_active = False

# ***************************************************************************
# ── بداية جزء معالجة خطة الاحتياجات وعرض النتائج على الواجهة ───────────────────
# ***************************************************************************

    if result_df.empty:
        st.warning("⚠️ لم يتم العثور على مكونات مطابقة بين الخطة والـ BOM.")
    else:
        # تجميع إجمالي لكل مكون × تاريخ × نوع الطلب
        # ✅ نعتمد على BOM Level (المحسوب تعاودياً) وليس hierarchy_level من ورقة Component
        merged_df = (
            result_df
            .groupby([
                col("component"),
                col("component_desc"),
                col("component_uom"),
                col("mrp_controller"),
                col("current_stock"),
                col("component_order_type"),
                "Order Type",
                "Date",
                "BOM Level",
            ], as_index=False)
            ["Required Component Quantity"]
            .sum()
        )

        actual_levels = sorted(merged_df["BOM Level"].unique())
#        st.success(
 #           f"✅ إجمالي صفوف الاحتياج: {len(merged_df):,} | "
  #          f"مكونات فريدة: {merged_df[col('component')].nunique():,} | "
   #         f"المستويات المحسوبة: {actual_levels}"
    #    )

        # 🔍 DEBUG: مساعدة في التشخيص — يمكن إخفاؤه بعد التحقق
#        with st.expander("🔍 تشخيص: عيّنة من نتائج result_df الخام (قبل التجميع)"):
  #          debug_sample = result_df[["Parent", col("component"), "Order Type", "Date",
  #                                    col("component_qty"), "Required Component Quantity", "BOM Level"]].copy()
   #         debug_sample["Date"] = debug_sample["Date"].astype(str)
    #        st.dataframe(debug_sample.sort_values(["BOM Level", "Parent", col("component")]).head(100),
     #                    width="stretch")
      #      st.caption(f"إجمالي الصفوف الخام: {len(result_df):,}")

        # عرض مبسط بالمستوى
#        display_cols = [
 #           col("component"), col("component_desc"),
  #          col("mrp_controller"), col("component_order_type"),
   #         "Order Type", "Date", "Required Component Quantity", "BOM Level"
    #    ]
    #    display_cols = [c for c in display_cols if c in merged_df.columns]
#        st.dataframe(merged_df[display_cols].sort_values(
 #           ["BOM Level", col("component"), "Date"]
  #      ), width="stretch")

    # ==============================================================================
    # C–J. حساب البيانات + الواجهة الجديدة (Tabs + Expanders)
    # ==============================================================================

    # ── حساب كل البيانات أولاً (بدون عرض) ──────────────────────────────────
    total_models     = plan_df[col("material")].nunique()
    total_components = component_df[col("component")].nunique()
    total_boms       = len(component_df)

    empty_mrp_count = (component_df.loc[component_df[col("mrp_controller")].replace(r'^\s*$', pd.NA, regex=True).fillna("غير محدد").eq("غير محدد"),col("component")].nunique())

  
    diff_uom = component_df.groupby(col("component"))[col("component_uom")].nunique()
    diff_uom = diff_uom[diff_uom > 1]
    total_diff_uom = len(diff_uom)

    # اضافة المسمى جانب الكود لاكثر من وحدة
#    diff_uom_str   = ", ".join(map(str, diff_uom.index)) if total_diff_uom > 0 else "لا يوجد"
    diff_uom_str = ", ".join(
        f"{comp_code} ({component_df.loc[component_df[col('component')] == comp_code, 'Component Description'].iloc[0]})"
        for comp_code in diff_uom.index) if total_diff_uom > 0 else "لا يوجد"

    diff_uom_color = "red" if total_diff_uom > 0 else "green"

    missing_boms      = set(plan_df[col("material")]) - set(component_df[col("material")])
    total_missing_boms = len(missing_boms)
    missing_boms_html  = (
        f"<span style='color:red;'>{', '.join(map(str, missing_boms))}</span>"
        if missing_boms else "<span style='color:green;'>لا يوجد</span>"
    )

    # إحصائيات نوع الطلب
    order_type_map = {"F": "شراء", "E": "تصنيع"}
    component_df["Order_Type_Label"] = component_df[col("component_order_type")].map(order_type_map).fillna("غير محدد")
    purchase_count      = component_df.loc[component_df["Order_Type_Label"] == "شراء",    col("component")].nunique()
    manufacturing_count = component_df.loc[component_df["Order_Type_Label"] == "تصنيع",   col("component")].nunique()
    undefined_count     = component_df.loc[component_df["Order_Type_Label"] == "غير محدد", col("component")].nunique()

    # المستويات الهرمية الموجودة
    levels_summary = (
        component_df.groupby(col("hierarchy_level"))[col("component")]
        .nunique()
        .reset_index()
        .rename(columns={col("component"): "عدد المكونات", col("hierarchy_level"): "المستوى"})
    )

  #  st.markdown(f"placeholder_summary") if False else None

    # ── D. Need_By_Date ──────────────────────────────────────────────────────

    if not result_df.empty:
        # تجميع كل المستويات: لكل مكون × تاريخ → جمع الاحتياجات
        result_date = (
            merged_df
            .groupby([
                col("component"), col("component_desc"), col("component_uom"),
                col("mrp_controller"), col("current_stock"), col("component_order_type"),
                "Date"
            ], as_index=False)
            ["Required Component Quantity"]
            .sum()
        )

        pivot_by_date = result_date.pivot_table(
            index=[
                col("component"), col("component_desc"), col("component_uom"),
                col("mrp_controller"), col("current_stock"), col("component_order_type"),
            ],
            columns="Date",
            values="Required Component Quantity",
            aggfunc="sum",
            fill_value=0
        ).reset_index()

        # تنسيق أسماء أعمدة التواريخ
        pivot_by_date.columns = [
            c.strftime("%d %b") if isinstance(c, (pd.Timestamp, datetime.datetime)) else c
            for c in pivot_by_date.columns
        ]

        pass  # البيانات جاهزة — العرض في التبويبات

    # ── E. Need_By_Order_Type ────────────────────────────────────────────────

    if not result_df.empty:
        result_order = (
            merged_df
            .groupby([
                col("component"), col("component_desc"), col("component_uom"),
                col("mrp_controller"), col("current_stock"), col("component_order_type"),
                "Order Type", "Date"
            ], as_index=False)
            ["Required Component Quantity"]
            .sum()
        )

        pivot_by_order = result_order.pivot_table(
            index=[
                col("component"), col("component_desc"), col("component_uom"),
                col("mrp_controller"), col("current_stock"), col("component_order_type"),
            ],
            columns=["Date", "Order Type"],
            values="Required Component Quantity",
            aggfunc="sum",
            fill_value=0
        ).reset_index()

        # تسطيح أسماء الأعمدة المركبة
        flat_cols = []
        for c in pivot_by_order.columns:
            if isinstance(c, tuple):
                date_part, ot_part = c
                if isinstance(date_part, (pd.Timestamp, datetime.datetime)):
                    flat_cols.append(f"{ot_part} - {date_part.strftime('%d %b')}")
                else:
                    flat_cols.append(str(date_part) if date_part else str(ot_part))
            else:
                flat_cols.append(c)
        pivot_by_order.columns = flat_cols

        pass  # البيانات جاهزة

    # ── F. تحليل الرصيد والتغطية ────────────────────────────────────────────

    if not result_df.empty:
        component_analysis = (
            merged_df
            .groupby([
                col("component"), col("component_desc"), col("component_uom"),
                col("current_stock"), col("component_order_type"),
                "BOM Level", col("mrp_controller"),
            ], as_index=False)
            .agg(
                Required_Qty=("Required Component Quantity", "sum"),
                Order_Types=("Order Type", lambda x: ", ".join(sorted(set(str(v) for v in x if pd.notna(v)))))
            )
            .rename(columns={
                "Required_Qty": "Required Component Quantity",
                "Order_Types": "Order Type",
            })
        )


        # 🔹 تنظيف وتحويل الأعمدة الرقمية
        numeric_cols = [col("current_stock"), "Required Component Quantity"]

        for c in numeric_cols:
                component_analysis[c] = component_analysis[c].astype(str).str.strip()
                component_analysis[c] = component_analysis[c].str.replace(r'[^\d\.]', '', regex=True)
                component_analysis[c] = pd.to_numeric(component_analysis[c], errors='coerce')

        # 🔹 حساب نسبة التغطية + تحويل الناتج + التقريب
        component_analysis["Coverage Percentage"] = pd.to_numeric(
                component_analysis[col("current_stock")] /
                component_analysis["Required Component Quantity"].replace(0, pd.NA) * 100,
                errors='coerce'
        ).round(1).fillna(0)



        component_analysis["Coverage Status"] = component_analysis["Coverage Percentage"].apply(
            lambda x: "🟢 كافية" if x >= 100 else ("🟡 جزئية" if x >= 50 else "🔴 غير كافية")
        )
        component_analysis["Priority"] = component_analysis.apply(
            lambda row: "🔥 عاجل" if row["Coverage Percentage"] < 30 and row["Required Component Quantity"] > 1000
            else ("⚠️ متوسط" if row["Coverage Percentage"] < 50 else "✅ منخفض"),
            axis=1
        )

        pass  # البيانات جاهزة
    else:
        component_analysis = pd.DataFrame()

    # ── G. Component in BOMs ─────────────────────────────────────────────────
 
    if not result_df.empty:
        # نُنشئ plan_unit: صف واحد لكل (Material, material_desc, Order Type) بكمية = 1
        unit_plan = (
            plan_melted[[col("material"), col("material_desc"), col("order_type")]]
            .drop_duplicates()
            .copy()
        )
        unit_plan["Planned Quantity"] = 1
        unit_plan["Date"] = pd.Timestamp("2000-01-01")   # تاريخ وهمي ثابت
 
        # نُشغّل explosion بكمية = 1 → يعطي النمطي التراكمي لكل منتج
        unit_result = bom_explosion(unit_plan, component_df)
 
        if not unit_result.empty:
            # 🔹 المفتاح: Material + Order Type فقط (بدون material_desc)
            # السبب: material_desc في unit_result يأتي من bom_explosion وقد يكون فارغاً
            # مما يُفشل الدمج ويُعيد plan_qty = NaN → 0
            plan_qty_map = (
                plan_melted.groupby([col("material"), col("order_type")])["Planned Quantity"]
                .sum()
                .reset_index()
                .rename(columns={"Planned Quantity": "plan_qty"})
            )
            # نُحضّر material_desc الصحيح من plan_melted بشكل منفصل
            mat_desc_map = (
                plan_melted[[col("material"), col("material_desc")]]
                .drop_duplicates(subset=[col("material")])
                .copy()
            )

            # 🔹 توحيد الأنواع
            unit_result[col("material")] = unit_result[col("material")].astype(str)
            unit_result["Order Type"]    = unit_result["Order Type"].astype(str)
            plan_qty_map[col("material")]   = plan_qty_map[col("material")].astype(str)
            plan_qty_map[col("order_type")] = plan_qty_map[col("order_type")].astype(str)
            mat_desc_map[col("material")]   = mat_desc_map[col("material")].astype(str)

            # 🔹 دمج plan_qty بـ Material + Order Type ← يضمن إيجاد الكمية دائماً
            unit_result = unit_result.merge(
                plan_qty_map,
                left_on=[col("material"), "Order Type"],
                right_on=[col("material"), col("order_type")],
                how="left",
                suffixes=("", "_plan")
            )
            # 🔹 دمج material_desc الصحيح (يُستخدم في رأس العمود فقط)
            if col("material_desc") in unit_result.columns:
                unit_result.drop(columns=[col("material_desc")], inplace=True)
            unit_result = unit_result.merge(
                mat_desc_map,
                on=col("material"),
                how="left"
            )
 
            # رأس العمود: كود المنتج , الكمية الفعلية , وصفه (نوع الطلب)
            unit_result["model_info"] = (
                unit_result[col("material")].astype(str) + " ; " +
                unit_result["plan_qty"].fillna(0).round(0).astype(int).astype(str) + " ; " +
                unit_result[col("material_desc")].astype(str).fillna("") + " (" +
                unit_result["Order Type"].astype(str).fillna("") + ")"
            )


            # ✅ نُلغي BOM Level من الـ pivot — نجمع كل المستويات في صف واحد
            pivot_index = [col("component"), col("component_desc"),
                           col("mrp_controller"), col("component_uom")]
            pivot_index = [c for c in pivot_index if c in unit_result.columns]
 

            component_bom_pivot = unit_result.pivot_table(
                index=pivot_index,
                columns="model_info",
                values="Required Component Quantity",
                aggfunc="sum"
            ).reset_index()

            component_bom_pivot.columns.name = None
        else:
            component_bom_pivot = pd.DataFrame()
    else:
        component_bom_pivot = pd.DataFrame()

    # ── H. الكميات الشهرية ───────────────────────────────────────────────────
    if date_cols:
        orders_summary = plan_df_orig.melt(
            id_vars=[col("material"), col("material_desc"), col("order_type")],
            value_vars=date_cols,
            var_name="Month",
            value_name="Quantity"
        )
        orders_summary["Quantity"] = pd.to_numeric(orders_summary["Quantity"], errors="coerce").fillna(0)
        try:
            orders_summary["Month"] = pd.to_datetime(orders_summary["Month"]).dt.month_name()
        except Exception:
            pass

        orders_grouped = (
            orders_summary
            .groupby(["Month", col("order_type")])
            .agg({"Quantity": "sum"})
            .reset_index()
        )
        pivot_monthly = orders_grouped.pivot_table(
            index="Month", columns=col("order_type"),
            values="Quantity", aggfunc="sum", fill_value=0
        ).reset_index()

        if "E" not in pivot_monthly.columns: pivot_monthly["E"] = 0
        if "L" not in pivot_monthly.columns: pivot_monthly["L"] = 0
        pivot_monthly["الإجمالي"] = pivot_monthly["E"] + pivot_monthly["L"]
        total_sum = pivot_monthly["الإجمالي"].sum()
        if total_sum > 0:
            pivot_monthly["E%"] = (pivot_monthly["E"] / pivot_monthly["الإجمالي"] * 100).round(1).astype(str) + "%"
            pivot_monthly["L%"] = (pivot_monthly["L"] / pivot_monthly["الإجمالي"] * 100).round(1).astype(str) + "%"
        else:
            pivot_monthly["E%"] = pivot_monthly["L%"] = "0.0%"

        month_order = {m: i for i, m in enumerate(calendar.month_name) if m}
        pivot_monthly = pivot_monthly.sort_values(
            by="Month", key=lambda x: x.map(lambda v: month_order.get(v, 99))
        )

        pass  # البيانات جاهزة للعرض في التبويبات

    # ── I. إعداد Summary للتصدير ─────────────────────────────────────────────
    coverage_stats_export = []
    if not result_df.empty:
        tc2 = max(len(component_analysis), 1)
        sc2  = len(component_analysis[component_analysis["Coverage Percentage"] >= 100])
        pc2  = len(component_analysis[(component_analysis["Coverage Percentage"] >= 50) & (component_analysis["Coverage Percentage"] < 100)])
        ic2  = len(component_analysis[component_analysis["Coverage Percentage"] < 50])
        crt2 = len(component_analysis[component_analysis["Priority"] == "🔥 عاجل"])
        coverage_stats_export = [
            ["🟢 مكونات تغطية كافية", sc2, f"{sc2/tc2*100:.1f}%"],
            ["🟡 مكونات تغطية جزئية", pc2, f"{pc2/tc2*100:.1f}%"],
            ["🔴 مكونات تغطية غير كافية", ic2, f"{ic2/tc2*100:.1f}%"],
            ["🔥 مكونات حرجة", crt2, ""],
        ]

    # ── بيانات الكميات الشهرية للـ Summary (تعديل ذكي لرصد النوع L المعوض) ────
    monthly_summary_rows = []
    if date_cols:
        monthly_summary_rows = [["", "", ""], ["📅 الكميات الشهرية", "", ""]]
        for _, mrow in pivot_monthly.iterrows():
            # التأكد من جلب قيم E و L و F حتى لو لم تكن متواجدة ببعض الشهور بدون حدوث KeyError
            val_e = int(mrow.get('E', 0)) if 'E' in mrow else 0
            val_l = int(mrow.get('L', 0)) if 'L' in mrow else 0
            val_f = int(mrow.get('F', 0)) if 'F' in mrow else 0 # تأمين نوع الشراء أيضاً
            
            monthly_summary_rows.append([
                mrow["Month"],
                int(mrow.get("الإجمالي", 0)),
                f"E: {val_e:,}  |  L: {val_l:,}  |  F: {val_f:,}"
            ])
            
        # حساب الإجماليات النهائية لكافة أنواع الأوامر المتاحة بالجدول المفصلي
        e_total = int(pivot_monthly.get("E", pd.Series([0])).sum())
        l_total = int(pivot_monthly.get("L", pd.Series([0])).sum())
        f_total = int(pivot_monthly.get("F", pd.Series([0])).sum())
        grand   = e_total + l_total + f_total
        
        monthly_summary_rows.append(["الإجمالي الكلي", grand, f"E: {e_total:,}  |  L: {l_total:,}  |  F: {f_total:,}"])

    # ==============================================================================
    # ⏱️ نهاية الحسابات وزمن معالجة العمليات (حماية كاملة من الـ NameError)
    # ==============================================================================
    end_time = time.time()
    
    # فحص ذكي للوقت المستغرق
    if 'start_time' in locals() or 'start_time' in globals():
        processing_time = end_time - start_time
    else:
        processing_time = 0.0

    proc_min = int(processing_time) // 60
    proc_sec = int(processing_time) % 60

    # تصفيف وعرض زمن المعالجة في الـ Sidebar بشكل منسق بدون تكرار
    if processing_time > 0:
        st.sidebar.caption(f"⏱️ زمن المعالجة والتفجير الهرمي: {proc_min} دقيقة و {proc_sec} ثانية")
    else:
        st.sidebar.caption("⏱️ زمن المعالجة: لم يتم رصد نقطة البداية start_time")

    # 🛡️ الحساب الفعلي الصحيح والديناميكي لـ missing_ot لمنع الـ NameError 
    # قمنا بربطه مباشرة بـ missing_ot المحسوب عند فحص الـ plan_df في البداية
    if 'missing_ot' not in locals() and 'missing_ot' not in globals():
        missing_ot = 0

    # بناء مصفوفة البيانات للملخص الشامل
    summary_data = [
        ["📌 ملخص نتائج الخطة", "", ""],
        ["موديلات بالخطة", total_models, ""],
        ["مكونات فريدة", total_components, ""],
        ["سطور BOM", total_boms, ""],
        ["مكونات بدون MRP Controller", empty_mrp_count, ""],
        ["مكونات بأكثر من وحدة", total_diff_uom, diff_uom_str],
        ["منتجات بالخطة بدون BOM", total_missing_boms, ""],
        ["", "", ""],
        ["مكونات شراء (F)", purchase_count, ""],
        ["مكونات تصنيع (E)", manufacturing_count, ""],
        ["مكونات غير محددة (معوضة بـ L)", missing_ot, ""], # تم تعديل المسمى هنا ليطابق الواقع الفعلي
        ["", "", ""],
        ["📈 إحصائيات التغطية", "", ""],
        *coverage_stats_export,
        *monthly_summary_rows,
        ["", "", ""],
        ["⏱️ زمن المعالجة", f"{proc_min} دقيقة و {proc_sec} ثانية", f"{processing_time:.2f} ثانية"],
        ["تاريخ الإنشاء", datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), ""],
    ]

    # إضافة صف التحذير بأمان الآن دون الخوف من الـ NameError بالرقم الفعلي للتعويض
    if missing_ot > 0:
        summary_data.append(["", "", ""])
        summary_data.append(["⚠️ Order Type", f"{missing_ot} قيمة معوضة", "تم استبدالها بـ 'L'"])

        st.session_state["missing_ot"] = missing_ot #تخزين الرقم داخل string ثم محاولة قراءته لاحقًا 

    summary_df = pd.DataFrame(summary_data, columns=["البند", "القيمة", "ملاحظات"])

    # تنسيق plan_df للتصدير
    plan_df_export = plan_df.copy()
    plan_df_export.columns = [
        c.strftime("%d %b") if isinstance(c, (datetime.datetime, pd.Timestamp)) else c
        for c in plan_df_export.columns
    ]

    # ── BOM Paths: df_bom_paths مُحسَّب مسبقاً (قبل المحرك المالي) — لا إعادة حساب ──
    # ==============================================================================
    # الواجهة الرئيسية — ملخص ثابت + 5 تبويبات
    # ==============================================================================

    def colored_metric(title, value, is_total=False, force_black=False):
        if force_black:
            color = "#000000"  # أسود ثابت
        elif is_total:
            color = "#1976d2"  # أزرق
        else:
            # تأمين الفحص إذا كانت القيمة نصية أو تحتوي على رموز قبل فحص الصفر
            try:
                val_numeric = float(str(value).replace(',', '').strip())
                color = "#2e7d32" if val_numeric == 0 else "#d32f2f"  # أخضر للـ 0 وأحمر لغيره
            except ValueError:
                color = "#2e7d32" if value == 0 else "#d32f2f"

        # إضافة تنسيق محكم (Card Style) يمنع الإزاحة تماماً
        return f"""
        <div style="text-align:center; padding: 5px 0; margin: 0; line-height: 1.2; direction: rtl;">
            <div style="font-size: 14px; color: #555555; margin-bottom: 2px;">
                {title}
            </div>
            <div style="font-size: 26px; font-weight: bold; color: {color}; margin: 0; padding: 0; letter-spacing: -0.5px;">
                {value}
            </div>
        </div>
        """
    # ==============================================================================
    # ⏱️ عرض زمن المعالجة الاحترافي وتطبيق تنسيقات الواجهة (CSS)
    # ==============================================================================
   # st.markdown(
  #      f"""
       # <div style="direction:rtl; text-align:right; margin:15px 0; font-size:14px;
      #              background-color:#e3f2fd; padding:10px; border-radius:6px; 
     #               border-right: 5px solid #1565c0; color: #0d47a1;">
    #        ⏱️ <b>زمن المعالجة والتفجير الهرمي:</b> {proc_min} دقيقة و {proc_sec} ثانية ({processing_time:.2f} ثانية)
    #    </div>
 #       """, unsafe_allow_html=True)

    # تطبيق ستايل التبويبات الملونة
    st.markdown("""
    <style>
    /* ====== شكل عام للتابات ====== */
    div[data-testid="stTabs"] button {
        font-size: 18px !important;
        font-weight: bold !important;
        padding: 10px 18px !important;
        border-radius: 8px !important;
        margin-right: 5px;
        transition: 0.2s;
    }

    /* ====== Tab 1: الخطة ====== */
    div[data-testid="stTabs"] button:nth-child(1) {
        background-color: #e3f2fd !important;
        color: #1565c0 !important;
    }

    /* ====== Tab 2: الاحتياج ====== */
    div[data-testid="stTabs"] button:nth-child(2) {
        background-color: #e8f5e9 !important;
        color: #2e7d32 !important;
    }

    /* ====== Tab 3: التغطية ====== */
    div[data-testid="stTabs"] button:nth-child(3) {
        background-color: #fff3e0 !important;
        color: #ef6c00 !important;
    }

    /* ====== Tab 4: BOM ====== */
    div[data-testid="stTabs"] button:nth-child(4) {
        background-color: #f3e5f5 !important;
        color: #6a1b9a !important;
    }

    /* ====== Tab 6: Excel تصدير النتائج ====== */
    div[data-testid="stTabs"] button:nth-child(5) {
        background-color: #e8f6e9 !important;
        color: #2e7d32 !important;
    }

    /* ====== Tab 5:  التقرير المالى ====== */
    div[data-testid="stTabs"] button:nth-child(6) {
        background-color: #fce4ec !important;
        color: #c62828 !important;
    }

    /* ====== التاب النشط المختار ====== */
    div[data-testid="stTabs"] button[aria-selected="true"] {
        border-bottom: 3px solid #1976d2 !important;
        transform: scale(1.05);
        box-shadow: 0px 4px 10px rgba(0, 0, 0, 0.05);
    }

    /* ✨ لون الخط في التاب النشط */
    div[data-testid="stTabs"] button[aria-selected="true"] p {
        color: inherit !important;}
        color: #ffeb3b !important;   /* أصفر مميز */
        font-weight: 800 !important;
    }

    /* ✨ لون  button */
    div[data-testid="stTabs"] button[aria-selected="true"] {
        background: linear-gradient(135deg, #1d4ed8, #2563eb) !important;
        color: white !important;
    }

    /* ====== النص داخل التاب ====== */
    div[data-testid="stTabs"] button p {
        font-size: 20px !important;
        font-weight: bold !important;
    }


        
    </style>

    """, unsafe_allow_html=True)


    # ==============================================================================
    # ── بداية بناء التبويبات الفعلي وعرض الجداول المخصصة لكل تبويب ──────────────────
    # ==============================================================================
    st.markdown("---")
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "🗓️ الخطة",
        "📦 الاحتياج",
        "📊 التغطية",
        "🌿 BOM",
        "💰 التحليل المالي",
        "📤 التصدير",
    ])

    # ══ TAB 1: الخطة (تم نقل الملخص الثابت والتحذيرات إلى هنا) ═══════════════════════════════
    with tab1:
        st.markdown(
            '<p style="font-size:22px;font-weight:bold;color:#1976d2;margin-top:10px;margin-bottom:5px;">📌 ملخص نتائج الخطة</p>',
            unsafe_allow_html=True
        )
        # ─── 👈 بطاقات المؤشرات (6 أعمدة ومحاذاة جهة اليسار) ────────────────
     #   st.markdown('<div style="direction: ltr !important; text-align: left !important;">', unsafe_allow_html=True)



        _k1, _k2, _k3, _k4, _k5, _k6, _k7 = st.columns(7)


        _k1.metric("🏭 موديلات", fmt_n(total_models))
        _k2.metric("🔩 مكونات فريدة", fmt_n(total_components))
        _k3.metric("📋 سطور BOM", fmt_n(total_boms))

        _k4.metric("❗ بدون MRP", fmt_n(empty_mrp_count))
        _k5.metric("🚫 بدون BOM", fmt_n(total_missing_boms))
        _k6.metric("⚠️ اختلاف وحدة القياس", fmt_n(total_diff_uom))
        _k7.metric("📉 موديلات بدون Order Type", fmt_n(st.session_state.get("missing_ot", 0)))

        # عرض التحذيرات التابعة للملخص داخل التبويب
        if total_missing_boms > 0:
            st.markdown(f"⚠️ منتجات بالخطة بدون BOM: {missing_boms_html}", unsafe_allow_html=True)
        if total_diff_uom > 0:
            st.markdown(
                f'<span style="color:red;">⚠️ مكونات بأكثر من وحدة: {diff_uom_str}</span>',
                unsafe_allow_html=True
            )
            
        st.markdown(f"""
        <div style="direction:rtl; text-align:right; margin: 10px 0; font-size:20px; line-height:1.8;">
            <div style="padding:5px 0;">🛒 عدد انواع المكونات المطلوبة لشراء (F): <b style="color:#2e7d32;">{purchase_count}</b></div>
            <div style="padding:5px 0;">🏭 عدد انواع المكونات المطلوبة تصنيع (E): <b style="color:#1565c0;">{manufacturing_count}</b></div>
            <div style="padding:5px 0;">❓ عدد انواع المكونات غير محدد شراء او تصنيع : <b style="color:#ef6c00;">{undefined_count}</b></div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("---") # خط فاصل جمالي بين الملخص والجداول

        # ─── جداول عرض الخطة الأصلية والتوزيع الشهري ───────────────────────────
        with st.expander("🗓️ عرض الخطة الأصلية", expanded=True):
            _plan_disp = plan_df_orig.copy()
            _plan_disp.columns = [
                c.strftime("%d %b") if isinstance(c, (datetime.datetime, pd.Timestamp)) else c
                for c in _plan_disp.columns
            ]
            st.dataframe(_plan_disp, width="stretch", hide_index=True)

        if date_cols and "pivot_monthly" in dir():
            with st.expander("📊 توزيع الكميات الشهرية", expanded=True):
                _mc1, _mc2 = st.columns([1, 2])
                with _mc1:
                    _html = (
                        "<table border='1' style='border-collapse:collapse;width:100%;"
                        "text-align:center;font-size:13px;'>"
                        "<tr style='background-color:#1976d2;color:white;'>"
                        "<th>الشهر</th><th>E</th><th>L</th><th>الإجمالي</th><th>E%</th><th>L%</th></tr>"
                    )
                    for _, mrow in pivot_monthly.iterrows():
                        _html += (
                            f"<tr><td style='color:#1976d2;font-weight:bold;'>{mrow['Month']}</td>"
                            f"<td>{int(mrow.get('E',0)):,}</td><td>{int(mrow.get('L',0)):,}</td>"
                            f"<td><b>{int(mrow.get('الإجمالي',0)):,}</b></td>"
                            f"<td>{mrow.get('E%','')}</td><td>{mrow.get('L%','')}</td></tr>"
                        )
                    _html += "</table>"
                    st.markdown(f"<div style='direction:rtl;'>{_html}</div>", unsafe_allow_html=True)
                with _mc2:
                    _fig_bar = px.bar(
                        pivot_monthly, x="Month", y=["E", "L"],
                        barmode="group", text_auto=True,
                        title="توزيع الكميات الشهرية",
                        labels={"value": "الكمية", "variable": "نوع الأمر", "Month": "الشهر"},
                        template="streamlit"
                    )
                    _fig_bar.update_layout(height=320, margin=dict(t=40, b=10))
                    st.plotly_chart(_fig_bar, width="stretch")

# ══ TAB 2: الاحتياج ═══════════════════════════════════════════════════════
with tab2:
    if result_df.empty:
        st.info("لا توجد بيانات احتياج.")

    else:

        # ─── عنوان القسم ─────────────────────────────────────
        st.markdown(
            '<p style="font-size:22px;font-weight:bold;color:#1976d2;margin-top:10px;margin-bottom:5px;">📦 ملخص الاحتياج</p>',
            unsafe_allow_html=True
        )

        # ─── حساب المؤشرات من الأعمدة الفعلية المتاحة في result_df ──────────
        _total_req_items = merged_df[col("component")].nunique() if not merged_df.empty else 0
        _gross_qty       = merged_df["Required Component Quantity"].sum() if not merged_df.empty else 0

        # صافي الاحتياج = الإجمالي − المخزون المتاح (الموجب فقط)
        if not merged_df.empty and col("current_stock") in merged_df.columns:
            _net_qty = max(0.0, float(
                _gross_qty - merged_df.drop_duplicates(subset=[col("component")])[col("current_stock")].clip(lower=0).sum()
            ))
        else:
            _net_qty = _gross_qty

        _net_ratio = (_net_qty / _gross_qty * 100) if _gross_qty > 0 else 0.0

        ctrl_col_name = col("mrp_controller")
        if ctrl_col_name in merged_df.columns:
            _mrp_ctrls_count = merged_df[merged_df[ctrl_col_name] != "غير محدد"][ctrl_col_name].nunique()
        else:
            _mrp_ctrls_count = 0

        # ─── 👈 بطاقات المؤشرات (5 أعمدة ومحاذاة جهة اليسار) ────────────────
        st.markdown('<div style="direction: ltr !important; text-align: left !important;">', unsafe_allow_html=True)

        _k1, _k2, _k3, _k4, _k5 = st.columns(5)

        _k1.metric("👤 مخططين مسؤولين",          fmt_n(_mrp_ctrls_count))
        _k2.metric("📦 خامات مطلوبة",          fmt_n(_total_req_items))
        _k3.metric("📊 الاحتياج الإجمالي",            fmt_n(_gross_qty))
        _k4.metric("📈 نسبة الصافي/الإجمالي",            fmt_pct(_net_ratio))
        _k5.metric("🎯 صافي الاحتياج",            fmt_n(_net_qty))

        st.markdown('</div>', unsafe_allow_html=True) # إغلاق حاوية المحاذاة اليسرى
        st.markdown("---")

        # ─── الجداول ────────────────────────────────────────
        with st.expander(
            "📅 Need by Date — الاحتياج حسب التاريخ",
            expanded=True):
            _date_disp = pivot_by_date.copy()

            _date_disp.columns = [
                c.strftime("%d %b")
                if isinstance(c, (datetime.datetime, pd.Timestamp))
                else c
                for c in _date_disp.columns]

            st.dataframe(
                _date_disp,
                width="stretch",
                hide_index=True)

        with st.expander(
            "📦 Need by Order Type — الاحتياج حسب نوع الطلب (E / L)",
            expanded=True):

            st.dataframe(
                pivot_by_order,
                width="stretch",
                hide_index=True)

    # ══ TAB 3: التغطية ════════════════════════════════════════════════════════
    with tab3:
        if component_analysis.empty:
            st.info("لا توجد بيانات تغطية.")
        else:
            # فلاتر
            _fc1, _fc2, _fc3 = st.columns(3)
            with _fc1:
             #   _mrp_opts = sorted(component_analysis[col("mrp_controller")].dropna().unique())

                _mrp_series = (component_analysis[col("mrp_controller")].replace(r"^\s*$", pd.NA, regex=True).fillna("غير محدد"))
                _mrp_opts = sorted(_mrp_series.unique())

                _sel_mrp = st.multiselect("🔍 MRP Controller", options=_mrp_opts, default=_mrp_opts, key="cov_mrp")
            with _fc2:
                _ot_opts = sorted(component_analysis[col("component_order_type")].dropna().unique())
                _sel_ot = st.multiselect("🔍 نوع طلب المكون", options=_ot_opts, default=_ot_opts, key="cov_ot")
            with _fc3:
                _lv_opts = sorted(component_analysis["BOM Level"].dropna().unique())
                _sel_lv = st.multiselect("🔍 المستوى الهرمي", options=_lv_opts, default=_lv_opts, key="cov_lv")

            filtered_analysis = component_analysis[
                component_analysis[col("mrp_controller")].isin(_sel_mrp) &
                component_analysis[col("component_order_type")].isin(_sel_ot) &
                component_analysis["BOM Level"].isin(_sel_lv)
            ]
            _tc  = max(len(filtered_analysis), 1)
            _sc  = len(filtered_analysis[filtered_analysis["Coverage Percentage"] >= 100])
            _pc  = len(filtered_analysis[(filtered_analysis["Coverage Percentage"] >= 50) & (filtered_analysis["Coverage Percentage"] < 100)])
            _ic  = len(filtered_analysis[filtered_analysis["Coverage Percentage"] < 50])
            _crt = len(filtered_analysis[filtered_analysis["Priority"] == "🔥 عاجل"])

            with st.expander("📋 جدول تحليل التغطية", expanded=True):
                _cc1,_cc2,_cc3,_cc4 = st.columns(4)
                _cc1.metric("🟢 كافية",     f"{fmt_n(_sc)} ({_sc/_tc*100:.0f}%)")
                _cc2.metric("🟡 جزئية",     f"{fmt_n(_pc)} ({_pc/_tc*100:.0f}%)")
                _cc3.metric("🔴 غير كافية", f"{fmt_n(_ic)} ({_ic/_tc*100:.0f}%)")
                _cc4.metric("🔥 عاجل",      fmt_n(_crt))
                st.dataframe(
                    filtered_analysis.sort_values("Coverage Percentage"),
                    width="stretch", hide_index=True)

            with st.expander("📊 الرسوم البيانية"):
                _rc1, _rc2 = st.columns(2)
                with _rc1:
                    _fp = px.pie(filtered_analysis, names="Coverage Status",
                                 title="توزيع حالة التغطية", color="Coverage Status",
                                 color_discrete_map={"🟢 كافية":"green","🟡 جزئية":"orange","🔴 غير كافية":"red"})
                    st.plotly_chart(_fp, width="stretch")
                with _rc2:
                    _fo = px.pie(filtered_analysis, names=col("component_order_type"),
                                 title="توزيع حسب نوع الطلب")
                    st.plotly_chart(_fo, width="stretch")

                _top10 = filtered_analysis.nsmallest(10, "Coverage Percentage").copy()
                if not _top10.empty:
                    _top10["Short_Label"] = (
                        _top10[col("component")].astype(str) + " - " +
                        _top10[col("component_desc")].astype(str).str[:25]
                    )
                    _fb = px.bar(
                        _top10.sort_values("Required Component Quantity", ascending=True),
                        y="Short_Label", x="Required Component Quantity",
                        color="Coverage Percentage", orientation='h',
                        title="أقل 10 مكونات في نسبة التغطية",
                        labels={"Required Component Quantity":"كمية الطلب","Short_Label":"المكون","Coverage Percentage":"نسبة التغطية %"},
                        color_continuous_scale="RdYlGn_r"
                    )
                    _fb.update_layout(height=400)
                    st.plotly_chart(_fb, width="stretch")

                if len(_sel_mrp) > 1:
                    _fsun = px.sunburst(
                        filtered_analysis,
                        path=[col("mrp_controller"), "BOM Level", "Coverage Status"],
                        values="Required Component Quantity",
                        title="توزيع الاحتياج حسب MRP Controller والمستوى"
                    )
                    st.plotly_chart(_fsun, width="stretch")

            with st.expander("📈 إحصائيات الأولويات"):
                st.markdown(f"""
                <div style="direction:rtl;font-size:16px;">
                <ul style="list-style-type:none;padding-right:0;">
                    <li>🟢 <b>{_sc}</b> مكونات تغطية كافية ({_sc/_tc*100:.1f}%)</li>
                    <li>🟡 <b>{_pc}</b> مكونات تغطية جزئية ({_pc/_tc*100:.1f}%)</li>
                    <li>🔴 <b>{_ic}</b> مكونات تغطية غير كافية ({_ic/_tc*100:.1f}%)</li>
                    <li>🔥 <b style="color:red;">{_crt}</b> مكونات حرجة تحتاج اهتمام عاجل</li>
                </ul></div>
                """, unsafe_allow_html=True)
                st.dataframe(levels_summary, width="stretch", hide_index=True)

    # ══ TAB 4: BOM ════════════════════════════════════════════════════════════
    with tab4:
        with st.expander("🌿 المسارات الكاملة للـ BOM (BOM Paths)", expanded=True):
            if df_bom_paths.empty:
                st.warning("⚠️ لا توجد مسارات — تحقق من نطاق الكودات (40000000–499999999).")
            else:
                _lc = [c for c in df_bom_paths.columns if c.startswith("Level_")]
                _nc = [c for c in df_bom_paths.columns if c.startswith("Name_")]
                st.success(f"✅ **{len(df_bom_paths):,}** مسار | أقصى عمق: **{len(_lc)}** مستويات")
                _sc1, _sc2 = st.columns([3, 1])
                with _sc1:
                    _search = st.text_input("بحث", placeholder="ابحث بكود أو جزء من اسم...", key="bom_path_search", label_visibility="collapsed")

                with _sc2:
                    _stype = st.radio("نوع البحث", ["بالكود", "بالاسم"], horizontal=True, key="bom_search_type")
                if _search.strip():
                    _term = _search.strip()
                    _mask = (
                        df_bom_paths[_lc].apply(lambda c: c.astype(str).str.strip() == _term).any(axis=1)
                        if _stype == "بالكود"
                        else df_bom_paths[_nc].apply(lambda c: c.astype(str).str.contains(_term, case=False, na=False)).any(axis=1)
                    )
                    _dfx = df_bom_paths[_mask].reset_index(drop=True)
                    if _dfx.empty:
                        st.warning(f"⚠️ لا توجد نتائج للبحث عن: **{_term}**")
                    else:
                        _ur = _dfx["Level_1"].nunique() if "Level_1" in _dfx.columns else 0
                        _fl = [lc for lc in _lc if (_dfx[lc].astype(str).str.strip() == _term).any()]
                        st.markdown(
                            f'<div style="direction:rtl;background:#e8f5e9;padding:8px;border-radius:6px;">'
                            f'📊 <b>{_term}</b> — يظهر في <b>{_ur}</b> منتج | '
                            f'<b>{len(_dfx)}</b> مسار | المستويات: <b>{"، ".join(_fl) or "—"}</b></div>',
                            unsafe_allow_html=True
                        )
                        st.dataframe(_dfx, width="stretch", hide_index=True)
                else:
                    st.dataframe(df_bom_paths, width="stretch", hide_index=True)

        with st.expander("📋 النمطي لكل منتج (Component in BOMs)", expanded=True):

            if component_bom_pivot.empty:
                st.info("لا توجد بيانات.")
            else:
                st.dataframe(component_bom_pivot.round(3).fillna(""), width="stretch")

 #       with st.expander("🔍 تشخيص: عيّنة من result_df الخام"):
    #        if not result_df.empty:
    #            _dbg = result_df[["Parent", col("component"), "Order Type", "Date",
     #                             col("component_qty"), "Required Component Quantity", "BOM Level"]].copy()
      #          _dbg["Date"] = _dbg["Date"].astype(str)
       #         st.dataframe(_dbg.sort_values(["BOM Level","Parent",col("component")]).head(100),
        #                     width="stretch")
         #       st.caption(f"إجمالي الصفوف الخام: {len(result_df):,}")

	# ══ TAB 5: التحليل المالي ═════════════════════════════════════════════════
    with tab5:
        if not financial_active:
            st.info("💡 فعّل التحليل المالي من الـ Sidebar باختيار مصدر الأسعار.")

        elif product_costing.empty:
            st.warning("⚠️ لا توجد بيانات مالية — تحقق من تطابق كودات الخامات مع ملف الأسعار.")
            if not debug_costing_df.empty:
                st.dataframe(debug_costing_df.head(20), width="stretch")
        else:
            # ==============================================================================
            # 🔍 1. بناء الفلتر العام للمنتج النهائي في أعلى التبويب
            # ==============================================================================
            # تجميع المنتجات المتاحة من جدول التكلفة التفصيلي لضمان دقة الفلترة
            available_products = sorted(debug_costing_df["Final_Product"].unique().tolist())
            
            # بناء قائمة منسدلة تحتوي على "الكل" بالإضافة إلى أكواد المنتجات
            product_options = ["الكل"] + available_products
            
            # عرض الفلتر بشكل بارز في الأعلى
            selected_product = st.selectbox(
                "🔍 فلتر عام لتبويب التحليل المالي (اختر كود المنتج النهائي):",
                options=product_options,
                key="global_financial_product_filter",
                help="عند اختيار منتج معين، سيتم تحديث المؤشرات والجداول والرسوم البيانية أدناه تلقائياً لعرض بيانات هذا المنتج فقط.")
            
            # تطبيق الفلترة الديناميكية على الجداول الأساسية بناءً على الاختيار
            if selected_product == "الكل":
                filtered_debug_df = debug_costing_df.copy()
                filtered_product_costing = product_costing.copy()
                filtered_ctrl_pivot = Controllers_Financial_Impact_Pivot.copy() if 'Controllers_Financial_Impact_Pivot' in locals() else pd.DataFrame()
            else:
                filtered_debug_df = debug_costing_df[debug_costing_df["Final_Product"] == selected_product].copy()
                filtered_product_costing = product_costing[product_costing["Final_Product"] == selected_product].copy()
                filtered_ctrl_pivot = Controllers_Financial_Impact_Pivot[Controllers_Financial_Impact_Pivot["Final_Product"] == selected_product].copy() if 'Controllers_Financial_Impact_Pivot' in locals() else pd.DataFrame()

# إعادة حساب الإجمالي والمؤشرات بناءً على البيانات المفلترة الحالية
            _grand = filtered_debug_df["Extended_Cost"].sum() if not filtered_debug_df.empty else 0
            
            # 🔍 تحديد عمود كود الخامة بقائمة أولوية صريحة
            # RM_Code هو الاسم بعد إعادة التسمية داخل المحرك المالي
            # نتجنب "material in col.lower()" لمنع التقاط Parent_Material بالخطأ
            _rm_priority = ["RM_Code", "Raw_Material", "Raw_Material_Code", "Material_Code", "Component", "خامة"]
            rm_col = "RM_Code"  # الاسم المعتمد بعد الـ rename
            for _c in _rm_priority:
                if _c in filtered_debug_df.columns:
                    rm_col = _c
                    break
            if rm_col not in filtered_debug_df.columns and not filtered_debug_df.empty:
                rm_col = filtered_debug_df.columns[0]  # بديل آمن عند غياب كل الخيارات

            # 🔍 محاولة تحديد اسم عمود وصف الخامة تلقائياً
            desc_col = "RM_Description"
            for col in filtered_debug_df.columns:
                if col in ["RM_Description", "Material_Description", "Description", "البيان", "اسم_الخامة"]:
                    desc_col = col
                    break
            if desc_col not in filtered_debug_df.columns:
                desc_col = rm_col # كحماية بديلة

            # إعادة حساب أثر الكنترولرات ديناميكياً بناءً على الفلتر الحالي (مؤمن تماماً)
            if not filtered_debug_df.empty and "MRP_Controller" in filtered_debug_df.columns and "Extended_Cost" in filtered_debug_df.columns:
                dynamic_ctrl = (
                    filtered_debug_df.groupby("MRP_Controller", as_index=False)
                    .agg(
                        Raw_Materials_Count=(rm_col, "nunique"),
                        Total_Cost_EGP=("Extended_Cost", "sum")
                    )
                )
                dynamic_ctrl["Cost_%"] = (dynamic_ctrl["Total_Cost_EGP"] / (_grand if _grand > 0 else 1) * 100).round(2)
                dynamic_ctrl = dynamic_ctrl.sort_values("Total_Cost_EGP", ascending=False).reset_index(drop=True)
            else:
                dynamic_ctrl = pd.DataFrame()

            # إعادة حساب الحساسية ديناميكياً للمنتج المحدد (مؤمن تماماً)
            if not filtered_debug_df.empty and "Extended_Cost" in filtered_debug_df.columns:
                # تجميع الأعمدة المتاحة فقط لمنع أي KeyError مستقبلي
                group_cols = [c for c in [rm_col, desc_col, "MRP_Controller"] if c in filtered_debug_df.columns]
                
                dynamic_sensitivity = (
                    filtered_debug_df.groupby(group_cols, as_index=False)
                    .agg(
                        Products_Used_In=("Final_Product", "nunique") if "Final_Product" in filtered_debug_df.columns else ("Extended_Cost", "count"),
                        Total_Cum_Qty=("Cum_Qty", "sum") if "Cum_Qty" in filtered_debug_df.columns else ("Extended_Cost", "sum"),
                        Unit_Price_EGP=("Unit_Price_EGP", "first") if "Unit_Price_EGP" in filtered_debug_df.columns else ("Extended_Cost", "first"),
                        Total_Cost=("Extended_Cost", "sum")
                    )
                )
                # إعادة تسمية الأعمدة لتطابق التنسيق في الجداول السفلية دون أخطاء
                if rm_col != "Raw_Material" and rm_col in dynamic_sensitivity.columns:
                    dynamic_sensitivity = dynamic_sensitivity.rename(columns={rm_col: "Raw_Material"})
                if desc_col != "RM_Description" and desc_col in dynamic_sensitivity.columns:
                    dynamic_sensitivity = dynamic_sensitivity.rename(columns={desc_col: "RM_Description"})
                
                dynamic_sensitivity["Impact_%"] = (dynamic_sensitivity["Total_Cost"] / (_grand if _grand > 0 else 1) * 100).round(3)
                dynamic_sensitivity["Critical"] = dynamic_sensitivity["Impact_%"].apply(
                    lambda x: "🔥 حرج" if x >= 5 else ("⚠️ متوسط" if x >= 1 else "✅ منخفض")
                )
                dynamic_sensitivity = dynamic_sensitivity.sort_values("Impact_%", ascending=False).reset_index(drop=True)
            else:
                dynamic_sensitivity = pd.DataFrame()

            st.markdown("---")

            # ── 2. المؤشرات المالية المفلترة (KPIs) ──────────────────────────
            _fm1, _fm2, _fm3, _fm4 = st.columns(4)
            _fm1.metric("💰 إجمالي التكلفة الحالية (EGP)", fmt_f(_grand))
            _fm2.metric("💱 سعر الصرف",                   fmt_f(exchange_rate, 1))
            _fm3.metric("📈 زيادة الأسعار",               fmt_pct(price_increase_pct, 0))
            
            if selected_product == "الكل":
                _fm4.metric("🏭 إجمالي المنتجات الكلية",     fmt_n(len(filtered_product_costing)))
            else:
                _fm4.metric("📦 حالة المنتج الحالي",         "مُفلتر ")

            # ── 3. تكلفة المنتجات النهائية ──────────────────────────────────────
            with st.expander("🏭 تكلفة المنتجات النهائية — Bottom-Up Costing", expanded=True):
                st.caption("✅ الحساب يعتمد فقط على الخامات الأساسية (Leaf Nodes) × الكمية التراكمية.")
                st.dataframe(
                    filtered_product_costing.style.format({
                        "Total_Cost_EGP":   "{:,.2f}",
                        "Plan_Total_Qty":   "{:,.0f}",
                        "Total_Plan_Cost":  "{:,.2f}",
                    }),
                    width="stretch", hide_index=True
                )

            # ── 4. الهيكل المتقاطع للأثر المالي للمخططين (Pivot Table) ──────
            if not filtered_ctrl_pivot.empty:
                with st.expander("📊 جدول التوزيع الأفقي المتقاطع للأثر المالي والمخططين (Pivot)", expanded=True):
                    st.caption("🔍 يوضح مساهمة كل مخطط (MRP Controller) داخل المنتج، متبوعاً بنسبة تأثيره في النهاية.")
                    
                    _pivot_fmt = {}
                    for _col in filtered_ctrl_pivot.columns:
                        if _col.startswith("نسبة_تأثير_"):
                            _pivot_fmt[_col] = "{:.2f}%"
                        elif _col not in ["Final_Product", "Final_Product_Name"]:
                            _pivot_fmt[_col] = "{:,.2f}"
                            
                    st.dataframe(
                        filtered_ctrl_pivot.style.format(_pivot_fmt),
                        width="stretch", hide_index=True
                    )

            # ── 5. Controllers Financial Impact ─────────────────────────────────
            with st.expander("👤 التأثير المالي المجمع حسب MRP Controller"):
                st.dataframe(
                    dynamic_ctrl.style.format({
                        "Total_Cost_EGP": "{:,.2f}",
                        "Cost_%":         "{:.2f}%",
                    }),
                    width="stretch", hide_index=True
                )

            # ── 6. Debug Sheet ──────────────────────────────────────────────────
            with st.expander("🔍 تفاصيل الخامات الحسابية (Debug Sheet)"):
                st.caption(
                    "كل صف = خامة واحدة داخل منتج واحد | "
                    "Parent_Material = أول مكون وسيط (SF) تحت المنتج النهائي مباشرة، "
                    "أو المنتج النهائي نفسه إذا كانت الخامة مرتبطة به مباشرةً"
                )
                _dbg_fmt = {
                    "Cum_Qty":        "{:,.4f}",
                    "Unit_Price_EGP": "{:,.2f}",
                    "Extended_Cost":  "{:,.2f}",
                }
                st.dataframe(
                    filtered_debug_df.head(100).style.format(_dbg_fmt),
                    width="stretch", hide_index=True
                )

            # ── 7. Sensitivity Analysis ─────────────────────────────────────────
            with st.expander("🎯 تحليل الحساسية — مساهمة كل خامة في التكلفة الحالية"):
                st.dataframe(
                    dynamic_sensitivity.style.format({
                        "Total_Cum_Qty":  "{:,.3f}",
                        "Unit_Price_EGP": "{:,.2f}",
                        "Total_Cost":     "{:,.2f}",
                        "Impact_%":       "{:.3f}%",
                    }),
                    width="stretch", hide_index=True
                )

            # ── 8. الرسوم البيانية الديناميكية ──────────────────────────────────────
            with st.expander("📊 الرسوم البيانية المالية المتغيرة"):
                _gc1, _gc2 = st.columns(2)
                with _gc1:
                    if not dynamic_ctrl.empty:
                        _fig_ctrl = px.pie(
                            dynamic_ctrl, names="MRP_Controller",
                            values="Total_Cost_EGP",
                            title=f"توزيع ميزانية الكنترولر للمنتج: {selected_product}",
                        )
                        st.plotly_chart(_fig_ctrl, width="stretch")
                with _gc2:
                    if not dynamic_sensitivity.empty:
                        _top10s = dynamic_sensitivity.head(10).copy()
                        _top10s["Short_Label"] = (
                            _top10s["Raw_Material"].astype(str) + " - " +
                            _top10s["RM_Description"].astype(str).str[:20]
                        )
                        _fig_top10 = px.bar(
                            _top10s.sort_values("Impact_%", ascending=True),
                            y="Short_Label", x="Impact_%", orientation="h",
                            title="أعلى 10 خامات تأثيراً في التكلفة الحالية",
                            color="Impact_%", color_continuous_scale="Reds",
                        )
                        _fig_top10.update_layout(height=400)
                        st.plotly_chart(_fig_top10, width="stretch")

                if not dynamic_ctrl.empty:
                    _fig_cb = px.bar(
                        dynamic_ctrl.sort_values("Total_Cost_EGP", ascending=True),
                        x="Total_Cost_EGP", y="MRP_Controller", orientation="h",
                        title="المساهمة المالية الإجمالية لكل كنترولر حالياً",
                        text="Cost_%", color="Total_Cost_EGP",
                        color_continuous_scale="Blues",
                    )
                    _fig_cb.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
                    _fig_cb.update_layout(height=max(300, len(dynamic_ctrl) * 45))
                    st.plotly_chart(_fig_cb, width="stretch")

    # ══ TAB 6: التصدير ════════════════════════════════════════════════════════
    with tab6:
        st.subheader("📤 تصدير النتائج إلى Excel")

        mrp_controller_col = "mrp_controller"
        if not mrp_df.empty and mrp_controller_col in mrp_df.columns:
            mrp_options = sorted(mrp_df[mrp_controller_col].dropna().unique().tolist())
        elif not result_df.empty and mrp_controller_col in result_df.columns:
            mrp_options = sorted(result_df[mrp_controller_col].dropna().unique().tolist())
        else:
            mrp_options = []

        if mrp_options:
            st.markdown(
                '<p style="font-size:16px;color:blue;font-weight:bold;">'
                '👤 اختر MRP Controller (يُطبَّق على جميع الأوراق):</p>',
                unsafe_allow_html=True
            )
            selected_mrp = st.multiselect("", options=mrp_options, default=mrp_options, key="export_mrp")
        else:
            selected_mrp = []
        # الحساب الديناميكي للمسارات العكسية ليعمل في المعاينة والتصدير
        df_reverse_paths = generate_reverse_bom_paths(df_bom_paths, debug_costing_df if 'debug_costing_df' in locals() or 'debug_costing_df' in globals() else None)

        available_sheets = {
            "📋 الخطة الأصلية (Original_Plan)":        ("Original_Plan",           True),
            "📌 الملخص (Summary)":                     ("Summary",                 True),
            "📅 الاحتياج بالتاريخ (Need_By_Date)":      ("Need_By_Date",            not result_df.empty),
            "📦 الاحتياج بنوع الأمر (Need_By_Order)":   ("Need_By_Order_Type",      not result_df.empty),
            "🔍 تحليل التغطية (Stock_Coverage)":        ("Stock_Coverage_Analysis", not component_analysis.empty),
     #       "🌳 BOM الكامل (BOM_All_Levels)":           ("BOM_All_Levels",          not result_df.empty),
            "📊 النمطي لكل منتج (Component_in_BOMs)":   ("Component_in_BOMs",       not component_bom_pivot.empty),
            "🌿 المسارات الأفقية (BOM_Paths)":          ("BOM_Paths",               not df_bom_paths.empty),
            "🗂️ المكونات الأصلية (Original_Component)": ("Original_Component",      True),
            "⚠️ تعارض وحدات القياس (Component UOM Issues)": ("Component_UOM_Issues", not component_uom_issues_sheet.empty),

      # ── أوراق مالية (تظهر فقط عند تفعيل الوحدة المالية) ──

            "💰 تكلفة المنتجات (Product_Costing)":      ("Final_Product_Costing",        financial_active and not product_costing.empty),
            "👤 تأثير Controllers المالي":               ("Controllers_Financial_Impact",  financial_active and not ctrl_financial.empty),
            "🎯 تحليل الحساسية (Sensitivity)":          ("Sensitivity_Analysis",          financial_active and not sensitivity_df.empty),
            "🔍 Debug تفصيل الحساب (Costing_Debug)":    ("Costing_Debug",                 financial_active and not debug_costing_df.empty),
            "📈 تحليل نسب تأثير الكنترول (Horizontal)":   ("Controllers_Impact_Analysis",   financial_active and not Controllers_Financial_Impact_Pivot.empty),
            "🔄 المسارات العكسية للخامات (Reverse_Paths)": ("Reverse_BOM_Paths",       not df_reverse_paths.empty),

        }
        default_checked = {"Original_Plan", "Need_By_Date", "Component_in_BOMs"}

        st.markdown(
            '<p style="font-size:17px;color:blue;font-weight:bold;">اختر الأوراق التي تريد تصديرها:</p>',
            unsafe_allow_html=True
        )


# ===== أزرار التحكم السريع (إلغاء الكل في أقصى اليسار) =====
        col_a, col_b, col_c, col_d = st.columns([1, 1, 1, 1], gap="small")

        # ===== 1. [يمين] اختيار الكل =====
        with col_a:
            if st.button("✅ اختيار الكل", width="stretch"):
                for _, (sname, avail) in available_sheets.items():
                    if avail:
                        st.session_state[f"sheet_{sname}"] = True
                st.rerun()

        # ===== 2. أوراق التخطيط =====
        with col_b:
            if st.button("📦 أوراق التخطيط", width="stretch"):
                planning_sheets = [
                    "Original_Plan", "Summary", "Need_By_Date", 
                    "Need_By_Order_Type", "Stock_Coverage_Analysis", 
                    "Component_in_BOMs", "BOM_Paths", "Original_Component", 
                    "MRP_Controller", "Reverse_BOM_Paths"
                ]
                # تصفير تمهيدي
                for _, (sname, _) in available_sheets.items():
                    st.session_state[f"sheet_{sname}"] = False
                # تفعيل التخطيط
                for sname in planning_sheets:
                    if f"sheet_{sname}" in st.session_state:
                        st.session_state[f"sheet_{sname}"] = True
                st.rerun()

        # ===== 3. التحليل المالي =====
        with col_c:
            if st.button("💰 التحليل المالي", width="stretch"):
                financial_sheets = [
                    "Final_Product_Costing", "Controllers_Financial_Impact", 
                    "Sensitivity_Analysis", "Costing_Debug", 
                    "Controllers_Impact_Analysis", "Summary"
                ]
                # تصفير تمهيدي
                for _, (sname, _) in available_sheets.items():
                    st.session_state[f"sheet_{sname}"] = False
                # تفعيل المالي
                for sname in financial_sheets:
                    if f"sheet_{sname}" in st.session_state:
                        st.session_state[f"sheet_{sname}"] = True
                st.rerun()

        # ===== 4. [يسار] إلغاء الكل =====
        with col_d:
            if st.button("❌ إلغاء الكل", width="stretch"):
                for _, (sname, _) in available_sheets.items():
                    st.session_state[f"sheet_{sname}"] = False
                st.rerun()

        # ===== الفورم ===== ===== ===== ===== ===== ===== ===== ===== ===== ===== ===== ===== ===== =====
        with st.form("export_form"):

            _ec1, _ec2 = st.columns(2)
            selected_sheets = {}

            for i, (label, (sheet_name, avail)) in enumerate(available_sheets.items()):
                _tc = _ec1 if i % 2 == 0 else _ec2

                with _tc:
                    if avail:

                        if f"sheet_{sheet_name}" not in st.session_state:
                            st.session_state[f"sheet_{sheet_name}"] = (
                                sheet_name in default_checked
                            )

                        selected_sheets[sheet_name] = st.checkbox(
                            label,
                            key=f"sheet_{sheet_name}"
                        )

                    else:
                        st.checkbox(
                            label + " *(غير متاح)*",
                            value=False,
                            disabled=True
                        )
            submit_export = st.form_submit_button("🗜️Excel-- اضغط هنا لإنشاء ملف ")

        # ===== إنشاء الملف ===== ===== ===== ===== ===== ===== ===== ===== ===== ===== =====
        if submit_export:
            chosen = [k for k, v in selected_sheets.items() if v]
            if not chosen:
                st.warning("⚠️ لم تختر أي ورقة.")
            else:
                with st.spinner("⏳ جاري إنشاء الملف..."):

                    # الحساب الديناميكي للمسارات العكسية
                    df_reverse_paths = generate_reverse_bom_paths(
                        df_bom_paths,
                        debug_costing_df if 'debug_costing_df' in locals() else None)

                    sheet_data_map = {
                        # ── الأوراق الخطط والاحتياحات ──────────────────────────────────

                        "Original_Plan":           plan_df_export,
                        "Summary":                 summary_df,
                        "Need_By_Date":            pivot_by_date        if not result_df.empty           else pd.DataFrame(),
                        "Need_By_Order_Type":      pivot_by_order       if not result_df.empty           else pd.DataFrame(),
                        "Stock_Coverage_Analysis": component_analysis   if not component_analysis.empty  else pd.DataFrame(),
             #           "BOM_All_Levels":          merged_df            if not result_df.empty           else pd.DataFrame(),
                        "Component_in_BOMs":       component_bom_pivot  if not component_bom_pivot.empty else pd.DataFrame(),
                        "BOM_Paths":               df_bom_paths         if not df_bom_paths.empty        else pd.DataFrame(),
                        "Original_Component":      component_df_orig,
             #           "MRP_Controller":          mrp_df               if not mrp_df.empty              else pd.DataFrame(),
                        "Component_UOM_Issues": component_uom_issues_sheet.copy() if not component_uom_issues_sheet.empty else pd.DataFrame(),

                        # الشيت الجديد المطور للمسارات العكسية
                        "Reverse_BOM_Paths": df_reverse_paths,
                        # ── الأوراق المالية ──────────────────────────────────
                        "Final_Product_Costing":        product_costing  if financial_active and not product_costing.empty  else pd.DataFrame(),
                        "Controllers_Financial_Impact": ctrl_financial   if financial_active and not ctrl_financial.empty   else pd.DataFrame(),
                        "Sensitivity_Analysis":         sensitivity_df   if financial_active and not sensitivity_df.empty   else pd.DataFrame(),
                        "Costing_Debug":                debug_costing_df if financial_active and not debug_costing_df.empty else pd.DataFrame(),
                        "Controllers_Impact_Analysis":  Controllers_Financial_Impact_Pivot if financial_active and not Controllers_Financial_Impact_Pivot.empty else pd.DataFrame(),
                    }

                    # فلترة الـ MRP Controller
                    if selected_mrp:

                        for sname, sdf in sheet_data_map.items():

                            if (
                                not sdf.empty
                                and mrp_controller_col in sdf.columns
                            ):

                                sheet_data_map[sname] = sdf[
                                    sdf[mrp_controller_col].isin(selected_mrp)
                                ]

                    # ✅ V3: Progress bar مع تسمية كل خطوة
                    _n_sheets   = len(chosen)
                    _prog_bar   = st.progress(0, text="⏳ جاري تجهيز الأوراق...")
                    _status_txt = st.empty()

                    # إنشاء Excel
                    output = BytesIO()

                    with pd.ExcelWriter(output, engine='openpyxl') as writer:

                        for _idx, sheet_name in enumerate(chosen):

                            # تحديث شريط التقدم + النص
                            _pct = int((_idx / _n_sheets) * 100)
                            _prog_bar.progress(
                                _pct,
                                text=f"📝 جاري كتابة: **{sheet_name}** ({_idx+1}/{_n_sheets})"
                            )
                            _status_txt.caption(
                                f"⚙️ معالجة الورقة {_idx+1} من {_n_sheets}: `{sheet_name}`"
                            )

                            df_export = sheet_data_map.get(
                                sheet_name,
                                pd.DataFrame()
                            )

                            if not df_export.empty:

                                safe_sheet_name = sheet_name[:31]

                                df_export.to_excel(
                                    writer,
                                    sheet_name=safe_sheet_name,
                                    index=False
                                )


                                # ===== Auto Fit Columns =====
 #                               ws = writer.sheets[safe_sheet_name]
  #                              for col in ws.columns:
   #                                 max_length = 0
    #                                column = col[0].column_letter
#                                    for cell in col:
#                                        try:
 #                                           max_length = max(
  #                                              max_length,
   #                                             len(str(cell.value)))
#                                        except:
 #                                           pass
#                                    ws.column_dimensions[column].width = min(
 #                                       max_length + 2,
  #                                      40)

                    # اكتمل التصدير
                    _prog_bar.progress(100, text="✅ اكتمل إنشاء الملف!")
                    _status_txt.empty()

                    output.seek(0)

                    # تخزين الملف داخل Session State
                    st.session_state["export_file"] = output.getvalue()

                    st.success(
                        f"✅ تم إنشاء الملف — {len(chosen)} ورقة: {', '.join(chosen)}"
                    )

                    st.balloons()

        # ===== زر التحميل ===== ===== ===== ===== ===== ===== ===== ===== ===== ===== =====
        if "export_file" in st.session_state:

            st.download_button(
                label="📊 تحميل ملف Excel",
                data=st.session_state["export_file"],
                file_name=f"MRP_Results&Financial_Impact_{datetime.datetime.now().strftime('%d_%b_%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# --- التذييل ---
st.markdown("""
<hr>
<div style="text-align:center; direction:rtl; font-size:14px; color:gray;">
    ✨ تــــم التنفيذ بواسطة <b>م / رضا رشدي</b> — جميع الحقوق محفوظة © 2026 ✨<br>
    <small style="color:#aaa;">V3 — Iterative BOM · Cached Paths · Progress Export · Unified Formatting</small>
</div>
""", unsafe_allow_html=True)
