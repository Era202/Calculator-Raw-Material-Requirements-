import pandas as pd
from pathlib import Path
from collections import defaultdict
from functools import lru_cache
import streamlit as st
import io
import datetime
from io import BytesIO
import calendar
import plotly.express as px

# Configure the page
st.set_page_config(page_title="MRP_Calculator Raw Material Requirements ", page_icon="📊", layout="wide")

# Title in Arabic
st.title("📊MRP_Calculator Raw Material Requirements (MRP)")
st.markdown("---")

class MRPCalculator:
    def __init__(self):
        self.relations = defaultdict(list)
        self.plan_df = None
        self.bom_df = None
        self.mrp_control_df = None
        self.material_descriptions = {}  # تخزين أوصاف المواد
        self.material_uoms = {}  # تخزين وحدات القياس للمواد
        self.standardized_uoms = {}  # تخزين الوحدات الموحدة
        self.mrp_control_values = {}  # تخزين قيم MRP Contor
        self.manufacturing_quantities = {}  # كميات التصنيع للمكونات الوسيطة
        self.bom_hierarchy = {}  # تخزين هيكل الـ BOM
        
    def load_data(self, uploaded_file) -> bool:
        """Load Plan, BOM and MRP Control sheets from uploaded Excel file"""
        try:
            # Read Excel file
            excel_file = pd.ExcelFile(uploaded_file)
            
            # Check if required sheets exist
            required_sheets = ["Plan", "BOM"]
            missing_sheets = [sheet for sheet in required_sheets if sheet not in excel_file.sheet_names]
            
            if missing_sheets:
                st.error(f"❌ الشيتات التالية غير موجودة في الملف: {', '.join(missing_sheets)}")
                return False
            
            self.plan_df = pd.read_excel(excel_file, sheet_name="Plan")
            self.bom_df = pd.read_excel(excel_file, sheet_name="BOM")
            
            # تحميل شيت MRP Contor إذا كان موجوداً (اختياري)
            if "MRP Contor" in excel_file.sheet_names:
                self.mrp_control_df = pd.read_excel(excel_file, sheet_name="MRP Contor")
                st.success("✅ تم تحميل البيانات بنجاح (بما في ذلك MRP Contor)")
            else:
                st.success("✅ تم تحميل البيانات بنجاح (بدون MRP Contor)")
                st.info("ℹ️ لم يتم العثور على شيت 'MRP Contor' - سيتم المتابعة بدونه")
            
            return True
            
        except Exception as e:
            st.error(f"❌ خطأ في تحميل الملف: {e}")
            return False

    def prepare_mrp_control_data(self):
        """تحضير بيانات MRP Contor"""
        if self.mrp_control_df is None:
            return True  # المتابعة بدون MRP Contor
            
        try:
            # تنظيف أعمدة MRP Contor
            self.mrp_control_df.columns = [str(c).strip() for c in self.mrp_control_df.columns]
            
            # البحث عن أعمدة MRP Contor
            cols_lower = {str(c).lower(): c for c in self.mrp_control_df.columns}
            
            def find_col(*names):
                for n in names:
                    if n.lower() in cols_lower:
                        return cols_lower[n.lower()]
                return None

            col_material = find_col("material", "code", "component", "item code", "raw_material")
            col_description = find_col("description", "material description", "item description", "component_description")
            col_mrp_control = find_col("mrp contor", "mrp control", "mrp", "control", "controller")
            
            if not col_material:
                st.warning("⚠️ عمود الأكواد غير موجود في شيت MRP Contor - سيتم تجاهل الشيت")
                return True
                
            if not col_mrp_control:
                st.warning("⚠️ عمود MRP Contor غير موجود في شيت MRP Contor - سيتم تجاهل الشيت")
                return True
            
            # بناء قاموس قيم MRP Contor
            mrp_control_count = 0
            for _, row in self.mrp_control_df.iterrows():
                material_code = str(row[col_material]).strip()
                if material_code and material_code != 'nan' and material_code != '':
                    # تخزين قيمة MRP Contor
                    mrp_control_value = row[col_mrp_control]
                    if pd.notna(mrp_control_value):
                        self.mrp_control_values[material_code] = str(mrp_control_value).strip()
                        mrp_control_count += 1
                    
                    # أيضا تخزين الوصف إذا كان متوفرا
                    if col_description and pd.notna(row[col_description]):
                        description = str(row[col_description]).strip()
                        if description and description != '':
                            # الأولوية لأوصاف MRP Contor
                            self.material_descriptions[material_code] = description
            
            st.info(f"✅ تم تحميل {mrp_control_count} قيمة MRP Contor")
            return True
            
        except Exception as e:
            st.warning(f"⚠️ خطأ في تحضير بيانات MRP Contor: {e} - سيتم المتابعة بدونه")
            return True

    def get_mrp_control_value(self, material_code: str) -> str:
        """Get MRP Contor value for material code"""
        return self.mrp_control_values.get(material_code, "")

    def prepare_bom_columns(self) -> tuple:
        """Identify and validate BOM columns"""
        self.bom_df.columns = [str(c).strip() for c in self.bom_df.columns]
        cols_lower = {str(c).lower(): c for c in self.bom_df.columns}
        
        def find_col(*names):
            for n in names:
                if n.lower() in cols_lower:
                    return cols_lower[n.lower()]
            return None

        col_parent = find_col("parent material", "parent", "parent code")
        col_component = find_col("component", "component material", "child")
        col_qty = find_col("component quantity", "component_quantity", "qty", "quantity")
        col_component_description = find_col("component description", "comp description", "component desc")
        col_uom = find_col("component uom", "uom", "unit of measure", "unit")
        
        missing_cols = []
        if not col_parent: missing_cols.append("Parent Material")
        if not col_component: missing_cols.append("Component")
        if not col_qty: missing_cols.append("Quantity")
        
        if missing_cols:
            st.error(f"⚠️ الأعمدة التالية غير موجودة في شيت BOM: {', '.join(missing_cols)}")
            return None, None, None, None, None
        
        st.info(f"✅ تم تحديد الأعمدة: Parent={col_parent}, Component={col_component}, Qty={col_qty}, Component Description={col_component_description}, UoM={col_uom}")
        return col_parent, col_component, col_qty, col_component_description, col_uom

    def clean_bom_data(self, col_parent: str, col_component: str, col_qty: str, col_uom: str):
        """تنظيف بيانات BOM وإزالة التكرار"""
        initial_rows = len(self.bom_df)
        
        # 1. إزالة الصفوف الفارغة تماماً
        self.bom_df = self.bom_df.dropna(how='all')
        
        # 2. إزالة الصفوف التي تفتقد بيانات أساسية
        essential_cols = [col_parent, col_component, col_qty]
        self.bom_df = self.bom_df.dropna(subset=essential_cols)
        
        # 3. إزالة التكرار الكامل (نفس Parent + Component + Qty + UoM)
        self.bom_df = self.bom_df.drop_duplicates(
            subset=[col_parent, col_component, col_qty, col_uom], 
            keep='first'
        )
        
        # 4. إزالة التكرار الجزئي (نفس Parent + Component) - نأخذ آخر تحديث
        self.bom_df = self.bom_df.drop_duplicates(
            subset=[col_parent, col_component], 
            keep='last'
        )
        
        final_rows = len(self.bom_df)
        removed_rows = initial_rows - final_rows
        
        if removed_rows > 0:
            st.info(f"🧹 تم تنظيف بيانات BOM: إزالة {removed_rows} صف (من {initial_rows} إلى {final_rows})")

    def convert_quantity(self, quantity: float, uom: str) -> tuple:
        """Convert quantity from G to KG only and return standardized UoM"""
        uom_clean = str(uom).strip().upper()
        
        # تحويل الجرام فقط إلى كيلوجرام
        if uom_clean in ['G', 'GR', 'GRAM', 'GRAMS']:
            return quantity * 0.001, 'KG'  # تحويل من جرام إلى كيلوجرام
        else:
            # باقي الوحدات تبقى كما هي
            return quantity, uom_clean

    def build_bom_relations(self, col_parent: str, col_component: str, col_qty: str, col_component_description: str, col_uom: str) -> bool:
        """Build BOM parent-component relationships and store descriptions"""
        try:
            # 🔥 تنظيف البيانات وإزالة التكرار قبل المعالجة
            self.clean_bom_data(col_parent, col_component, col_qty, col_uom)
            
            # بناء قاموس أوصاف المواد ووحدات القياس من شيت BOM
            for _, row in self.bom_df.iterrows():
                material_code = str(row[col_component]).strip()
                
                # تخزين الوصف من BOM (إذا لم يكن موجوداً في MRP Contor)
                if col_component_description and pd.notna(row[col_component_description]):
                    description = str(row[col_component_description]).strip()
                    if material_code and description and material_code != 'nan' and material_code not in self.material_descriptions:
                        self.material_descriptions[material_code] = description
                
                # تخزين وحدة القياس الأصلية
                if col_uom and pd.notna(row[col_uom]):
                    uom = str(row[col_uom]).strip()
                    if material_code and uom and material_code != 'nan':
                        self.material_uoms[material_code] = uom
                        # تحديد الوحدة الموحدة
                        converted_qty, standardized_uom = self.convert_quantity(1.0, uom)
                        self.standardized_uoms[material_code] = standardized_uom
                
                # أيضا تخزين وصف ووحدات القياس للمواد الأب
                parent_code = str(row[col_parent]).strip()
                if col_component_description and pd.notna(row[col_component_description]):
                    parent_desc = str(row[col_component_description]).strip()
                    if parent_code and parent_desc and parent_code != 'nan' and parent_code not in self.material_descriptions:
                        self.material_descriptions[parent_code] = parent_desc
                
                if col_uom and pd.notna(row[col_uom]):
                    parent_uom = str(row[col_uom]).strip()
                    if parent_code and parent_uom and parent_code != 'nan':
                        self.material_uoms[parent_code] = parent_uom
                        converted_qty, standardized_uom = self.convert_quantity(1.0, parent_uom)
                        self.standardized_uoms[parent_code] = standardized_uom
            
            # بناء علاقات BOM مع تحويل الوحدات
            for _, row in self.bom_df.iterrows():
                parent = str(row[col_parent]).strip()
                comp = str(row[col_component]).strip()
                
                # Skip empty rows
                if not parent or not comp or parent.lower() == 'nan' or comp.lower() == 'nan':
                    continue
                
                # Handle quantity conversion
                try:
                    qty_val = row[col_qty]
                    if pd.isna(qty_val):
                        continue
                    qty = float(str(qty_val).replace(",", "."))
                    
                    # تحويل الكمية بناءً على وحدة القياس
                    if col_uom and pd.notna(row[col_uom]):
                        uom = str(row[col_uom]).strip()
                        converted_qty, _ = self.convert_quantity(qty, uom)
                    else:
                        converted_qty = qty
                        
                except (ValueError, TypeError):
                    st.warning(f"⚠️ كمية غير صالحة لتخطيط {parent} -> {comp}: {qty_val}")
                    continue
                    
                if converted_qty > 0:
                    self.relations[parent].append((comp, converted_qty))
            
            st.success(f"✅ تم بناء علاقات BOM لـ {len(self.relations)} مادة أب")
            st.info(f"✅ تم تخزين أوصاف لـ {len(self.material_descriptions)} مادة")
            st.info(f"✅ تم تخزين وحدات قياس لـ {len(self.material_uoms)} مادة")
            
            # عرض إحصائيات التحويل
            g_materials = [code for code, uom in self.material_uoms.items() 
                          if str(uom).upper() in ['G', 'GR', 'GRAM', 'GRAMS']]
            if g_materials:
                st.info(f"🔁 سيتم تحويل {len(g_materials)} مادة من الجرام إلى الكيلوجرام")
            
            return True
            
        except Exception as e:
            st.error(f"❌ خطأ في بناء علاقات BOM: {e}")
            return False

    def get_material_description(self, material_code: str) -> str:
        """Get description for material code, return empty if not found"""
        return self.material_descriptions.get(material_code, "")

    def get_standardized_uom(self, material_code: str) -> str:
        """Get standardized UoM for material code"""
        return self.standardized_uoms.get(material_code, self.material_uoms.get(material_code, ""))

    @lru_cache(maxsize=None)
    def explode_unit(self, item_code: str) -> dict:
        """Recursively explode BOM to raw materials"""
        item = str(item_code).strip()
        
        # If item starts with '1' or has no components, treat as raw material
        if item.startswith("1") or item not in self.relations or not self.relations[item]:
            return {item: 1.0}
            
        total = defaultdict(float)
        for comp, qty in self.relations[item]:
            sub_map = self.explode_unit(comp)
            for material, quantity in sub_map.items():
                total[material] += quantity * qty
                
        return dict(total)

    def calculate_requirements(self) -> pd.DataFrame:
        """Calculate material requirements based on production plan"""
        plan_cols = list(self.plan_df.columns)
        
        # Identify FG and month columns
        if "Material Description" in plan_cols:
            fg_col = plan_cols[0]
            month_cols = plan_cols[2:]
        else:
            fg_col = plan_cols[0]
            month_cols = plan_cols[1:]
        
        st.info(f"✅ تم تحديد أعمدة الخطة: FG={fg_col}, الشهور={len(month_cols)}")
        
        results = defaultdict(lambda: defaultdict(float))
        material_codes = set()  # لتخزين جميع أكواد المواد
        
        # Progress bar
        progress_bar = st.progress(0)
        total_items = len(self.plan_df)
        
        for idx, (_, prow) in enumerate(self.plan_df.iterrows()):
            fg = str(prow[fg_col]).strip()
            if not fg or fg.lower() in ['nan', 'none', '']:
                continue
                
            bom_map = self.explode_unit(fg)
            
            for month in month_cols:
                try:
                    planned_qty = prow[month]
                    if pd.isna(planned_qty):
                        continue
                    planned = float(str(planned_qty).replace(",", "."))
                except (ValueError, TypeError):
                    continue
                    
                if planned == 0:
                    continue
                    
                for raw_material, per_unit in bom_map.items():
                    results[raw_material][month] += planned * per_unit
                    material_codes.add(raw_material)
            
            # Update progress
            progress_bar.progress((idx + 1) / total_items)
        
        st.success(f"✅ تم معالجة {len(self.plan_df)} من مواد التخطيط")
        
        # Create output DataFrame with descriptions and STANDARDIZED UoM
        raw_list = sorted(material_codes)
        
        # إضافة أعمدة الوصف ووحدة القياس الموحدة و MRP Contor
        descriptions = [self.get_material_description(material) for material in raw_list]
        standardized_uoms = [self.get_standardized_uom(material) for material in raw_list]
        mrp_controls = [self.get_mrp_control_value(material) for material in raw_list]
        
        # إنشاء DataFrame النهائي
        out_df = pd.DataFrame({
            'Raw_Material': raw_list,
            'Component_Description': descriptions,
            'UoM': standardized_uoms,  # استخدام الوحدة الموحدة
            'MRP_Contor': mrp_controls  # إضافة عمود MRP Contor
        })
        
        # إضافة أعمدة الشهور
        for month in month_cols:
            month_data = [results[material].get(month, 0.0) for material in raw_list]
            out_df[str(month)] = month_data
        
        return out_df

    def calculate_manufacturing_quantities(self):
        """حساب كميات التصنيع للمكونات الوسيطة تلقائياً"""
        try:
            # حساب كميات التصنيع من الخطة للمواد الموجودة في BOM كمواد أب
            for material in self.relations.keys():
                if material in self.plan_df.iloc[:, 0].values:
                    material_plan = self.plan_df[self.plan_df.iloc[:, 0] == material]
                    if not material_plan.empty:
                        # جمع جميع الكميات من أعمدة الشهور
                        month_cols = self.plan_df.columns[2:] if "Material Description" in self.plan_df.columns else self.plan_df.columns[1:]
                        total_manufacturing = material_plan[month_cols].sum().sum()
                        if total_manufacturing > 0:
                            self.manufacturing_quantities[material] = total_manufacturing
            
            st.info(f"✅ تم حساب كميات التصنيع لـ {len(self.manufacturing_quantities)} مكون وسيط")
            return True
        except Exception as e:
            st.warning(f"⚠️ خطأ في حساب كميات التصنيع: {e}")
            return False

    def calculate_all_levels_requirements(self):
        """حساب الكميات المطلوبة لجميع مستويات الـ BOM"""
        try:
            # تحديد أعمدة الشهور من الخطة
            month_cols = self.plan_df.columns[2:] if "Material Description" in self.plan_df.columns else self.plan_df.columns[1:]
            
            # نتائج جميع المستويات
            all_levels_results = defaultdict(lambda: defaultdict(float))
            
            # معالجة كل مادة في الخطة
            for _, row in self.plan_df.iterrows():
                parent = str(row.iloc[0]).strip()
                if not parent:
                    continue
                
                # حساب الكميات لكل شهر
                for month in month_cols:
                    try:
                        planned_qty = row[month]
                        if pd.isna(planned_qty) or planned_qty == 0:
                            continue
                        planned = float(str(planned_qty).replace(",", "."))
                    except (ValueError, TypeError):
                        continue
                    
                    # إضافة المادة الأصلية (المستوى 0)
                    all_levels_results[parent][month] += planned
                    
                    # حساب الكميات لجميع المستويات باستخدام BOM
                    self._calculate_component_requirements(parent, planned, month, all_levels_results)
            
            # إنشاء DataFrame لجميع المستويات
            all_materials = sorted(all_levels_results.keys())
            
            all_levels_data = []
            for material in all_materials:
                row_data = {
                    'Material_Code': material,
                    'Material_Description': self.get_material_description(material),
                    'Standardized_UoM': self.get_standardized_uom(material),
                    'MRP_Contor': self.get_mrp_control_value(material),
                    'Level': self._get_material_level(material),
                    'Is_Raw_Material': material.startswith("1") or material not in self.relations or not self.relations[material],
                    'Total_Required': sum(all_levels_results[material].values())
                }
                
                # إضافة الكميات لكل شهر
                for month in month_cols:
                    row_data[str(month)] = all_levels_results[material].get(month, 0.0)
                
                all_levels_data.append(row_data)
            
            all_levels_df = pd.DataFrame(all_levels_data)
            
            # إعادة ترتيب الأعمدة
            base_cols = ['Material_Code', 'Material_Description', 'Standardized_UoM', 'MRP_Contor', 'Level', 'Is_Raw_Material', 'Total_Required']
            month_cols_sorted = [str(col) for col in month_cols]
            all_cols = base_cols + month_cols_sorted
            
            all_levels_df = all_levels_df[all_cols]
            all_levels_df = all_levels_df.sort_values(['Level', 'Material_Code'])
            
            st.info(f"✅ تم حساب الكميات لـ {len(all_levels_df)} مادة في جميع مستويات الـ BOM")
            return all_levels_df
            
        except Exception as e:
            st.error(f"❌ خطأ في حساب جميع المستويات: {e}")
            return pd.DataFrame()

    def _calculate_component_requirements(self, parent, parent_qty, month, results_dict):
        """دالة مساعدة لحساب متطلبات المكونات بشكل متكرر"""
        if parent not in self.relations:
            return
        
        for comp, comp_qty in self.relations[parent]:
            required_qty = parent_qty * comp_qty
            results_dict[comp][month] += required_qty
            # استدعاء متكرر للمكونات التالية
            self._calculate_component_requirements(comp, required_qty, month, results_dict)

    def _get_material_level(self, material_code):
        """تحديد مستوى المادة في هيكل الـ BOM"""
        # المواد في الخطة هي المستوى 0
        if material_code in set(self.plan_df.iloc[:, 0].astype(str).str.strip()):
            return 0
        
        # البحث عن المستوى في العلاقات
        def find_level(current_material, visited=None):
            if visited is None:
                visited = set()
            
            if current_material in visited:
                return -1  # منع التكرار اللانهائي
            visited.add(current_material)
            
            # إذا كانت المادة موجودة في الخطة، فهي المستوى 0
            if current_material in set(self.plan_df.iloc[:, 0].astype(str).str.strip()):
                return 1
            
            # البحث عن الآباء
            for parent, components in self.relations.items():
                for comp, _ in components:
                    if comp == current_material:
                        parent_level = find_level(parent, visited)
                        if parent_level >= 0:
                            return parent_level + 1
            
            return -1  # لم يتم العثور على مسار إلى الخطة
        
        level = find_level(material_code)
        return level if level >= 0 else 999  # مستوى عالي إذا لم يتم العثور على مسار

    def generate_bom_level1_sheet(self):
        """إنشاء شيت للمكونات الموجودة في المستوى الأول فقط مع الكميات المطلوبة حسب التاريخ"""
        try:
            # استخدام دالة حساب جميع المستويات ثم تصفية المستوى الأول فقط
            all_levels_df = self.calculate_all_levels_requirements()
            if all_levels_df.empty:
                return pd.DataFrame()
            
            # تصفية المستوى الأول فقط (Level = 1)
            level1_df = all_levels_df[all_levels_df['Level'] == 1].copy()
            
            st.info(f"✅ تم تحديد {len(level1_df)} مكون في المستوى الأول")
            return level1_df
            
        except Exception as e:
            st.error(f"❌ خطأ في إنشاء شيت المستوى الأول: {e}")
            return pd.DataFrame()

    def create_monthly_summary(self):
        """إنشاء ملخص شهري للكميات حسب نوع الأمر"""
        try:
            if "Order Type" not in self.plan_df.columns:
                st.warning("⚠️ عمود 'Order Type' غير موجود في شيت Plan")
                return pd.DataFrame()
            
            # تحديد أعمدة الشهور
            date_cols = [c for c in self.plan_df.columns if isinstance(c, (datetime.datetime, pd.Timestamp))]
            if not date_cols:
                # إذا لم تكن هناك تواريخ، استخدم الأعمدة الرقمية بعد العمودين الأولين
                date_cols = self.plan_df.columns[2:] if "Material Description" in self.plan_df.columns else self.plan_df.columns[1:]
            
            orders_summary = self.plan_df.melt(
                id_vars=["Material", "Order Type"],
                value_vars=date_cols,
                var_name="Month",
                value_name="Quantity"
            )
            
            # تحويل الشهور إلى أسماء إذا كانت تواريخ
            try:
                orders_summary["Month"] = pd.to_datetime(orders_summary["Month"], errors="coerce")
                orders_summary = orders_summary.dropna(subset=["Month"])
                orders_summary["Month"] = orders_summary["Month"].dt.month_name()
            except:
                pass  # إذا لم تكن تواريخ، استخدم الأسماء كما هي

            orders_grouped = orders_summary.groupby(
                ["Month", "Order Type"]
            ).agg({"Quantity": "sum"}).reset_index()

            pivot_df = orders_grouped.pivot_table(
                index="Month", columns="Order Type", values="Quantity", aggfunc="sum", fill_value=0
            ).reset_index()

            pivot_df["الإجمالي"] = pivot_df.sum(axis=1, numeric_only=True)
            
            # حساب النسب المئوية
            if 'E' in pivot_df.columns:
                pivot_df["E%"] = (pivot_df["E"] / pivot_df["الإجمالي"] * 100).round(1).astype(str) + "%"
            if 'L' in pivot_df.columns:
                pivot_df["L%"] = (pivot_df["L"] / pivot_df["الإجمالي"] * 100).round(1).astype(str) + "%"
            
            return pivot_df
            
        except Exception as e:
            st.error(f"❌ خطأ في إنشاء الملخص الشهري: {e}")
            return pd.DataFrame()

    def run(self):
        """Main execution method"""
        # File upload section
        st.header("📁 رفع ملف الخطة")
        
        uploaded_file = st.file_uploader(
            "اختر ملف Excel الذي يحتوي على شيت Plan وBOM (واختياري: MRP Contor)",
            type=["xlsx", "xls"],
            help="يجب أن يحتوي الملف على شيتين: 'Plan' و 'BOM' - واختياري: 'MRP Contor'"
        )
        
        if uploaded_file is not None:
            try:
                # Load data
                if not self.load_data(uploaded_file):
                    return
                
                # Show data preview
                if self.mrp_control_df is not None:
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        st.subheader("معاينة بيانات الخطة (Plan)")
                        st.dataframe(self.plan_df.head(), use_container_width=True)
                    
                    with col2:
                        st.subheader("معاينة بيانات BOM")
                        st.dataframe(self.bom_df.head(), use_container_width=True)
                    
                    with col3:
                        st.subheader("معاينة بيانات MRP Contor")
                        st.dataframe(self.mrp_control_df.head(), use_container_width=True)
                else:
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.subheader("معاينة بيانات الخطة (Plan)")
                        st.dataframe(self.plan_df.head(), use_container_width=True)
                    
                    with col2:
                        st.subheader("معاينة بيانات BOM")
                        st.dataframe(self.bom_df.head(), use_container_width=True)
                
                # Process MRP Control data first
                if not self.prepare_mrp_control_data():
                    return
                
                # Process BOM
                col_parent, col_component, col_qty, col_component_description, col_uom = self.prepare_bom_columns()
                if not all([col_parent, col_component, col_qty]):
                    return
                
                if not self.build_bom_relations(col_parent, col_component, col_qty, col_component_description, col_uom):
                    return
                
                # Show material info sample
                if self.material_descriptions or self.material_uoms:
                    st.subheader("📝 عينة من بيانات المواد")
                    sample_data = []
                    materials = list(self.material_descriptions.keys())[:10]
                    for material in materials:
                        original_uom = self.material_uoms.get(material, '')
                        standardized_uom = self.get_standardized_uom(material)
                        mrp_control = self.get_mrp_control_value(material)
                        sample_data.append({
                            'كود المادة': material,
                            'وصف المكون': self.material_descriptions.get(material, ''),
                            'الوحدة الأصلية': original_uom,
                            'الوحدة الموحدة': standardized_uom,
                            'MRP Contor': mrp_control
                        })
                    if sample_data:
                        sample_df = pd.DataFrame(sample_data)
                        st.dataframe(sample_df, use_container_width=True)
                    
                    if len(self.material_descriptions) > 10:
                        st.info(f"... وعرض {len(self.material_descriptions) - 10} مادة أخرى")
                
                # Calculate manufacturing quantities
                self.calculate_manufacturing_quantities()
                
                # Show manufacturing quantities
                if self.manufacturing_quantities:
                    st.subheader("🏭 كميات التصنيع للمكونات الوسيطة")
                    manuf_data = []
                    for material, qty in list(self.manufacturing_quantities.items())[:10]:
                        manuf_data.append({
                            'المادة': material,
                            'الوصف': self.get_material_description(material),
                            'كمية التصنيع': f"{qty:,.0f}",
                            'MRP Contor': self.get_mrp_control_value(material)
                        })
                    if manuf_data:
                        manuf_df = pd.DataFrame(manuf_data)
                        st.dataframe(manuf_df, use_container_width=True)
                    
                    if len(self.manufacturing_quantities) > 10:
                        st.info(f"... وعرض {len(self.manufacturing_quantities) - 10} مكون آخر")
                
                # Calculate requirements
                if st.button("🚀 حساب متطلبات المواد", type="primary"):
                    with st.spinner("جاري حساب متطلبات المواد..."):
                        requirements_df = self.calculate_requirements()
                        all_levels_df = self.calculate_all_levels_requirements()
                        level1_df = self.generate_bom_level1_sheet()
                        monthly_summary = self.create_monthly_summary()
                    
                    # Display results
                    st.header("📊 نتائج متطلبات المواد")
                    
                    # إحصائيات سريعة
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("عدد المواد الخام", len(requirements_df))
                    with col2:
                        total_req = requirements_df.select_dtypes(include=['number']).sum().sum()
                        st.metric("إجمالي المتطلبات", f"{total_req:,.2f}")
                    with col3:
                        kg_materials = (requirements_df['UoM'] == 'KG').sum()
                        st.metric("المواد بالكيلوجرام", kg_materials)
                    with col4:
                        materials_with_mrp = (requirements_df['MRP_Contor'] != '').sum()
                        st.metric("مواد ذات MRP Contor", f"{materials_with_mrp}/{len(requirements_df)}")
                    
                    # عرض البيانات مع التنسيق
                    st.subheader("متطلبات المواد الخام")
                    st.dataframe(requirements_df, use_container_width=True)
                    
                    # عرض جميع مستويات الـ BOM
                    if not all_levels_df.empty:
                        st.subheader("🏗️ جميع مستويات الـ BOM مع الكميات المطلوبة")
                        
                        # إحصائيات سريعة لجميع المستويات
                        col1, col2, col3, col4 = st.columns(4)
                        with col1:
                            st.metric("إجمالي المواد", len(all_levels_df))
                        with col2:
                            total_all_levels = all_levels_df['Total_Required'].sum()
                            st.metric("إجمالي الكميات", f"{total_all_levels:,.0f}")
                        with col3:
                            max_level = all_levels_df['Level'].max()
                            st.metric("أعلى مستوى", max_level)
                        with col4:
                            raw_count = all_levels_df['Is_Raw_Material'].sum()
                            st.metric("مواد خام", f"{raw_count}/{len(all_levels_df)}")
                        
                        # عرض حسب المستويات
                        for level in sorted(all_levels_df['Level'].unique()):
                            level_data = all_levels_df[all_levels_df['Level'] == level]
                            st.subheader(f"📋 المستوى {level} ({len(level_data)} مادة)")
                            st.dataframe(level_data, use_container_width=True)
                    
                    # عرض مكونات المستوى الأول مع الكميات
                    if not level1_df.empty:
                        st.subheader("📋 المكونات في المستوى الأول مع الكميات المطلوبة")
                        st.dataframe(level1_df, use_container_width=True)
                    
                    # عرض الملخص الشهري
                    if not monthly_summary.empty:
                        st.subheader("📅 الملخص الشهري للكميات")
                        
                        # عرض كجدول HTML منسق
                        html_table = "<table border='1' style='border-collapse: collapse; width:100%; text-align:center; color:black;'>"
                        html_table += "<tr style='background-color:#d9d9d9; color:blue;'><th>الشهر</th><th>E</th><th>L</th><th>الإجمالي</th><th>E%</th><th>L%</th></tr>"

                        for idx, row in monthly_summary.iterrows():
                            bg_color = "#f2f2f2" if idx % 2 == 0 else "#ffffff"
                            html_table += f"<tr style='background-color:{bg_color};'>"
                            html_table += f"<td style='color:blue;'>{row['Month']}</td>"
                            html_table += f"<td>{int(row.get('E',0))}</td>"
                            html_table += f"<td>{int(row.get('L',0))}</td>"
                            html_table += f"<td>{int(row.get('الإجمالي',0))}</td>"
                            html_table += f"<td>{row.get('E%','')}</td>"
                            html_table += f"<td>{row.get('L%','')}</td>"
                            html_table += "</tr>"

                        html_table += "</table>"
                        st.markdown(f"<div style='direction:rtl;'>{html_table}</div>", unsafe_allow_html=True)
                        
                        # رسم بياني
                        st.subheader("📊 رسم بياني للكميات الشهرية")
                        numeric_cols = [c for c in monthly_summary.columns if c not in ["Month", "الإجمالي", "E%", "L%"]]
                        monthly_summary[numeric_cols] = monthly_summary[numeric_cols].apply(pd.to_numeric, errors="coerce")
                        
                        fig = px.bar(
                            monthly_summary,
                            x="Month",
                            y=numeric_cols,
                            barmode="group",
                            text_auto=True,
                            title="توزيع الكميات حسب نوع الأمر",
                            template="streamlit"
                        )
                        st.plotly_chart(fig, use_container_width=True)
                    
                    # تحميل النتائج
                    self.download_results(requirements_df, all_levels_df, level1_df, monthly_summary)
                    
            except Exception as e:
                st.error(f"❌ حدث خطأ غير متوقع: {e}")

    def download_results(self, requirements_df, all_levels_df, level1_df, monthly_summary):
        """Handle downloading results"""
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            self.plan_df.to_excel(writer, sheet_name="Plan", index=False)
            self.bom_df.to_excel(writer, sheet_name="BOM", index=False)
            if self.mrp_control_df is not None:
                self.mrp_control_df.to_excel(writer, sheet_name="MRP_Contor", index=False)
            requirements_df.to_excel(writer, sheet_name="RawMaterial_Requirements", index=False)
            
            # إضافة الشيتات الجديدة
            if not all_levels_df.empty:
                all_levels_df.to_excel(writer, sheet_name="All_BOM_Levels", index=False)
            if not level1_df.empty:
                level1_df.to_excel(writer, sheet_name="Level1_Components", index=False)
            if not monthly_summary.empty:
                monthly_summary.to_excel(writer, sheet_name="Monthly_Summary", index=False)
            
            # إضافة شيت كميات التصنيع
            if self.manufacturing_quantities:
                manuf_df = pd.DataFrame([
                    {
                        'Material': mat,
                        'Description': self.get_material_description(mat),
                        'Manufacturing_Quantity': qty,
                        'MRP_Contor': self.get_mrp_control_value(mat)
                    }
                    for mat, qty in self.manufacturing_quantities.items()
                ])
                manuf_df.to_excel(writer, sheet_name="Manufacturing_Quantities", index=False)
        
        output.seek(0)
        
        st.download_button(
            label="📥 تحميل النتائج كملف Excel",
            data=output,
            file_name=f"MRP_Results_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        st.balloons()

# Run the application
if __name__ == "__main__":
    calculator = MRPCalculator()
    calculator.run()

# --- التذييل ---
st.markdown(
    """
    <hr>
    <div style="text-align:center; direction:rtl; font-size:14px; color:gray;">
        ✨ تم التنفيذ بواسطة <b>م / رضا رشدي</b> – جميع الحقوق محفوظة © 2025 ✨
    </div>
    """,
    unsafe_allow_html=True
)
