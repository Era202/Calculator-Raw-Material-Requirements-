import pandas as pd
from pathlib import Path
from collections import defaultdict
from functools import lru_cache
import streamlit as st
import io

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
        self.material_descriptions = {}  # تخزين أوصاف المواد
        self.material_uoms = {}  # تخزين وحدات القياس للمواد
        self.standardized_uoms = {}  # تخزين الوحدات الموحدة
        
    def load_data(self, uploaded_file) -> bool:
        """Load Plan and BOM sheets from uploaded Excel file"""
        try:
            # Read Excel file
            excel_file = pd.ExcelFile(uploaded_file)
            
            # Check if required sheets exist
            if "Plan" not in excel_file.sheet_names:
                st.error("❌ شيت 'Plan' غير موجود في الملف")
                return False
            if "BOM" not in excel_file.sheet_names:
                st.error("❌ شيت 'BOM' غير موجود في الملف")
                return False
            
            self.plan_df = pd.read_excel(excel_file, sheet_name="Plan")
            self.bom_df = pd.read_excel(excel_file, sheet_name="BOM")
            
            st.success("✅ تم تحميل البيانات بنجاح")
            return True
            
        except Exception as e:
            st.error(f"❌ خطأ في تحميل الملف: {e}")
            return False

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
                
                # تخزين الوصف
                if col_component_description and pd.notna(row[col_component_description]):
                    description = str(row[col_component_description]).strip()
                    if material_code and description and material_code != 'nan':
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
                    if parent_code and parent_desc and parent_code != 'nan':
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

    # ... باقي الدوال كما هي بدون تغيير
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
        
        # إضافة أعمدة الوصف ووحدة القياس الموحدة
        descriptions = [self.get_material_description(material) for material in raw_list]
        standardized_uoms = [self.get_standardized_uom(material) for material in raw_list]
        
        # إنشاء DataFrame النهائي
        out_df = pd.DataFrame({
            'Raw_Material': raw_list,
            'Component_Description': descriptions,
            'UoM': standardized_uoms  # استخدام الوحدة الموحدة
        })
        
        # إضافة أعمدة الشهور
        for month in month_cols:
            month_data = [results[material].get(month, 0.0) for material in raw_list]
            out_df[str(month)] = month_data
        
        return out_df

    def run(self):
        """Main execution method"""
        # File upload section
        st.header("📁 رفع ملف الخطة")
        
        uploaded_file = st.file_uploader(
            "اختر ملف Excel الذي يحتوي على شيت Plan وBOM",
            type=["xlsx", "xls"],
            help="يجب أن يحتوي الملف على شيتين: 'Plan' و 'BOM'"
        )
        
        if uploaded_file is not None:
            try:
                # Load data
                if not self.load_data(uploaded_file):
                    return
                
                # Show data preview
                col1, col2 = st.columns(2)
                
                with col1:
                    st.subheader("معاينة بيانات الخطة (Plan)")
                    st.dataframe(self.plan_df.head(), use_container_width=True)
                
                with col2:
                    st.subheader("معاينة بيانات BOM")
                    st.dataframe(self.bom_df.head(), use_container_width=True)
                
                # Process BOM
                col_parent, col_component, col_qty, col_component_description, col_uom = self.prepare_bom_columns()
                if not all([col_parent, col_component, col_qty]):
                    return
                
                if not self.build_bom_relations(col_parent, col_component, col_qty, col_component_description, col_uom):
                    return
                
                # Show material info sample
                if self.material_descriptions or self.material_uoms:
                    st.subheader("📝 عينة من بيانات المواد (قبل التحويل)")
                    sample_data = []
                    materials = list(self.material_descriptions.keys())[:10]
                    for material in materials:
                        original_uom = self.material_uoms.get(material, '')
                        standardized_uom = self.get_standardized_uom(material)
                        sample_data.append({
                            'كود المادة': material,
                            'وصف المكون': self.material_descriptions.get(material, ''),
                            'الوحدة الأصلية': original_uom,
                            'الوحدة الموحدة': standardized_uom
                        })
                    if sample_data:
                        sample_df = pd.DataFrame(sample_data)
                        st.dataframe(sample_df, use_container_width=True)
                    
                    if len(self.material_descriptions) > 10:
                        st.info(f"... وعرض {len(self.material_descriptions) - 10} مادة أخرى")
                
                # Calculate requirements
                if st.button("🚀 حساب متطلبات المواد", type="primary"):
                    with st.spinner("جاري حساب متطلبات المواد..."):
                        requirements_df = self.calculate_requirements()
                    
                    # Display results
                    st.header("📊 نتائج متطلبات المواد")
                    
                    # إحصائيات سريعة
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("عدد المواد الخام", len(requirements_df))
                    with col2:
                        total_req = requirements_df.select_dtypes(include=['number']).sum().sum()
                        st.metric("إجمالي المتطلبات", f"{total_req:,.2f}")
                    with col3:
                        kg_materials = (requirements_df['UoM'] == 'KG').sum()
                        st.metric("المواد بالكيلوجرام", kg_materials)
                    
                    # عرض البيانات مع التنسيق
                    st.dataframe(requirements_df, use_container_width=True)
                    
                    # تحميل النتائج
                    self.download_results(requirements_df)
                    
            except Exception as e:
                st.error(f"❌ حدث خطأ غير متوقع: {e}")

    def download_results(self, requirements_df):
        """Handle downloading results"""
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            self.plan_df.to_excel(writer, sheet_name="Plan", index=False)
            self.bom_df.to_excel(writer, sheet_name="BOM", index=False)
            requirements_df.to_excel(writer, sheet_name="RawMaterial_Requirements", index=False)
        
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
