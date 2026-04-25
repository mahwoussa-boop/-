import streamlit as st
import pandas as pd
import json
import io
import os
import re
import time
from datetime import datetime
import google.generativeai as genai
from openpyxl import load_workbook

# ─── PAGE CONFIG ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="مهووس | معالج المنتجات الذكي",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    body, .stApp { direction: rtl; font-family: 'Segoe UI', Tahoma, Arial, sans-serif; }
    .brand-card {
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
        border: 1px solid #0f3460;
        border-radius: 12px;
        padding: 20px;
        color: white;
        margin-bottom: 16px;
    }
    .stat-box {
        background: rgba(255,255,255,0.05);
        border-radius: 8px;
        padding: 12px;
        text-align: center;
    }
    .section-header {
        background: linear-gradient(90deg, #0f3460, #533483);
        color: white;
        padding: 10px 16px;
        border-radius: 8px;
        font-weight: bold;
        margin: 12px 0;
    }
</style>
""", unsafe_allow_html=True)

# ─── SYSTEM INSTRUCTION ──────────────────────────────────────────────────────
SYSTEM_INSTRUCTION = """أنت خبير عالمي في كتابة أوصاف منتجات العطور محسّنة لمحركات البحث التقليدية (Google SEO) ومحركات بحث الذكاء الصناعي (GEO/AIO). تعمل حصرياً لمتجر "مهووس" (Mahwous) - الوجهة الأولى للعطور الفاخرة في السعودية.

## هويتك ومهمتك
أنت خبير عطور محترف مع 15+ سنة خبرة في صناعة العطور الفاخرة، متخصص في SEO و Generative Engine Optimization (GEO)، وكاتب محتوى عربي بارع بأسلوب راقٍ ودود وعاطفي وتسويقي مقنع. تمثل صوت متجر "مهووس" بكل احترافية وثقة.

مهمتك: كتابة أوصاف منتجات عطور شاملة واحترافية ومحسّنة لتحقيق:
1. تصدر نتائج البحث في Google
2. الظهور في إجابات محركات بحث الذكاء الاصطناعي (ChatGPT, Gemini, Perplexity)
3. زيادة معدل التحويل بنسبة 40-60%
4. تعزيز ثقة العملاء (E-E-A-T: Experience, Expertise, Authoritativeness, Trustworthiness)

## قواعد الكلمات المفتاحية (إلزامية)
**المستوى 1: الكلمة الرئيسية** - الصيغة: "عطر [الماركة] [اسم العطر] [التركيز] [الحجم] [للجنس]" - التكرار: 5-7 مرات في 1200 كلمة
**المستوى 2: الكلمات الثانوية (3 كلمات)** - التكرار: 3-5 مرات لكل كلمة
**المستوى 3: الكلمات الدلالية (10-15 كلمة)** - صفات، مكونات، أحاسيس، مناسبات - التكرار: 2-3 مرات لكل كلمة
**المستوى 4: الكلمات الحوارية (5-8 عبارات)** - في قسم FAQ

## بنية الوصف الإلزامية (1200-1500 كلمة):
1. الفقرة الافتتاحية (100-150 كلمة) - الكلمة الرئيسية في أول 50 كلمة
2. تفاصيل المنتج (نقاط نقطية)
3. رحلة العطر: الهرم العطري (200-250 كلمة) - وصف حسي للـ Top/Heart/Base Notes
4. لماذا تختار هذا العطر؟ (200-250 كلمة) - 4-6 نقاط بمزايا وفوائد
5. متى وأين ترتدي هذا العطر؟ (150-200 كلمة) - الفصول، الأوقات، المناسبات
6. لمسة خبير من مهووس (200-250 كلمة) - تحليل حسي، ثبات، مقارنات [إلزامي]
7. الأسئلة الشائعة FAQ (250-300 كلمة) - 6-8 أسئلة بكلمات مفتاحية حوارية
8. اكتشف أكثر من مهووس - روابط داخلية وخارجية
9. الفقرة الختامية "عالمك العطري يبدأ من مهووس"

## الأسلوب: راقٍ (40%) + ودود (25%) + عاطفي (20%) + تسويقي (15%)
لا تستخدم الإيموجي. استخدم Bold للكلمات المهمة. اكتب بطبيعية دون حشو.

## قواعد خاصة للمعالجة الآلية:
- عند طلب JSON، أعد فقط JSON صارم بدون أي نص إضافي أو markdown code blocks
- لا تبدأ الرد بـ ``` أو json أو أي نص قبل القوس الأول {
- أعد JSON واحد فقط يبدأ بـ { وينتهي بـ }
"""

# ─── COLUMN DETECTION ────────────────────────────────────────────────────────
ARABIC_COL_KEYS = {
    'name':          ['اسم المنتج', 'اسم'],
    'type':          ['نوع المنتج'],
    'category':      ['فئة المنتج', 'فئة'],
    'images':        ['صورة المنتج', 'صورة'],
    'option_name':   ['اسم خيار'],
    'option_value':  ['اسم الخيار'],
    'price':         ['سعر المنتج', 'سعر'],
    'quantity':      ['الكمية'],
    'description':   ['الوصف'],
    'accepts_orders':['هل يقبل'],
    'sku':           ['sku', 'رمز المنتج'],
    'barcode':       ['الباركود', 'رمز الباركود'],
    'brand':         ['الماركة'],
    'status':        ['حالة المنتج', 'حالة'],
}

FALLBACK_POSITIONS = {
    'name': 1, 'type': 2, 'category': 3, 'images': 4,
    'price': 7, 'quantity': 8, 'description': 9,
    'sku': 11, 'brand': 22, 'status': 24,
}


def find_col(df: pd.DataFrame, key: str) -> str | None:
    """Find column by Arabic name keywords or positional fallback."""
    keywords = ARABIC_COL_KEYS.get(key, [])
    for col in df.columns:
        col_str = str(col)
        for kw in keywords:
            if kw.lower() in col_str.lower():
                return col
    # Positional fallback
    if key in FALLBACK_POSITIONS:
        idx = FALLBACK_POSITIONS[key]
        cols = list(df.columns)
        if idx < len(cols):
            return cols[idx]
    return None


def get_brand_col(df: pd.DataFrame) -> str | None:
    """Find brand column: try name match, then content pattern (values contain '|')."""
    col = find_col(df, 'brand')
    if col:
        return col
    for c in df.columns:
        sample = df[c].dropna().astype(str).head(30)
        pipe_count = sample.str.contains(r'\|').sum()
        avg_len = sample.str.len().mean()
        if pipe_count > 5 and avg_len < 60:
            return c
    return None


def is_tester(name: str) -> bool:
    if not isinstance(name, str):
        return False
    return any(t in name.lower() for t in ['تستر', 'tester', 'testr'])


def calc_tester_price(original_price: float) -> float:
    if original_price >= 1000:
        return max(original_price - 150, 0)
    return max(original_price - 70, 0)


def load_products(file) -> pd.DataFrame:
    """Load products file, auto-detecting header row."""
    name = file.name.lower()
    if name.endswith('.csv'):
        for enc in ['utf-8-sig', 'utf-8', 'cp1256']:
            try:
                file.seek(0)
                df = pd.read_csv(file, encoding=enc)
                break
            except Exception:
                continue
    else:
        file.seek(0)
        df = pd.read_excel(file, header=1)
        # Validate: first column should be 'No.' or a known name
        first_col = str(df.columns[0])
        if first_col not in ('No.',) and not any(k in first_col for k in ['اسم', 'No']):
            file.seek(0)
            df = pd.read_excel(file, header=0)
    return df


def extract_json(text: str) -> dict:
    """Robustly extract JSON from Gemini response."""
    text = text.strip()
    text = re.sub(r'^```(?:json)?\s*', '', text, flags=re.MULTILINE)
    text = re.sub(r'\s*```\s*$', '', text, flags=re.MULTILINE)
    text = text.strip()
    start = text.find('{')
    end = text.rfind('}')
    if start == -1 or end == -1:
        raise json.JSONDecodeError("No JSON object found", text, 0)
    return json.loads(text[start:end + 1])


# ─── BATCHING CONFIG ─────────────────────────────────────────────────────────
BATCH_SIZE = 20            # منتجات لكل دفعة
BATCH_SLEEP_SEC = 4        # استراحة بين الدفعات لتفادي Rate Limit
MAX_RETRIES = 3            # عدد محاولات الإعادة لكل دفعة
BACKUP_FILE = "mahwous_backup.json"


def _save_backup(brand_results: dict) -> None:
    """احفظ نسخة احتياطية محلية بعد إكمال كل ماركة."""
    try:
        with open(BACKUP_FILE, 'w', encoding='utf-8') as f:
            json.dump(brand_results, f, ensure_ascii=False, indent=2)
    except Exception:
        pass  # لا نوقف المعالجة بسبب فشل الحفظ


def _load_backup() -> dict:
    try:
        if os.path.exists(BACKUP_FILE):
            with open(BACKUP_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def _build_batch_prompt(brand_name: str, products: list, use_grounding: bool,
                        batch_idx: int, total_batches: int) -> str:
    products_summary = json.dumps(
        [{'id': p.get('id', ''), 'name': p.get('name', ''),
          'price': p.get('price', 0),
          'desc_snippet': str(p.get('description', ''))[:150]}
         for p in products],
        ensure_ascii=False, indent=2
    )
    # المهمتان 2 و 3 (التساتر والنواقص) تُطلب فقط في الدفعة الأخيرة
    # لتجنب التكرار، أما المهمة 1 فتُطلب لكل دفعة.
    is_last = (batch_idx == total_batches - 1)
    extra_tasks = ""
    if is_last:
        extra_tasks = f"""
### المهمة 2: التساتر (للماركة كاملة)
{f'ابحث الآن في الإنترنت عن عطور ماركة "{brand_name}" المتاحة كـ "تستر" في السوق السعودي.' if use_grounding else f'بناءً على معرفتك، أي عطور ماركة "{brand_name}" يتوفر منها تستر في الأسواق؟'}
- قاعدة التسعير: خصم 70 ريال للمنتجات تحت 1000 ريال، خصم 150 ريال للمنتجات فوق 1000 ريال.

### المهمة 3: النواقص (للماركة كاملة)
{f'ابحث في الإنترنت عن عطور ماركة "{brand_name}" غير موجودة في قائمتي.' if use_grounding else f'بناءً على معرفتك، ما هي عطور ماركة "{brand_name}" الرئيسية غير الموجودة في قائمتي؟'}
لكل منتج ناقص، اكتب بياناته الكاملة مع وصف HTML محسّن جاهز للرفع على سلة.
"""

    return f"""أنت تعالج منتجات ماركة "{brand_name}" في متجر مهووس.

هذه الدفعة رقم {batch_idx + 1} من أصل {total_batches}.
عدد المنتجات في هذه الدفعة: {len(products)}

قائمة المنتجات (ملخص):
{products_summary}

**المطلوب:**

### المهمة 1: تحديث الأوصاف والسيو (لمنتجات هذه الدفعة فقط)
لكل منتج حالي، اكتب وصفاً HTML احترافياً محسّناً للـ SEO كامل (1200-1500 كلمة) وفق هويتك كخبير مهووس. الوصف يجب أن يكون بـ HTML: <h2>, <h3>, <ul>, <li>, <strong>, <p>.
{extra_tasks}
**أعد JSON صارم فقط:**

{{
  "brand": "{brand_name}",
  "products_updated": [
    {{
      "product_id": "string",
      "name": "اسم المنتج الكامل",
      "new_description": "<h2>...</h2><p>...</p>",
      "seo_title": "عنوان SEO أقل من 60 حرف",
      "seo_description": "وصف ميتا أقل من 155 حرف"
    }}
  ],
  "testers_updated": [
    {{
      "product_id": "string أو null",
      "name": "اسم المنتج كامل مع كلمة تستر",
      "is_new": true,
      "tester_available_in_market": true,
      "original_price": 0,
      "new_price": 0,
      "new_description": "<p>HTML وصف...</p>",
      "notes": "ملاحظة"
    }}
  ],
  "missing_products": [
    {{
      "name": "اسم المنتج الكامل",
      "type": "عطر مفرد",
      "category": "العطور > عطور رجالية",
      "price": 0,
      "description": "<h2>...</h2><p>...</p>",
      "brand": "{brand_name}",
      "is_tester": false
    }}
  ]
}}

ملاحظة: في الدفعات غير الأخيرة، أعد testers_updated و missing_products كقوائم فارغة [].
"""


def _call_gemini_single_batch(
    brand_name: str,
    products: list,
    api_key: str,
    model_name: str,
    use_grounding: bool,
    batch_idx: int,
    total_batches: int,
) -> dict:
    """نداء واحد للـ API لدفعة واحدة، مع response_mime_type=application/json."""
    genai.configure(api_key=api_key)

    model_kwargs = dict(
        model_name=model_name,
        system_instruction=SYSTEM_INSTRUCTION,
    )
    if use_grounding:
        model_kwargs['tools'] = [genai.protos.Tool(
            google_search_retrieval=genai.protos.GoogleSearchRetrieval()
        )]

    model = genai.GenerativeModel(**model_kwargs)

    prompt = _build_batch_prompt(brand_name, products, use_grounding,
                                 batch_idx, total_batches)

    # Structured Outputs: إجبار JSON نقي
    # ملاحظة: response_mime_type لا يتوافق مع google_search_retrieval في بعض الإصدارات،
    # لذلك نُفعّله فقط حين لا يوجد grounding.
    gen_config_kwargs = {'temperature': 0.0}
    if not use_grounding:
        gen_config_kwargs['response_mime_type'] = 'application/json'

    response = model.generate_content(
        prompt,
        generation_config=genai.GenerationConfig(**gen_config_kwargs),
    )
    return extract_json(response.text)


def _call_batch_with_retry(*args, **kwargs) -> dict:
    """إعادة المحاولة مع Exponential Backoff."""
    last_err = None
    for attempt in range(MAX_RETRIES):
        try:
            return _call_gemini_single_batch(*args, **kwargs)
        except Exception as e:
            last_err = e
            if attempt < MAX_RETRIES - 1:
                wait = (2 ** attempt) * 5  # 5s, 10s, 20s
                time.sleep(wait)
    raise last_err if last_err else RuntimeError("فشل غير معروف في الدفعة")


def call_gemini_brand(
    brand_name: str,
    products: list,
    api_key: str,
    model_name: str = 'gemini-2.0-flash',
    use_grounding: bool = True,
    progress_cb=None,
) -> dict:
    """معالجة ماركة كاملة عبر دفعات صغيرة، مع دمج النتائج وحفظ احتياطي."""
    if not products:
        return {
            'brand': brand_name,
            'products_updated': [],
            'testers_updated': [],
            'missing_products': [],
        }

    batches = [products[i:i + BATCH_SIZE]
               for i in range(0, len(products), BATCH_SIZE)]
    total_batches = len(batches)

    merged = {
        'brand': brand_name,
        'products_updated': [],
        'testers_updated': [],
        'missing_products': [],
    }
    failed_batches = []

    for b_idx, batch in enumerate(batches):
        if progress_cb:
            progress_cb(b_idx, total_batches, 'start', None)
        try:
            part = _call_batch_with_retry(
                brand_name, batch, api_key, model_name,
                use_grounding, b_idx, total_batches,
            )
            merged['products_updated'].extend(part.get('products_updated', []) or [])
            merged['testers_updated'].extend(part.get('testers_updated', []) or [])
            merged['missing_products'].extend(part.get('missing_products', []) or [])
            if progress_cb:
                progress_cb(b_idx, total_batches, 'done', part)
        except Exception as e:
            failed_batches.append({'batch_idx': b_idx, 'error': str(e),
                                   'product_ids': [p.get('id') for p in batch]})
            if progress_cb:
                progress_cb(b_idx, total_batches, 'failed', str(e))

        # استراحة بين الدفعات (ليس بعد الأخيرة)
        if b_idx < total_batches - 1:
            time.sleep(BATCH_SLEEP_SEC)

    if failed_batches:
        merged['_failed_batches'] = failed_batches

    return merged


def build_output_excel(result: dict, original_df: pd.DataFrame, template_bytes: bytes) -> bytes:
    """Build Salla-compatible Excel from AI results."""
    brand_col = get_brand_col(original_df)
    name_col  = find_col(original_df, 'name')
    price_col = find_col(original_df, 'price')
    desc_col  = find_col(original_df, 'description')
    cat_col   = find_col(original_df, 'category')
    qty_col   = find_col(original_df, 'quantity')

    brand_name = result.get('brand', '')

    # Filter brand rows
    if brand_col:
        mask = original_df[brand_col].astype(str).str.contains(
            re.escape(brand_name), case=False, na=False
        )
        brand_df = original_df[mask].copy()
    else:
        brand_df = original_df.copy()

    updated_map = {
        str(p['product_id']): p
        for p in result.get('products_updated', [])
        if p.get('product_id')
    }

    rows = []

    # Existing products with updated descriptions
    for _, row in brand_df.iterrows():
        pid = str(row.get('No.', row.name))
        new_row = row.copy()
        if pid in updated_map and desc_col:
            new_row[desc_col] = updated_map[pid].get('new_description', row.get(desc_col, ''))
        rows.append(new_row)

    all_cols = list(original_df.columns)

    # New tester products
    for tester in result.get('testers_updated', []):
        if tester.get('is_new'):
            nr = {c: '' for c in all_cols}
            if name_col:  nr[name_col] = tester.get('name', '')
            if price_col: nr[price_col] = tester.get('new_price', 0)
            if desc_col:  nr[desc_col] = tester.get('new_description', '')
            if brand_col: nr[brand_col] = brand_name
            if cat_col:   nr[cat_col] = 'العطور > عطور التساتر'
            if qty_col:   nr[qty_col] = 10
            rows.append(pd.Series(nr))

    # Missing products
    for missing in result.get('missing_products', []):
        nr = {c: '' for c in all_cols}
        if name_col:  nr[name_col] = missing.get('name', '')
        if price_col: nr[price_col] = missing.get('price', 0)
        if desc_col:  nr[desc_col] = missing.get('description', '')
        if brand_col: nr[brand_col] = missing.get('brand', brand_name)
        if cat_col:   nr[cat_col] = missing.get('category', '')
        if qty_col:   nr[qty_col] = 10
        rows.append(pd.Series(nr))

    output_df = pd.DataFrame(rows) if rows else brand_df.copy()

    # Load template and write
    wb = load_workbook(io.BytesIO(template_bytes))
    ws = wb.active

    # Find header row (row with 'اسم' or 'No.' cells)
    header_row = 2
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=6, values_only=True), 1):
        if any(cell and ('اسم' in str(cell) or cell == 'No.') for cell in row):
            header_row = i
            break
    data_start = header_row + 1

    template_headers = [
        ws.cell(row=header_row, column=c).value
        for c in range(1, ws.max_column + 1)
    ]

    # Map template columns to output_df columns
    col_map = {}
    for t_idx, t_hdr in enumerate(template_headers):
        if not t_hdr:
            continue
        t_str = str(t_hdr)
        for df_col in output_df.columns:
            if t_str in str(df_col) or str(df_col) in t_str:
                col_map[t_idx + 1] = df_col
                break

    # Write rows
    for r_idx, (_, row) in enumerate(output_df.iterrows()):
        excel_row = data_start + r_idx
        for t_col, df_col in col_map.items():
            val = row.get(df_col, '')
            if pd.isna(val) if not isinstance(val, str) else False:
                val = ''
            ws.cell(row=excel_row, column=t_col, value=val)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ─── SESSION STATE INIT ───────────────────────────────────────────────────────
def init_state():
    defaults = {
        'df': None,
        'brand_col': None,
        'brands_list': [],
        'filtered_brands': [],
        'current_brand_idx': 0,
        'brand_results': {},
        'processing': False,
        'waiting_confirm': False,
        'current_result': None,
        'template_bytes': None,
        'api_key': os.environ.get('GEMINI_API_KEY', ''),
        'model_name': 'gemini-2.0-flash',
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


init_state()

# ─── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ الإعدادات")

    api_key = st.text_input(
        "🔑 Gemini API Key",
        value=st.session_state.api_key,
        type="password",
        placeholder="AIza...",
        help="احصل على مفتاحك من Google AI Studio",
    )
    st.session_state.api_key = api_key

    model_name = st.selectbox(
        "🤖 النموذج",
        ['gemini-2.0-flash', 'gemini-1.5-pro', 'gemini-1.5-flash'],
        index=0,
        help="gemini-2.0-flash: أسرع وأرخص | gemini-1.5-pro: أدق وأشمل",
    )
    st.session_state.model_name = model_name

    st.divider()
    st.markdown("## 📁 رفع الملفات")

    uploaded_file = st.file_uploader(
        "ملف المنتجات (Excel / CSV)",
        type=['xlsx', 'xls', 'csv'],
        key='products_file',
    )

    template_file = st.file_uploader(
        "قالب سلة الفارغ (اختياري للتصدير)",
        type=['xlsx'],
        key='template_file',
    )

    if template_file:
        st.session_state.template_bytes = template_file.read()
        st.success("✅ تم تحميل القالب")

    # ─── استرجاع نسخة احتياطية ───────────────────────────────────────────
    if os.path.exists(BACKUP_FILE):
        st.divider()
        st.markdown("## 💾 النسخة الاحتياطية")
        backup_data = _load_backup()
        st.caption(f"📦 {len(backup_data)} ماركة محفوظة محلياً")
        col_r1, col_r2 = st.columns(2)
        with col_r1:
            if st.button("♻️ استرجاع", use_container_width=True):
                st.session_state.brand_results = backup_data
                st.success(f"✅ تم استرجاع {len(backup_data)} ماركة")
                st.rerun()
        with col_r2:
            try:
                with open(BACKUP_FILE, 'rb') as _bf:
                    st.download_button(
                        "⬇️ تحميل",
                        data=_bf.read(),
                        file_name=BACKUP_FILE,
                        mime="application/json",
                        use_container_width=True,
                    )
            except Exception:
                pass

    if uploaded_file and st.button("📊 تحليل الملف", use_container_width=True, type="primary"):
        with st.spinner("جاري تحليل الملف..."):
            try:
                df = load_products(uploaded_file)
                st.session_state.df = df
                brand_col = get_brand_col(df)
                st.session_state.brand_col = brand_col
                if brand_col:
                    brands = sorted(df[brand_col].dropna().unique().tolist())
                    st.session_state.brands_list = brands
                    st.session_state.filtered_brands = brands
                    st.session_state.current_brand_idx = 0
                    st.session_state.brand_results = {}
                    st.success(f"✅ {len(df):,} منتج | {len(brands):,} ماركة")
                else:
                    st.error("❌ لم يُعثر على عمود الماركة")
            except Exception as e:
                st.error(f"❌ خطأ في قراءة الملف: {e}")

    # ─── FILTERS ─────────────────────────────────────────────────────────────
    if st.session_state.brands_list:
        st.divider()
        st.markdown("## 🔍 الفلاتر المتقدمة")

        selected_brands = st.multiselect(
            "الماركات",
            options=st.session_state.brands_list,
            default=[],
            placeholder="كل الماركات",
        )

        df_now = st.session_state.df
        cat_col = find_col(df_now, 'category') if df_now is not None else None

        selected_cats = []
        if cat_col is not None and df_now is not None:
            all_cats = sorted(df_now[cat_col].dropna().astype(str).unique().tolist())
            selected_cats = st.multiselect(
                "الفئات",
                options=all_cats,
                default=[],
                placeholder="كل الفئات",
            )

        price_range = None
        price_col_now = find_col(df_now, 'price') if df_now is not None else None
        if price_col_now is not None and df_now is not None:
            prices = pd.to_numeric(df_now[price_col_now], errors='coerce').dropna()
            if len(prices) > 0:
                min_p, max_p = int(prices.min()), int(prices.max())
                price_range = st.slider(
                    "نطاق السعر (ريال)", min_p, max_p, (min_p, max_p), step=10
                )

        only_with_testers = st.checkbox("ماركات بها تساتر فقط", value=False)
        sort_by = st.selectbox(
            "ترتيب الماركات حسب",
            ['الاسم أبجدياً', 'عدد المنتجات (تنازلي)', 'عدد المنتجات (تصاعدي)']
        )

        if st.button("✅ تطبيق الفلاتر", use_container_width=True):
            df_f = st.session_state.df
            brand_col_f = st.session_state.brand_col
            filtered = st.session_state.brands_list.copy()

            if selected_brands:
                filtered = [b for b in filtered if b in selected_brands]

            if df_f is not None and selected_cats and cat_col:
                cat_df = df_f[df_f[cat_col].astype(str).isin(selected_cats)]
                cat_brands = set(cat_df[brand_col_f].dropna().astype(str).tolist())
                filtered = [b for b in filtered if b in cat_brands]

            if df_f is not None and price_range and price_col_now:
                p_df = df_f[
                    pd.to_numeric(df_f[price_col_now], errors='coerce').between(
                        price_range[0], price_range[1]
                    )
                ]
                p_brands = set(p_df[brand_col_f].dropna().astype(str).tolist())
                filtered = [b for b in filtered if b in p_brands]

            name_col_f = find_col(df_f, 'name')
            if only_with_testers and df_f is not None and name_col_f:
                tester_brands = set()
                for brand in filtered:
                    bd = df_f[df_f[brand_col_f].astype(str).str.contains(
                        re.escape(str(brand)), case=False, na=False
                    )]
                    if bd[name_col_f].apply(is_tester).any():
                        tester_brands.add(brand)
                filtered = [b for b in filtered if b in tester_brands]

            # Sort
            if sort_by == 'عدد المنتجات (تنازلي)' and df_f is not None:
                filtered.sort(
                    key=lambda b: len(df_f[df_f[brand_col_f].astype(str).str.contains(
                        re.escape(str(b)), case=False, na=False
                    )]),
                    reverse=True
                )
            elif sort_by == 'عدد المنتجات (تصاعدي)' and df_f is not None:
                filtered.sort(
                    key=lambda b: len(df_f[df_f[brand_col_f].astype(str).str.contains(
                        re.escape(str(b)), case=False, na=False
                    )])
                )
            else:
                filtered.sort()

            st.session_state.filtered_brands = filtered
            st.session_state.current_brand_idx = 0
            st.session_state.brand_results = {}
            st.session_state.processing = False
            st.session_state.waiting_confirm = False
            st.success(f"✅ {len(filtered)} ماركة بعد التصفية")
            st.rerun()

# ─── MAIN AREA ────────────────────────────────────────────────────────────────
st.markdown("# 🌿 مهووس | معالج المنتجات الذكي")
st.caption("أتمتة تحديث الأوصاف · اكتشاف التساتر · سد الفجوات — ماركة بماركة")

if st.session_state.df is None:
    st.info("👆 ارفع ملف المنتجات من الشريط الجانبي للبدء")
    st.stop()

df = st.session_state.df
filtered_brands = st.session_state.filtered_brands
brand_col = st.session_state.brand_col
current_idx = st.session_state.current_brand_idx
total_brands = len(filtered_brands)

if total_brands == 0:
    st.warning("⚠️ لا توجد ماركات بعد تطبيق الفلاتر — عدّل الفلاتر من الشريط الجانبي")
    st.stop()

# ─── OVERALL PROGRESS BAR ────────────────────────────────────────────────────
completed = len(st.session_state.brand_results)
col_p1, col_p2, col_p3, col_p4 = st.columns([3, 1, 1, 1])
with col_p1:
    st.markdown(f"### التقدم الإجمالي")
    overall_pct = completed / total_brands if total_brands > 0 else 0
    st.progress(overall_pct, text=f"{completed}/{total_brands} ماركة مكتملة ({overall_pct*100:.0f}%)")
with col_p2:
    st.metric("مكتمل", completed)
with col_p3:
    st.metric("متبقي", total_brands - completed)
with col_p4:
    st.metric("إجمالي", total_brands)

# ─── BRANDS OVERVIEW TABLE ───────────────────────────────────────────────────
with st.expander(f"📋 قائمة الماركات المختارة ({total_brands})", expanded=False):
    name_col = find_col(df, 'name')
    price_col = find_col(df, 'price')
    summary_rows = []
    for i, brand in enumerate(filtered_brands):
        bd = df[df[brand_col].astype(str).str.contains(
            re.escape(str(brand)), case=False, na=False
        )]
        tester_cnt = bd[name_col].apply(is_tester).sum() if name_col else 0
        avg_price = round(
            pd.to_numeric(bd[price_col], errors='coerce').mean(), 0
        ) if price_col else 0
        status_icon = (
            "✅" if brand in st.session_state.brand_results else
            ("🔄" if i == current_idx and st.session_state.processing else
             ("⏸️" if i > current_idx else "⏭️"))
        )
        summary_rows.append({
            '#': i + 1,
            'الماركة': brand,
            'المنتجات': len(bd),
            'التساتر': tester_cnt,
            'متوسط السعر': avg_price,
            'الحالة': status_icon,
        })
    st.dataframe(
        pd.DataFrame(summary_rows),
        use_container_width=True,
        hide_index=True,
        column_config={
            'متوسط السعر': st.column_config.NumberColumn(format="%.0f ريال"),
        }
    )

st.divider()

# ─── ALL DONE ────────────────────────────────────────────────────────────────
if current_idx >= total_brands:
    st.balloons()
    st.success("🎉 تمت معالجة جميع الماركات بنجاح!")
    results_json = json.dumps(
        st.session_state.brand_results, ensure_ascii=False, indent=2
    )
    st.download_button(
        "⬇️ تحميل جميع النتائج (JSON)",
        data=results_json.encode('utf-8'),
        file_name=f"mahwous_all_results_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
        mime="application/json",
        use_container_width=True,
        type="primary",
    )
    if st.button("🔄 بدء جلسة جديدة", use_container_width=True):
        for k in ['brand_results', 'current_brand_idx', 'processing',
                  'waiting_confirm', 'current_result']:
            if k in st.session_state:
                del st.session_state[k]
        st.rerun()
    st.stop()

# ─── CURRENT BRAND ───────────────────────────────────────────────────────────
current_brand = filtered_brands[current_idx]
brand_df = df[df[brand_col].astype(str).str.contains(
    re.escape(str(current_brand)), case=False, na=False
)]
name_col  = find_col(df, 'name')
price_col = find_col(df, 'price')
desc_col  = find_col(df, 'description')
tester_count = brand_df[name_col].apply(is_tester).sum() if name_col else 0

avg_price_val = 0
if price_col:
    prices_series = pd.to_numeric(brand_df[price_col], errors='coerce').dropna()
    avg_price_val = int(prices_series.mean()) if len(prices_series) > 0 else 0

# Brand dashboard
st.markdown(f"""
<div class="brand-card">
  <h2 style="margin:0 0 16px 0;">📦 {current_brand}</h2>
  <div style="display:flex; gap:16px; flex-wrap:wrap;">
    <div class="stat-box" style="flex:1; min-width:120px;">
      <div style="font-size:2em; font-weight:bold;">{len(brand_df):,}</div>
      <div style="color:#aaa; font-size:.9em;">إجمالي المنتجات</div>
    </div>
    <div class="stat-box" style="flex:1; min-width:120px;">
      <div style="font-size:2em; font-weight:bold;">{tester_count:,}</div>
      <div style="color:#aaa; font-size:.9em;">التساتر الحالية</div>
    </div>
    <div class="stat-box" style="flex:1; min-width:120px;">
      <div style="font-size:2em; font-weight:bold;">{avg_price_val:,}</div>
      <div style="color:#aaa; font-size:.9em;">متوسط السعر (ريال)</div>
    </div>
    <div class="stat-box" style="flex:1; min-width:120px;">
      <div style="font-size:2em; font-weight:bold;">{current_idx + 1}/{total_brands}</div>
      <div style="color:#aaa; font-size:.9em;">ترتيب الماركة</div>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

# ─── WAITING FOR CONFIRMATION (results ready) ────────────────────────────────
if st.session_state.waiting_confirm and st.session_state.current_result:
    result = st.session_state.current_result
    n_upd = len(result.get('products_updated', []))
    n_tst = len(result.get('testers_updated', []))
    n_mis = len(result.get('missing_products', []))

    st.success(f"✅ اكتملت معالجة **{current_brand}** | {n_upd} وصف محدّث · {n_tst} تستر · {n_mis} منتج ناقص")

    tabs = st.tabs([
        f"📝 الأوصاف المحدّثة ({n_upd})",
        f"🏷️ التساتر ({n_tst})",
        f"🔍 المنتجات الناقصة ({n_mis})",
    ])

    with tabs[0]:
        updated = result.get('products_updated', [])
        if updated:
            disp_cols = [c for c in ['product_id', 'name', 'seo_title', 'seo_description']
                         if c in pd.DataFrame(updated).columns]
            st.dataframe(
                pd.DataFrame(updated)[disp_cols] if disp_cols else pd.DataFrame(updated),
                use_container_width=True, hide_index=True
            )
            with st.expander("👁️ معاينة أول وصف HTML"):
                if updated[0].get('new_description'):
                    st.markdown(updated[0]['new_description'], unsafe_allow_html=True)
        else:
            st.info("لا توجد تحديثات للأوصاف")

    with tabs[1]:
        testers = result.get('testers_updated', [])
        if testers:
            disp_cols = [c for c in ['name', 'is_new', 'original_price', 'new_price', 'tester_available_in_market', 'notes']
                         if c in pd.DataFrame(testers).columns]
            st.dataframe(
                pd.DataFrame(testers)[disp_cols] if disp_cols else pd.DataFrame(testers),
                use_container_width=True, hide_index=True
            )
        else:
            st.info("لم يتم العثور على تساتر جديدة")

    with tabs[2]:
        missing = result.get('missing_products', [])
        if missing:
            disp_cols = [c for c in ['name', 'type', 'category', 'price', 'is_tester']
                         if c in pd.DataFrame(missing).columns]
            st.dataframe(
                pd.DataFrame(missing)[disp_cols] if disp_cols else pd.DataFrame(missing),
                use_container_width=True, hide_index=True
            )
        else:
            st.info("لا توجد منتجات ناقصة مقترحة")

    # Download buttons
    st.markdown("### تحميل النتائج")
    dl_col1, dl_col2, dl_col3 = st.columns(3)

    with dl_col1:
        safe_brand = re.sub(r'[^\w]', '_', current_brand)
        st.download_button(
            f"JSON — {current_brand}",
            data=json.dumps(result, ensure_ascii=False, indent=2).encode('utf-8'),
            file_name=f"mahwous_{safe_brand}_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
            mime="application/json",
            use_container_width=True,
        )

    with dl_col2:
        if st.session_state.template_bytes:
            try:
                excel_out = build_output_excel(result, df, st.session_state.template_bytes)
                safe_b = re.sub(r'[^\w]', '_', current_brand)
                st.download_button(
                    f"Excel سلة — {current_brand}",
                    data=excel_out,
                    file_name=f"salla_{safe_b}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as e:
                st.caption(f"⚠️ Excel: {e}")
        else:
            st.caption("🔺 ارفع قالب سلة للحصول على Excel")

    with dl_col3:
        if st.button("⏭️ تخطي هذه الماركة", use_container_width=True, type="secondary"):
            st.session_state.current_brand_idx += 1
            st.session_state.waiting_confirm = False
            st.session_state.current_result = None
            st.session_state.processing = False
            st.rerun()

    st.markdown("---")
    if st.button(
        f"✅ تأكيد واستخراج — ثم الانتقال لـ {filtered_brands[current_idx + 1] if current_idx + 1 < total_brands else 'النهاية'}",
        type="primary",
        use_container_width=True,
    ):
        st.session_state.brand_results[current_brand] = result
        _save_backup(st.session_state.brand_results)  # حفظ احتياطي تلقائي
        st.session_state.current_brand_idx += 1
        st.session_state.waiting_confirm = False
        st.session_state.current_result = None
        st.session_state.processing = False
        st.rerun()

    st.stop()

# ─── START BUTTON (idle, not processing) ─────────────────────────────────────
if not st.session_state.processing:
    if not st.session_state.api_key:
        st.warning("⚠️ أدخل Gemini API Key في الشريط الجانبي أولاً")
        st.stop()

    # Preview products
    with st.expander(f"👁️ معاينة منتجات {current_brand}", expanded=True):
        preview_rows = []
        if name_col:
            for _, row in brand_df.head(15).iterrows():
                preview_rows.append({
                    'اسم المنتج': str(row.get(name_col, '')),
                    'السعر': row.get(price_col, '') if price_col else '',
                    'تستر؟': '✅' if is_tester(str(row.get(name_col, ''))) else '',
                })
        if preview_rows:
            st.dataframe(pd.DataFrame(preview_rows), use_container_width=True, hide_index=True)
        if len(brand_df) > 15:
            st.caption(f"... و {len(brand_df) - 15} منتج آخر")

    col_start, col_skip = st.columns([4, 1])
    with col_start:
        if st.button(
            f"🚀 بدء معالجة {current_brand} ({len(brand_df)} منتج) بالذكاء الاصطناعي",
            type="primary",
            use_container_width=True,
        ):
            st.session_state.processing = True
            st.rerun()
    with col_skip:
        if st.button("⏭️ تخطي", use_container_width=True, type="secondary"):
            st.session_state.current_brand_idx += 1
            st.rerun()

    st.stop()

# ─── ACTIVE PROCESSING ───────────────────────────────────────────────────────
st.markdown(f"""
<div class="section-header">
🔄 جاري معالجة ماركة: {current_brand} — {len(brand_df)} منتج
</div>
""", unsafe_allow_html=True)

# Dual progress bars
brand_lbl  = st.empty()
brand_bar  = st.progress(0)
prod_lbl   = st.empty()
prod_bar   = st.progress(0)
status_msg = st.empty()

brand_lbl.markdown("**الخطوة 1/3:** جاري تجهيز بيانات المنتجات...")
brand_bar.progress(5)

# Build products payload — تنظيف NaN قبل التحويل لـ JSON
brand_df_clean = brand_df.fillna("")
products_payload = []
if name_col:
    for _, row in brand_df_clean.iterrows():
        try:
            price_val = float(pd.to_numeric(row.get(price_col, 0), errors='coerce') or 0)
        except Exception:
            price_val = 0.0
        products_payload.append({
            'id': str(row.get('No.', row.name)),
            'name': str(row.get(name_col, '')),
            'price': price_val,
            'description': str(row.get(desc_col, ''))[:300] if desc_col else '',
        })

n = len(products_payload)
total_batches_est = max(1, (n + BATCH_SIZE - 1) // BATCH_SIZE)

brand_bar.progress(15)
brand_lbl.markdown(
    f"**الخطوة 2/3:** إرسال {n} منتج على {total_batches_est} دفعة "
    f"(حجم الدفعة={BATCH_SIZE}) إلى Gemini AI..."
)
prod_bar.progress(0.0)
prod_lbl.markdown("⏳ بدء معالجة الدفعات...")
status_msg.info(
    f"🤖 يعمل الذكاء الاصطناعي على:\n"
    f"- تحديث {n} وصف بكلمات مفتاحية SEO (دفعات من {BATCH_SIZE})\n"
    f"- البحث عن التساتر المتاحة لـ {current_brand} في السوق السعودي\n"
    f"- اكتشاف المنتجات الناقصة\n"
    f"- استراحة {BATCH_SLEEP_SEC} ثانية بين الدفعات لتفادي Rate Limit"
)


def _ui_progress(b_idx, total, phase, payload):
    frac = b_idx / max(total, 1)
    if phase == 'start':
        prod_bar.progress(min(frac + 0.01, 1.0))
        prod_lbl.markdown(
            f"📦 الدفعة {b_idx + 1}/{total} — جاري الإرسال إلى Gemini..."
        )
    elif phase == 'done':
        done_frac = (b_idx + 1) / max(total, 1)
        prod_bar.progress(min(done_frac, 1.0))
        n_p = len(payload.get('products_updated', [])) if payload else 0
        prod_lbl.markdown(
            f"✅ الدفعة {b_idx + 1}/{total} اكتملت — {n_p} وصف في هذه الدفعة"
        )
        brand_bar.progress(min(15 + int(done_frac * 70), 95))
    elif phase == 'failed':
        prod_lbl.markdown(
            f"⚠️ الدفعة {b_idx + 1}/{total} فشلت بعد {MAX_RETRIES} محاولات — تم تجاوزها"
        )


try:
    result = call_gemini_brand(
        brand_name=current_brand,
        products=products_payload,
        api_key=st.session_state.api_key,
        model_name=st.session_state.model_name,
        use_grounding=True,
        progress_cb=_ui_progress,
    )

    brand_bar.progress(75)
    prod_bar.progress(0.7)
    brand_lbl.markdown("**الخطوة 3/3:** معالجة النتائج...")

    n_upd = len(result.get('products_updated', []))
    n_tst = len(result.get('testers_updated', []))
    n_mis = len(result.get('missing_products', []))

    # Animate product completion
    for i in range(n_upd):
        frac = 0.7 + (i + 1) / max(n_upd, 1) * 0.3
        prod_bar.progress(min(frac, 1.0))
        prod_lbl.markdown(f"✅ معالجة الوصف: {i + 1}/{n_upd}")
        time.sleep(0.04)

    brand_bar.progress(100)
    prod_bar.progress(1.0)
    brand_lbl.markdown(f"✅ **اكتملت معالجة {current_brand}!**")
    prod_lbl.markdown(f"✅ {n_upd} وصف · {n_tst} تستر · {n_mis} ناقص")
    status_msg.success(f"🎯 اكتملت المعالجة — {n_upd} وصف محدّث | {n_tst} تستر | {n_mis} منتج ناقص")

    st.session_state.current_result = result
    st.session_state.waiting_confirm = True
    st.session_state.processing = False
    time.sleep(0.5)
    st.rerun()

except Exception as e:
    err = str(e)
    brand_bar.progress(0)

    # Try without grounding if it's a grounding error
    if any(x in err.lower() for x in ['grounding', 'search', 'tool', 'billing']):
        status_msg.warning("⚠️ Google Search غير متاح — جاري إعادة المحاولة بدون بحث مباشر...")
        prod_lbl.markdown("🔄 جاري الإعادة...")
        prod_bar.progress(0.1)
        try:
            result = call_gemini_brand(
                brand_name=current_brand,
                products=products_payload,
                api_key=st.session_state.api_key,
                model_name=st.session_state.model_name,
                use_grounding=False,
                progress_cb=_ui_progress,
            )
            brand_bar.progress(100)
            prod_bar.progress(1.0)
            status_msg.success("✅ اكتملت المعالجة (بدون Google Search)")
            st.session_state.current_result = result
            st.session_state.waiting_confirm = True
            st.session_state.processing = False
            time.sleep(0.5)
            st.rerun()
        except Exception as e2:
            status_msg.error(f"❌ فشل أيضاً بدون grounding: {e2}")
            st.session_state.processing = False

    elif 'api_key' in err.lower() or 'api key' in err.lower() or 'invalid' in err.lower():
        status_msg.error("❌ Gemini API Key غير صحيح — تحقق من المفتاح في الشريط الجانبي")
        st.session_state.processing = False

    elif 'quota' in err.lower() or 'rate' in err.lower() or '429' in err:
        status_msg.error("❌ تجاوز حد الاستخدام (Rate Limit) — انتظر دقيقة وأعد المحاولة")
        st.session_state.processing = False

    else:
        status_msg.error(f"❌ خطأ: {err}")
        st.session_state.processing = False

    col_retry, col_skip2 = st.columns(2)
    with col_retry:
        if st.button("🔄 إعادة المحاولة", use_container_width=True, type="primary"):
            st.session_state.processing = True
            st.rerun()
    with col_skip2:
        if st.button("⏭️ تخطي هذه الماركة", use_container_width=True):
            st.session_state.current_brand_idx += 1
            st.session_state.processing = False
            st.rerun()
