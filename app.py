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

# ─── COMPETITOR STORES (Saudi market — strict scope) ─────────────────────────
COMPETITOR_STORES = [
    "https://saeedsalah.com/", "https://vanilla.sa/", "https://sara-makeup.com/",
    "https://alkhabeershop.com/", "https://www.goldenscent.com/", "https://leesanto.com/",
    "https://azalperfume.com/", "https://candyniche.com/", "https://luxuryperfumesnish.com/",
    "https://hanan-store55.com/", "https://areejamwaj.com/", "https://niceonesa.com/",
    "https://www.sephora.me/sa-ar", "https://www.faces.sa/ar", "https://niche.sa/",
    "https://worldgivenchy.com/", "https://sarahmakeup37.com/", "https://aromaticcloud.com/",
    "https://tatayab.com/", "https://kayan9.com/",
]

# ─── HTML TEMPLATES (mandatory output structures) ────────────────────────────
HTML_TEMPLATE_NEW = """<p><strong>[مقدمة تسويقية جذابة]</strong></p>
<h3>التفاصيل والخطوط العطرية باختصار</h3>
<ul>
  <li><strong>الماركة:</strong> [اسم الماركة]</li>
  <li><strong>اسم العطر:</strong> [اسم العطر]</li>
  <li><strong>الجنس:</strong> [رجالي / نسائي / للجنسين]</li>
  <li><strong>الخط العطري (العائلة):</strong> [مثل: حمضي - أروماتك]</li>
  <li><strong>الحجم:</strong> [الحجم مل]</li>
  <li><strong>التركيز:</strong> [مثل: أودي بارفيوم]</li>
  <li><strong>سنة الإصدار:</strong> [السنة والتوقيع إن وجد]</li>
</ul>
<h3>رحلة العطر (النوتات)</h3>
<ul>
  <li><strong>الافتتاحية:</strong> [وصف]</li>
  <li><strong>القلب العطري:</strong> [وصف]</li>
  <li><strong>القاعدة الأساسية:</strong> [وصف]</li>
</ul>
<h3>لماذا تختار هذا العطر؟</h3>
<ul>
  <li><strong>رائحة متوازنة:</strong> [وصف]</li>
  <li><strong>مثالي لجميع الأوقات:</strong> [وصف]</li>
  <li><strong>أداء قوي:</strong> [وصف]</li>
</ul>
<h3>الأسئلة الشائعة</h3>
<p><strong>هل هذا العطر مناسب للمناخ الحار؟</strong><br>[إجابة]</p>
<p><strong>هل يمكن استخدامه يومياً؟</strong><br>[إجابة]</p>
<p><strong>ما هي المناسبة الأفضل لاستخدامه؟</strong><br>[إجابة]</p>
<h3>اكتشف المزيد من مهووس</h3>
<ul>
  <li><a href="#">استكشف أحدث العطور الرجالية هنا</a></li>
  <li><a href="#">تصفح أجمل العطور النسائية الجذابة</a></li>
  <li><a href="#">للباحثين عن التميز، استكشف عطور النيش الفاخرة</a></li>
</ul>"""

HTML_TEMPLATE_TESTER = """<p><strong>استمتع بالفخامة المطلقة بتكلفة أذكى! نقدم لك تستر "[اسم العطر]" من [الماركة] الأصلي 100%، ليمنحك نفس التجربة، الثبات، والفوحان للإصدار المغلف ولكن بسعر استثنائي. [مقدمة عن العطر].</strong></p>
<h3>التفاصيل والخطوط باختصار</h3>
<ul>
  <li><strong>الماركة:</strong> [اسم الماركة]</li>
  <li><strong>الاسم:</strong> [اسم العطر]</li>
  <li><strong>حالة المنتج:</strong> تستر (Tester) أصلي 100%.</li>
  <li><strong>الجنس:</strong> [رجالي / نسائي / للجنسين]</li>
  <li><strong>الخط (العائلة):</strong> [مثل: حمضي - أروماتك]</li>
  <li><strong>الحجم:</strong> [الحجم مل]</li>
  <li><strong>التركيز:</strong> [مثل: أودي بارفيوم]</li>
</ul>
<h3>رحلة النوتات</h3>
<ul>
  <li><strong>الافتتاحية:</strong> [وصف]</li>
  <li><strong>القلب:</strong> [وصف]</li>
  <li><strong>القاعدة الأساسية:</strong> [وصف]</li>
</ul>
<h3>لماذا تختار هذا الإصدار؟</h3>
<ul>
  <li><strong>رائحة متوازنة:</strong> [وصف]</li>
  <li><strong>مثالي لجميع الأوقات:</strong> [وصف]</li>
  <li><strong>أداء قوي:</strong> [وصف]</li>
</ul>
<h3>الدليل الشامل للتساتر من متجر مهووس</h3>
<p>هل تتساءل عن سر التساتر ولماذا تحظى بشعبية هائلة بين عشاق الروائح الفاخرة؟ يسعدنا في متجر مهووس أن نكشف لك هذا السر، لنجعل تجربة تسوقك أكثر ذكاءً وثقة.</p>
<p><strong>ما هو التستر؟</strong><br>التستر هو نسخة أصلية 100% تصدرها الشركة المصنعة (الماركات العالمية) جنباً إلى جنب مع المنتجات التجارية. الهدف الأساسي من إنتاجه هو وضعه في المتاجر والبوتيكات الفاخرة ليتمكن العملاء من تجربة الرائحة والأداء قبل الشراء.</p>
<p><strong>ما الفرق بين التستر والإصدار العادي المغلف؟</strong><br>الفرق الوحيد والأساسي يكمن في "الشكل الخارجي" فقط، ولا مساومة أبداً على الجودة:</p>
<ul>
  <li><strong>السائل:</strong> متطابق 100% من حيث المكونات، التركيز، الثبات، والفوحان. أنت تحصل على نفس القطرة الأصلية تماماً.</li>
  <li><strong>الزجاجة:</strong> يأتي في نفس الزجاجة الأصلية الفاخرة للماركة، وقد يُطبع عليها أحياناً عبارة (Tester) أو (Demonstration).</li>
  <li><strong>العلبة الخارجية:</strong> بهدف تقليل التكاليف، تُصدر الشركات التساتر في علب كرتونية بسيطة (غالباً بيضاء أو بنية صديقة للبيئة)، وتأتي بدون الغلاف البلاستيكي الشفاف (السلوفان).</li>
  <li><strong>الغطاء:</strong> تأتي معظم التساتر بغطائها الأصلي الفاخر، وفي حالات نادرة جداً قد تأتي بدون غطاء بناءً على تصميم الشركة المصنعة.</li>
</ul>
<p><strong>لماذا يعتبر التستر استثماراً ذكياً؟</strong><br>إذا كنت تشتري لاقتنائك الشخصي وليس لتقديمه كهدية رسمية، فإن التستر هو الخيار الأكثر ذكاءً وتوفيراً. فهو يتيح لك الاستمتاع بأرقى الروائح العالمية وإصدارات النيش بأسعار اقتصادية مخفضة جداً، لتحصل على أقصى قيمة مقابل ما تدفعه.</p>
<p><strong>ضمان مهووس الذهبي</strong><br>نحن في متجر مهووس نضع ثقتك في المقام الأول. نضمن لك أصالة جميع التساتر المتوفرة لدينا بنسبة 100%. يتم توفيرها من نفس الموزعين المعتمدين للماركات العالمية، لتعيش تجربة الفخامة المطلقة براحة بال تامة.</p>
<h3>اكتشف المزيد من مهووس</h3>
<ul>
  <li><a href="#">تصفح تشكيلتنا الواسعة من التساتر الأصلية</a></li>
  <li><a href="#">تسوق المزيد من إصدارات النيش الرجالية الفاخرة</a></li>
  <li><a href="#">اكتشف أحدث إصدارات النيش النسائية</a></li>
</ul>"""

# ─── SYSTEM INSTRUCTION TEMPLATE (حقن ديناميكي لبصمة الكتابة) ───────────────
# لا تستخدم f-string هنا — القيم تُحقن في call_gemini_brand
SYSTEM_INSTRUCTION_TEMPLATE = """## هويتك ومهمتك
أنت **خبير عطور محترف بخبرة 20 سنة** + محلل تنافسي لمتجر مهووس في السوق السعودي.
مهمتك الوحيدة: **اكتشاف الفجوات** — التساتر الناقصة والمنتجات الناقصة — وكتابة وصف **فقط للجديد** الذي تقترح إضافته.

## ❌ ممنوع منعاً باتاً
- لا تُعيد كتابة أو تُحدّث وصف أي منتج موجود مسبقاً في قائمتنا أبداً
- products_updated يجب أن يكون [] قائمة فارغة في كل رد
- لا تقترح منتجاً موجوداً في قائمتنا ولو بصيغة مختلفة

## قواعد التساتر (إلزامية)
**قاعدة 1 — فحص وجود التستر أولاً:**
لكل عطر أساسي في قائمتنا، تحقق: هل يوجد في "التساتر الموجودة" تستر يحمل نفس الاسم؟
- إذا نعم → تخطّ تماماً، لا تفعل شيئاً
- إذا لا → انتقل لقاعدة 2

**قاعدة 2 — البحث عند المنافسين:**
ابحث في المتاجر السعودية المحددة: هل يبيعون تستر لهذا العطر؟
- إذا نعم → أضفه في testers_to_add مع حجم التستر ومتجر المصدر
- إذا لا → لا تقترح التستر

**قاعدة 3 — صورة التستر (حرجة):**
صورة التستر تُؤخذ **حرفياً وآلياً** من حقل image_url للمنتج الأساسي في قائمتنا.
إذا كان المنتج الأساسي يحتوي أكثر من صورة (مفصولة بفاصلة) → خذ الأولى فقط.
لا تبحث عن صورة جديدة للتستر أبداً — الصورة تأتي من العطر الأساسي دائماً.

**قاعدة 4 — التساتر بلا عطر أساسي:**
لكل تستر في قائمتنا الحالية، تحقق: هل يوجد عطر أساسي (غير تستر) بنفس الاسم؟
- إذا لا → ضع اقتراح إضافة العطر الأساسي في orphan_testers

## قواعد المنتجات الناقصة
- قارن قائمتنا الكاملة بما يبيعه المنافسون السعوديون لنفس الماركة
- ركّز على: الأكثر مبيعاً أولاً، ثم الإصدارات الجديدة، ثم الأحجام الشائعة المختلفة
- كل مقترح يجب أن يكون **متوفراً للشراء الآن** في متجر محدد (اذكر المتجر في source_store)
- لا تخترع عطراً — كل مقترح يجب أن تراه فعلاً في أحد المتاجر السعودية

## صرامة مطلقة ضد الاختراع
- لا تخترع عطراً أو سعراً أو رابط صورة غير موجود فعلياً
- إذا لم تجد معلومة موثوقة، اترك الحقل فارغاً ولا تخمّن
- كل اقتراح يجب أن يكون موجوداً في متجر سعودي محدد

## أسلوب الكتابة — تعلّم من هذه الأمثلة الحقيقية من متجرنا
{writing_dna}

## قوالب HTML الإلزامية (للمنتجات الجديدة فقط)
### قالب العطور الجديدة/الأساسية:
{HTML_TEMPLATE_NEW}

### قالب التساتر الجديدة:
{HTML_TEMPLATE_TESTER}

## المتاجر السعودية للمقارنة (حصراً):
{competitors_json}

## قواعد التسعير
- تستر لمنتج أقل من 1000 ريال: السعر الأساسي ناقص 70 ريال
- تستر لمنتج 1000 ريال فأكثر: السعر الأساسي ناقص 150 ريال

## تعليمات الإخراج (إلزامية)
- JSON صارم فقط يبدأ بـ {{ وينتهي بـ }}, بلا markdown, بلا نص خارج JSON
- products_updated: [] دائماً — لا استثناء
- جميع الأسعار بالريال السعودي
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


def _norm_ar(s: str) -> str:
    """Normalize Arabic for column matching: alef variants, ya, ta marbuta, diacritics."""
    s = str(s).strip().lower()
    s = re.sub(r'[ً-ٰٟ]', '', s)  # diacritics
    s = s.replace('أ', 'ا').replace('إ', 'ا').replace('آ', 'ا').replace('ٱ', 'ا')
    s = s.replace('ى', 'ي').replace('ئ', 'ي')
    s = s.replace('ة', 'ه')
    s = re.sub(r'\s+', ' ', s)
    return s


def find_col(df: pd.DataFrame, key: str) -> str | None:
    """Find column by Arabic keywords with alef-normalization.

    For 'name'/'price', exclude option columns (containing 'خيار'/'option'/brackets like [1]).
    For 'images', exclude description-of-image columns ('وصف صورة').
    """
    keywords = ARABIC_COL_KEYS.get(key, [])
    EXCLUDE = {
        'name':   ['خيار', 'option', '[1]', '[2]', '[3]'],
        'price':  ['خيار', 'option', 'تكلفه', 'مخفض'],
        'images': ['وصف صوره', 'وصف صورة'],
    }
    excludes = [_norm_ar(x) for x in EXCLUDE.get(key, [])]

    cols_norm = [(col, _norm_ar(col)) for col in df.columns]

    # Pass 1: exact normalized match
    for kw in keywords:
        kn = _norm_ar(kw)
        for col, cn in cols_norm:
            if cn == kn and not any(x in cn for x in excludes):
                return col

    # Pass 2: substring normalized match, excluding bad ones
    for kw in keywords:
        kn = _norm_ar(kw)
        for col, cn in cols_norm:
            if kn in cn and not any(x in cn for x in excludes):
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
    if not isinstance(name, str) or not name.strip():
        return False
    n = name.lower()
    if any(t in n for t in ['tester', 'testr']):
        return True
    # Arabic forms with possible alef hamza variants
    n_norm = re.sub(r'[أإآ]', 'ا', n)
    return any(t in n_norm for t in ['تستر', 'تستير', 'تيستر'])


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


def extract_writing_dna(df: pd.DataFrame, max_samples: int = 5) -> str:
    """
    استخرج 'بصمة الكتابة' من ملف المنتجات لتعليم Gemini الأسلوب الصحيح.
    تأخذ عينات حقيقية من الأوصاف + التصنيفات + أنماط التسمية.
    """
    name_col  = find_col(df, 'name')
    desc_col  = find_col(df, 'description')
    cat_col   = find_col(df, 'category')
    brand_col = get_brand_col(df)
    price_col = find_col(df, 'price')

    samples = []
    if name_col and desc_col:
        for _, row in df.iterrows():
            name = str(row.get(name_col, ''))
            desc = str(row.get(desc_col, ''))
            # فقط منتجات أساسية (غير تساتر) ذات وصف HTML حقيقي
            if not is_tester(name) and len(desc) > 300 and '<' in desc:
                samples.append({
                    'name':     name,
                    'brand':    str(row.get(brand_col, '')) if brand_col else '',
                    'category': str(row.get(cat_col,   '')) if cat_col  else '',
                    'price':    row.get(price_col, 0)       if price_col else 0,
                    'desc_html': desc[:800],
                })
            if len(samples) >= max_samples:
                break

    all_categories = []
    if cat_col:
        all_categories = sorted(df[cat_col].dropna().astype(str).unique().tolist())

    dna = "### التصنيفات المتاحة في المتجر (استخدمها حرفياً دون تعديل):\n"
    dna += "\n".join(f"- {c}" for c in all_categories) if all_categories else "- غير محدد"
    dna += "\n\n### أمثلة حقيقية من أوصاف متجر مهووس (انسخ الأسلوب والتنسيق والمصطلحات):\n"

    for i, s in enumerate(samples, 1):
        dna += f"""
--- مثال {i} ---
الاسم: {s['name']}
الماركة: {s['brand']}
التصنيف: {s['category']}
السعر: {s['price']} ريال
مقطع الوصف (HTML):
{s['desc_html']}
--- نهاية المثال {i} ---
"""
    if not samples:
        dna += "\n(لا توجد أمثلة — استخدم القوالب الإلزامية المرفقة)\n"

    return dna


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


def call_gemini_brand(
    brand_name: str,
    products: list,              # كل منتجات الماركة (أساسية + تساتر)
    full_brand_products: list,   # نفس القائمة — لمنع التهلوس
    api_key: str,
    writing_dna: str,            # بصمة الكتابة المستخرجة من الملف
    model_name: str = 'gemini-2.5-flash',
    use_grounding: bool = True,
    include_missing_search: bool = True,
    batch_index: int = 0,
    total_batches: int = 1,
) -> dict:
    """
    Call Gemini API for a single brand batch.
    المهمة الوحيدة: اكتشاف الفجوات (تساتر ناقصة + منتجات ناقصة).
    لا يُعيد كتابة أوصاف المنتجات الموجودة أبداً.
    """
    genai.configure(api_key=api_key)

    # حقن بصمة الكتابة ديناميكياً في system instruction
    system_instruction = SYSTEM_INSTRUCTION_TEMPLATE.format(
        writing_dna=writing_dna,
        HTML_TEMPLATE_NEW=HTML_TEMPLATE_NEW,
        HTML_TEMPLATE_TESTER=HTML_TEMPLATE_TESTER,
        competitors_json=json.dumps(COMPETITOR_STORES, ensure_ascii=False, indent=2),
    )

    model_kwargs = dict(
        model_name=model_name,
        system_instruction=system_instruction,
    )
    if use_grounding:
        model_kwargs['tools'] = [genai.protos.Tool(
            google_search_retrieval=genai.protos.GoogleSearchRetrieval()
        )]

    model = genai.GenerativeModel(**model_kwargs)

    # فصل العطور الأساسية عن التساتر لإرسالهما بوضوح لـ Gemini
    base_perfumes  = [p for p in full_brand_products if not p.get('is_tester', False)]
    tester_products = [p for p in full_brand_products if p.get('is_tester', False)]

    base_catalog_json = json.dumps(
        [{'id': p['id'], 'name': p['name'], 'price': p['price'],
          'image_url': p.get('image_url', '')}
         for p in base_perfumes],
        ensure_ascii=False, indent=2
    )

    tester_catalog_json = json.dumps(
        [{'id': p['id'], 'name': p['name'], 'price': p['price']}
         for p in tester_products],
        ensure_ascii=False, indent=2
    )

    competitors_str = json.dumps(COMPETITOR_STORES, ensure_ascii=False)

    missing_section = ""
    if include_missing_search:
        missing_section = f"""
## المهمة 3: المنتجات الناقصة (مرة واحدة فقط لكل ماركة)
قارن قائمتنا الكاملة ({len(full_brand_products)} منتج) بما يبيعه المنافسون السعوديون لماركة "{brand_name}":
{competitors_str}

الأولوية: الأكثر مبيعاً أولاً → الإصدارات الجديدة → الأحجام الشائعة المختلفة.
لكل منتج ناقص: ابحث عن image_url_1 (زجاجة لحالها بخلفية بيضاء) و image_url_2 (زجاجة + كرتون).
اكتب وصفاً بقالب العطور الجديدة. اذكر المتجر السعودي المصدر في source_store."""
    else:
        missing_section = """
## المهمة 3: المنتجات الناقصة
تخطَّ هذه المهمة في الدفعة الحالية → missing_products: [] (نُفِّذت في الدفعة الأولى)."""

    prompt = f"""أنت تعالج ماركة "{brand_name}" — الدفعة {batch_index + 1} من {total_batches}.

## العطور الأساسية لدينا (غير التساتر) — {len(base_perfumes)} عطر:
{base_catalog_json}

## التساتر الموجودة لدينا حالياً — {len(tester_products)} تستر:
{tester_catalog_json}

---
⚠️ تذكير: products_updated يجب أن يكون [] فارغاً دائماً — لا تُعيد كتابة وصف أي منتج موجود.

## المهمة 1: التساتر الناقصة
لكل عطر في "العطور الأساسية":
1. هل يوجد في "التساتر الموجودة" تستر بنفس الاسم (بأي صيغة عربية أو إنجليزية)؟
   - نعم → تخطّ (لا تفعل شيئاً)
   - لا → ابحث في المتاجر السعودية: هل يبيعون تستر لهذا العطر؟
     * نعم → أضفه في testers_to_add:
       - image_url: انسخه **حرفياً** من حقل image_url للعطر الأساسي (أول صورة إذا كانت مفصولة بفاصلة)
       - size_ml: حجم التستر الذي وجدته عند المنافس
       - new_price: مطبقاً قاعدة التسعير (ناقص 70 أو 150 حسب السعر)
       - new_description: قالب التستر الإلزامي مكتملاً
     * لا → لا تقترح التستر

## المهمة 2: التساتر التي ليس لها عطر أساسي
لكل تستر في "التساتر الموجودة":
1. هل يوجد في "العطور الأساسية" منتج بنفس الاسم (بدون كلمة تستر)؟
   - نعم → تخطّ
   - لا → أضف في orphan_testers: تستر موجود لكن بلا عطر أساسي
{missing_section}

**أعد JSON صارم فقط يبدأ بـ {{ وينتهي بـ }} بلا أي نص خارجه:**

{{
  "brand": "{brand_name}",
  "batch_index": {batch_index},
  "products_updated": [],
  "testers_to_add": [
    {{
      "base_product_id": "id من قائمة العطور الأساسية",
      "name": "اسم العطر تستر",
      "size_ml": 100,
      "original_price": 0,
      "new_price": 0,
      "image_url": "منسوخ حرفياً من image_url العطر الأساسي",
      "source_store": "اسم المتجر السعودي",
      "tester_available_in_market": true,
      "new_description": "<p>...</p>",
      "seo_title": "عنوان أقل من 60 حرف",
      "seo_description": "وصف ميتا أقل من 155 حرف"
    }}
  ],
  "orphan_testers": [
    {{
      "tester_product_id": "id التستر الذي ليس له أساسي",
      "tester_name": "اسم التستر",
      "suggested_base_name": "اسم العطر الأساسي المقترح إضافته"
    }}
  ],
  "missing_products": [
    {{
      "name": "اسم العطر الكامل",
      "type": "عطر مفرد",
      "category": "من التصنيفات المتاحة حرفياً",
      "price": 0,
      "size_ml": 100,
      "concentration": "EDP",
      "gender": "رجالي",
      "description": "<p>...</p>",
      "brand": "{brand_name}",
      "is_tester": false,
      "is_bestseller": true,
      "source_store": "اسم المتجر السعودي",
      "image_url_1": "زجاجة بخلفية بيضاء",
      "image_url_2": "زجاجة + كرتون بخلفية بيضاء",
      "seo_title": "عنوان أقل من 60 حرف",
      "seo_description": "وصف ميتا أقل من 155 حرف"
    }}
  ]
}}"""

    response = model.generate_content(
        prompt,
        generation_config=genai.GenerationConfig(temperature=0.0),
    )
    return extract_json(response.text)


def _normalize_perfume_name(name: str) -> str:
    """Normalize perfume name for duplicate detection across Arabic/English variants."""
    if not name:
        return ''
    s = str(name).lower().strip()
    # Remove diacritics and punctuation
    s = re.sub(r'[ً-ٰٟ]', '', s)
    s = re.sub(r'[^\w؀-ۿ\s]', ' ', s)
    # Normalize EDP/EDT variants
    replacements = {
        'eau de parfum': 'edp', 'أو دو بارفان': 'edp', 'بارفان': 'edp',
        'بارفيوم': 'edp', 'parfum': 'edp',
        'eau de toilette': 'edt', 'أو دو تواليت': 'edt', 'تواليت': 'edt',
        'pour homme': 'men', 'للرجال': 'men', 'رجالي': 'men',
        'pour femme': 'women', 'للنساء': 'women', 'نسائي': 'women',
        'tester': 'tstr', 'تستر': 'tstr',
    }
    for k, v in replacements.items():
        s = s.replace(k, v)
    # Strip Arabic 'ال'
    s = re.sub(r'\bال', '', s)
    # Collapse whitespace
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def filter_duplicates(result: dict, existing_products: list) -> dict:
    """Remove suggestions that already exist in our catalog (defense in depth)."""
    existing_norms = {_normalize_perfume_name(p.get('name', '')) for p in existing_products}

    def not_dup(item):
        norm = _normalize_perfume_name(item.get('name', ''))
        if not norm:
            return False
        for ex in existing_norms:
            if not ex:
                continue
            if norm == ex or (len(ex) > 8 and ex in norm) or (len(norm) > 8 and norm in ex):
                return False
        return True

    # فلترة المنتجات الناقصة
    if 'missing_products' in result:
        result['missing_products'] = [m for m in result['missing_products'] if not_dup(m)]

    # فلترة التساتر الجديدة (الحقل الجديد)
    if 'testers_to_add' in result:
        result['testers_to_add'] = [t for t in result['testers_to_add'] if not_dup(t)]

    # توافق خلفي مع testers_updated القديم (إذا وُجد)
    if 'testers_updated' in result:
        kept = []
        for t in result['testers_updated']:
            if t.get('is_new'):
                if not_dup(t):
                    kept.append(t)
            else:
                kept.append(t)
        result['testers_updated'] = kept

    # products_updated يُفرَّغ دائماً بغض النظر عما أرسله النموذج
    result['products_updated'] = []

    return result


def merge_batch_results(accum: dict, new: dict) -> dict:
    """Merge a new batch result into the accumulator for the brand."""
    if not accum:
        return {
            'brand':           new.get('brand', ''),
            'products_updated': [],   # دائماً فارغ
            'testers_to_add':  list(new.get('testers_to_add', [])),
            'orphan_testers':  list(new.get('orphan_testers', [])),
            'missing_products': list(new.get('missing_products', [])),
        }

    # products_updated يبقى فارغاً دائماً
    accum['products_updated'] = []

    # دمج testers_to_add مع منع التكرار بالاسم المُنظَّم
    existing_tester_norms = {
        _normalize_perfume_name(t.get('name', ''))
        for t in accum.get('testers_to_add', [])
    }
    for t in new.get('testers_to_add', []):
        norm = _normalize_perfume_name(t.get('name', ''))
        if norm and norm not in existing_tester_norms:
            accum.setdefault('testers_to_add', []).append(t)
            existing_tester_norms.add(norm)

    # دمج orphan_testers مع منع التكرار بالـ id
    existing_orphan_ids = {
        t.get('tester_product_id') for t in accum.get('orphan_testers', [])
    }
    for t in new.get('orphan_testers', []):
        oid = t.get('tester_product_id')
        if oid and oid not in existing_orphan_ids:
            accum.setdefault('orphan_testers', []).append(t)
            existing_orphan_ids.add(oid)

    # دمج missing_products مع منع التكرار بالاسم المُنظَّم
    existing_missing_norms = {
        _normalize_perfume_name(m.get('name', ''))
        for m in accum.get('missing_products', [])
    }
    for m in new.get('missing_products', []):
        norm = _normalize_perfume_name(m.get('name', ''))
        if norm and norm not in existing_missing_norms:
            accum.setdefault('missing_products', []).append(m)
            existing_missing_norms.add(norm)

    return accum


def build_output_excel(result: dict, original_df: pd.DataFrame, template_bytes: bytes) -> bytes:
    """
    Build Salla-compatible Excel — يحتوي فقط على المنتجات الجديدة المقترحة.
    لا يُعيد تصدير أي منتج موجود مسبقاً في قائمتنا.
    """
    brand_col = get_brand_col(original_df)
    name_col  = find_col(original_df, 'name')
    price_col = find_col(original_df, 'price')
    desc_col  = find_col(original_df, 'description')
    cat_col   = find_col(original_df, 'category')
    qty_col   = find_col(original_df, 'quantity')
    img_col   = find_col(original_df, 'images')

    brand_name = result.get('brand', '')
    all_cols   = list(original_df.columns)
    rows       = []

    # 1. التساتر الجديدة المقترحة فقط
    for tester in result.get('testers_to_add', []):
        nr = {c: '' for c in all_cols}
        if name_col:  nr[name_col]  = tester.get('name', '')
        if price_col: nr[price_col] = tester.get('new_price', 0)
        if desc_col:  nr[desc_col]  = tester.get('new_description', '')
        if brand_col: nr[brand_col] = brand_name
        if cat_col:   nr[cat_col]   = 'العطور > عطور التساتر'
        if qty_col:   nr[qty_col]   = 10
        if img_col:
            img = tester.get('image_url', '')
            # Fallback: ابحث عن صورة المنتج الأساسي من original_df
            if not img and tester.get('base_product_id') and 'No.' in original_df.columns:
                base_match = original_df[
                    original_df['No.'].astype(str) == str(tester['base_product_id'])
                ]
                if not base_match.empty and img_col:
                    raw = str(base_match.iloc[0].get(img_col, '') or '')
                    img = raw.split(',')[0].strip()   # أول صورة فقط
            nr[img_col] = img
        rows.append(pd.Series(nr))

    # 2. المنتجات الناقصة الجديدة (صورتان مفصولتان بفاصلة لـ Salla)
    for missing in result.get('missing_products', []):
        nr = {c: '' for c in all_cols}
        if name_col:  nr[name_col]  = missing.get('name', '')
        if price_col: nr[price_col] = missing.get('price', 0)
        if desc_col:  nr[desc_col]  = missing.get('description', '')
        if brand_col: nr[brand_col] = missing.get('brand', brand_name)
        if cat_col:   nr[cat_col]   = missing.get('category', '')
        if qty_col:   nr[qty_col]   = 10
        if img_col:
            imgs = [missing.get('image_url_1', ''), missing.get('image_url_2', '')]
            nr[img_col] = ','.join(u for u in imgs if u)
        rows.append(pd.Series(nr))

    # إذا لم تكن هناك منتجات جديدة → أعد القالب فارغاً
    if not rows:
        buf = io.BytesIO(template_bytes)
        buf.seek(0)
        return buf.read()

    output_df = pd.DataFrame(rows)

    # تحميل قالب سلة وكتابة الصفوف
    wb = load_workbook(io.BytesIO(template_bytes))
    ws = wb.active

    # اكتشاف صف الرأس (يحتوي 'اسم' أو 'No.')
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

    col_map = {}
    for t_idx, t_hdr in enumerate(template_headers):
        if not t_hdr:
            continue
        t_str = str(t_hdr)
        for df_col in output_df.columns:
            if t_str in str(df_col) or str(df_col) in t_str:
                col_map[t_idx + 1] = df_col
                break

    for r_idx, (_, row) in enumerate(output_df.iterrows()):
        excel_row = data_start + r_idx
        for t_col, df_col in col_map.items():
            val = row.get(df_col, '')
            if not isinstance(val, str) and pd.isna(val):
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
        'model_name': 'gemini-2.5-flash',
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
        ['gemini-2.5-flash', 'gemini-2.5-pro', 'gemini-2.5-flash-lite', 'gemini-1.5-pro', 'gemini-1.5-flash'],
        index=0,
        help="gemini-2.5-flash: أسرع وأرخص | gemini-2.5-pro: أدق وأشمل",
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
    n_tst = len(result.get('testers_to_add', []))
    n_orp = len(result.get('orphan_testers', []))
    n_mis = len(result.get('missing_products', []))

    st.success(
        f"✅ اكتملت معالجة **{current_brand}** | "
        f"🏷️ {n_tst} تستر جديد · ⚠️ {n_orp} تستر بلا أساسي · 🔍 {n_mis} منتج ناقص"
    )

    tabs = st.tabs([
        f"🏷️ التساتر الجديدة ({n_tst})",
        f"⚠️ تساتر بلا عطر أساسي ({n_orp})",
        f"🔍 المنتجات الناقصة ({n_mis})",
    ])

    with tabs[0]:
        testers = result.get('testers_to_add', [])
        if testers:
            df_t = pd.DataFrame(testers)
            show_cols = [c for c in ['name', 'size_ml', 'original_price', 'new_price', 'source_store', 'tester_available_in_market']
                         if c in df_t.columns]
            st.dataframe(
                df_t[show_cols] if show_cols else df_t,
                use_container_width=True, hide_index=True
            )
            with st.expander("👁️ معاينة أول وصف HTML للتستر"):
                if testers[0].get('new_description'):
                    st.markdown(testers[0]['new_description'], unsafe_allow_html=True)
        else:
            st.info("✅ كل منتجاتك لديها تساتر — لا يوجد ناقص")

    with tabs[1]:
        orphans = result.get('orphan_testers', [])
        if orphans:
            st.warning("⚠️ هذه التساتر موجودة في متجرك لكن ليس لها عطر أساسي — يُنصح بإضافة العطر الأساسي:")
            st.dataframe(pd.DataFrame(orphans), use_container_width=True, hide_index=True)
        else:
            st.info("✅ كل التساتر لديها عطر أساسي مطابق")

    with tabs[2]:
        missing = result.get('missing_products', [])
        if missing:
            df_m = pd.DataFrame(missing)
            show_cols = [c for c in ['name', 'category', 'price', 'size_ml', 'concentration', 'gender', 'is_bestseller', 'source_store']
                         if c in df_m.columns]
            st.dataframe(
                df_m[show_cols] if show_cols else df_m,
                use_container_width=True, hide_index=True
            )
            with st.expander("👁️ معاينة أول وصف HTML للمنتج الناقص"):
                if missing[0].get('description'):
                    st.markdown(missing[0]['description'], unsafe_allow_html=True)
        else:
            st.info("✅ لا توجد منتجات ناقصة")

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
        # Clean up autosave for this brand
        try:
            _safe = re.sub(r'[^\w]', '_', current_brand)
            _p = os.path.join(".mahwous_autosave", f"{_safe}.json")
            if os.path.exists(_p):
                os.remove(_p)
        except Exception:
            pass
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

brand_lbl.markdown("**الخطوة 1/3:** جاري تحليل أسلوب الكتابة وتصنيف المنتجات...")
brand_bar.progress(5)

# استخرج بصمة الكتابة من الملف الكامل (مرة واحدة فقط)
writing_dna = extract_writing_dna(df)

# Build products payload — الكل (أساسي + تساتر) مع تمييز النوع
img_col_now = find_col(df, 'images')
products_payload = []
if name_col:
    for _, row in brand_df.iterrows():
        raw_img = str(row.get(img_col_now, '') or '') if img_col_now else ''
        prod_name = str(row.get(name_col, ''))
        products_payload.append({
            'id':        str(row.get('No.', row.name)),
            'name':      prod_name,
            'price':     float(pd.to_numeric(row.get(price_col, 0), errors='coerce') or 0),
            'description': '',        # لا نرسل الوصف الحالي — نوفّر tokens
            'image_url': raw_img.split(',')[0].strip(),   # أول صورة فقط
            'is_tester': is_tester(prod_name),
        })

n = len(products_payload)

# ─── BATCHING ────────────────────────────────────────────────────────────────
BATCH_SIZE = 25
batches = [products_payload[i:i + BATCH_SIZE] for i in range(0, n, BATCH_SIZE)] or [[]]
total_batches = len(batches)

brand_bar.progress(10)
brand_lbl.markdown(f"**الخطوة 1/3:** {n} منتج → {total_batches} دفعة (حجم الدفعة {BATCH_SIZE})")

# Auto-save key for resuming
SAVE_DIR = ".mahwous_autosave"
os.makedirs(SAVE_DIR, exist_ok=True)
safe_brand_key = re.sub(r'[^\w]', '_', current_brand)
autosave_path = os.path.join(SAVE_DIR, f"{safe_brand_key}.json")

accumulated = {}
# Resume from autosave if present
if os.path.exists(autosave_path):
    try:
        with open(autosave_path, 'r', encoding='utf-8') as f:
            accumulated = json.load(f)
    except Exception:
        accumulated = {}

start_batch = accumulated.get('_completed_batches', 0)

try:
    for b_idx in range(start_batch, total_batches):
        batch = batches[b_idx]
        brand_lbl.markdown(f"**الدفعة {b_idx + 1}/{total_batches}:** إرسال إلى Gemini AI...")
        prod_bar.progress((b_idx) / total_batches)
        prod_lbl.markdown(f"📦 الدفعة {b_idx + 1}/{total_batches} — {len(batch)} منتج")
        status_msg.info(
            f"🔍 الدفعة {b_idx + 1}/{total_batches} — {current_brand}\n"
            f"{'🔎 فحص التساتر الناقصة + المنتجات عند المنافسين' if b_idx == 0 else '🔎 فحص التساتر الناقصة'}\n"
            f"⚠️ لا يتم تحديث أوصاف المنتجات الموجودة"
        )

        batch_result = call_gemini_brand(
            brand_name=current_brand,
            products=batch,
            full_brand_products=products_payload,
            api_key=st.session_state.api_key,
            writing_dna=writing_dna,
            model_name=st.session_state.model_name,
            use_grounding=True,
            include_missing_search=(b_idx == 0),
            batch_index=b_idx,
            total_batches=total_batches,
        )
        batch_result = filter_duplicates(batch_result, products_payload)
        accumulated = merge_batch_results(accumulated, batch_result)
        accumulated['_completed_batches'] = b_idx + 1

        # Auto-save after each batch
        with open(autosave_path, 'w', encoding='utf-8') as f:
            json.dump(accumulated, f, ensure_ascii=False, indent=2)

        prod_bar.progress((b_idx + 1) / total_batches)

    result = {k: v for k, v in accumulated.items() if not k.startswith('_')}

    brand_bar.progress(75)
    prod_bar.progress(0.7)
    brand_lbl.markdown("**الخطوة 3/3:** معالجة النتائج...")

    n_upd = 0  # دائماً صفر — لا تحديث للموجود
    n_tst = len(result.get('testers_to_add', []))
    n_orp = len(result.get('orphan_testers', []))
    n_mis = len(result.get('missing_products', []))

    # شريط التقدم النهائي
    brand_bar.progress(100)
    prod_bar.progress(1.0)
    brand_lbl.markdown(f"✅ **اكتملت معالجة {current_brand}!**")
    prod_lbl.markdown(f"🏷️ {n_tst} تستر جديد · ⚠️ {n_orp} تستر بلا أساسي · 🔍 {n_mis} منتج ناقص")
    status_msg.success(f"🎯 اكتملت المعالجة — {n_tst} تستر جديد | {n_orp} تستر بلا أساسي | {n_mis} منتج ناقص")

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
            fallback_acc = {}
            for b_idx, batch in enumerate(batches):
                br = call_gemini_brand(
                    brand_name=current_brand,
                    products=batch,
                    full_brand_products=products_payload,
                    api_key=st.session_state.api_key,
                    writing_dna=writing_dna,
                    model_name=st.session_state.model_name,
                    use_grounding=False,
                    include_missing_search=(b_idx == 0),
                    batch_index=b_idx,
                    total_batches=len(batches),
                )
                br = filter_duplicates(br, products_payload)
                fallback_acc = merge_batch_results(fallback_acc, br)
            result = {k: v for k, v in fallback_acc.items() if not k.startswith('_')}
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
