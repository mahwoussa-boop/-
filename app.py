import streamlit as st
import pandas as pd
import json
import io
import os
import re
import time
import threading
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from google import genai
from google.genai import types as genai_types
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from difflib import SequenceMatcher

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
    "https://www.noon.com/saudi-ar/", "https://www.amazon.sa/",
    "https://en.ounass.com/saudi-arabia/", "https://www.namshi.com/sa-ar/",
    "https://www.brandsforless.com/en-sa/", "https://www.sivvi.com/en-sa/",
    "https://haraj.com.sa/", "https://shukran.com/",
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

# ─── SYSTEM INSTRUCTION TEMPLATE (writing_dna injected dynamically) ──────────
SYSTEM_INSTRUCTION_TEMPLATE = """## هويتك ومهمتك
أنت **خبير عطور محترف بخبرة 20 سنة** + محلل تنافسي لمتجر مهووس في السوق السعودي.
مهمتك **اكتشاف الفجوات فقط** — التساتر الناقصة والمنتجات الناقصة — وكتابة وصف احترافي **فقط للمنتجات الجديدة** التي تقترحها.

## ❌ ممنوع منعاً باتاً
- لا تُعيد كتابة أو تُحدّث وصف أي منتج موجود مسبقاً في قائمتنا
- products_updated يجب أن يكون دائماً [] قائمة فارغة
- لا تقترح منتجاً موجوداً مسبقاً ولو بصيغة مختلفة

## قواعد التساتر (الأهم)
**القاعدة 1 — فحص وجود التستر:**
- لكل عطر أساسي في قائمتنا، تحقق: هل يوجد منتج آخر في القائمة يحتوي اسمه على "تستر" أو "Tester" لنفس العطر؟
- إذا وُجد التستر → **تخطّ، لا تقترح شيئاً**
- إذا لم يُوجد → انتقل للقاعدة 2

**القاعدة 2 — البحث الشامل في السوق السعودي:**
- ابحث في Google بهذه الصيغ المتعددة:
  1. "[اسم العطر] tester Saudi Arabia buy"
  2. "[اسم العطر] تستر المملكة العربية السعودية"
  3. "[اسم الماركة] tester site:sa"
- ابحث في **أي** متجر سعودي يظهر في نتائج Google، سواء كان في قائمة المتاجر المرجعية أم لا
- سجّل اسم المتجر كما يظهر في URL النتيجة (مثال: "noon.com" أو "goldenscent.com")
- إذا لم تجد في أي مكان → لا تقترح التستر

**القاعدة 3 — صورة التستر:**
- الصورة تُؤخذ حرفياً من حقل image_url للمنتج الأساسي الموجود في قائمتنا
- لا تبحث عن صورة جديدة للتستر أبداً
- إذا كان للمنتج الأساسي أكثر من صورة (مفصولة بفاصلة)، خذ الأولى فقط

**القاعدة 4 — التساتر بلا عطر أساسي:**
- مرّ على كل تستر في قائمتنا
- تحقق: هل يوجد منتج أساسي (غير تستر) بنفس الاسم؟
- إذا لم يوجد → أضفه في missing_products مع وصفه كعطر أساسي جديد

## قواعد المنتجات الناقصة
- قارن قائمتنا الكاملة بما يبيعه المنافسون لنفس الماركة
- ركّز على: الأكثر مبيعاً، الإصدارات الجديدة، والأحجام المختلفة الشائعة
- كل مقترح يجب أن يكون متوفراً للشراء الآن في متجر محدد (اذكر المتجر)
- الأولوية للمنتجات الأكثر مبيعاً (bestsellers)

## أسلوب الكتابة — تعلّم من هذه الأمثلة الحقيقية
{writing_dna}

## قوالب HTML الإلزامية للمنتجات الجديدة فقط
### قالب العطور الجديدة/الأساسية:
{HTML_TEMPLATE_NEW}

### قالب التساتر الجديدة:
{HTML_TEMPLATE_TESTER}

## صرامة مطلقة ضد الاختراع
- لا تخترع عطراً أو سعراً أو رابط صورة غير موجود فعلياً
- إذا لم تجد معلومة موثوقة، اترك الحقل فارغاً
- كل مقترح يجب أن يكون موجوداً في متجر سعودي محدد

## 🚫 ممنوع التكرار الداخلي (قواعد صارمة)
- ممنوع منعاً باتاً تكرار نفس العطر أو التستر داخل المصفوفة. إذا وجدت العطر في أكثر من متجر منافس، اختر المتجر الأفضل أو الأرخص واذكره **مرة واحدة فقط**، وتجاهل البقية تماماً.
- لا تكرر نفس base_product_id في testers_to_add. كل base_product_id يظهر مرة واحدة فقط مهما تعددت المتاجر التي تبيع التستر.
- لا تكرر نفس المنتج في missing_products بصيغ مختلفة (مثل "Fame Parfum 80ml" و "فيم بارفان 80 مل") — كلها نفس المنتج، اذكره مرة واحدة فقط.

## 🏷️ قواعد source_store
- يُفضَّل أن يكون من قائمة المتاجر المرجعية competitors_json
- لكن إذا وجدت المنتج في متجر سعودي آخر موثوق، سجّله بنطاقه الفعلي (مثال: "noon.com")
- ممنوع اختراع متجر غير موجود — يجب أن يكون رابط المنتج قابلاً للتحقق

## 📐 صرامة مخطط JSON
- يجب أن تحتوي **جميع** عناصر missing_products على الحقلين image_url_1 و image_url_2 بشكل دائم. إذا لم تجد صورة ثانية، اجعل قيمتها سلسلة نصية فارغة "" — ولا تحذف المفتاح أبداً.
- جميع المفاتيح المذكورة في مخطط الإخراج إلزامية في كل عنصر؛ القيم الفارغة تُمثَّل بـ "" أو 0 وليس بحذف المفتاح.

## المتاجر السعودية للمقارنة (انسخ source_store حرفياً من هذه القائمة):
{competitors_json}

## قواعد التسعير
- تستر لمنتج أقل من 1000 ريال: السعر الأساسي ناقص 70 ريال
- تستر لمنتج 1000 ريال فأكثر: السعر الأساسي ناقص 150 ريال

## تعليمات الإخراج
- JSON صارم فقط، يبدأ بـ {{ وينتهي بـ }}
- لا markdown، لا نص خارج JSON
- products_updated: [] دائماً (لا تحديث للموجود)
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
    """استخرج 'بصمة الكتابة' من ملف المنتجات لتعليم Gemini الأسلوب."""
    name_col = find_col(df, 'name')
    desc_col = find_col(df, 'description')
    cat_col = find_col(df, 'category')
    brand_col = get_brand_col(df)
    price_col = find_col(df, 'price')

    samples = []
    if name_col and desc_col:
        for _, row in df.iterrows():
            name = str(row.get(name_col, ''))
            desc = str(row.get(desc_col, ''))
            if (not is_tester(name) and len(desc) > 200 and '<' in desc):
                samples.append({
                    'name': name,
                    'brand': str(row.get(brand_col, '')) if brand_col else '',
                    'category': str(row.get(cat_col, '')) if cat_col else '',
                    'price': row.get(price_col, 0) if price_col else 0,
                    'description_sample': desc[:800],
                })
            if len(samples) >= max_samples:
                break

    all_categories = []
    if cat_col:
        all_categories = sorted(df[cat_col].dropna().astype(str).unique().tolist())

    dna = "## أسلوب الكتابة المُتَّبع في متجر مهووس (تعلّم منه ولا تخرج عنه)\n\n"
    dna += "### التصنيفات المتاحة (استخدمها حرفياً):\n"
    dna += "\n".join(f"- {c}" for c in all_categories) + "\n\n"
    dna += "### أمثلة فعلية من أوصاف المنتجات (انسخ الأسلوب والتنسيق):\n"
    for i, s in enumerate(samples, 1):
        dna += (
            f"\n--- مثال {i} ---\n"
            f"الاسم: {s['name']}\n"
            f"الماركة: {s['brand']}\n"
            f"التصنيف: {s['category']}\n"
            f"السعر: {s['price']} ريال\n"
            f"مقطع من الوصف:\n{s['description_sample']}\n---\n"
        )
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
        snippet = (text[:300] + '...') if len(text) > 300 else text
        raise ValueError(
            f"لم يُرجع Gemini JSON صالحاً. مقتطف من الرد ({len(text)} حرف): {snippet!r}"
        )
    body = text[start:end + 1]
    try:
        return json.loads(body, strict=False)
    except json.JSONDecodeError:
        pass
    # Use json_repair (handles LLM-broken JSON: unescaped quotes, trailing commas, etc.)
    try:
        from json_repair import repair_json
        repaired = repair_json(body, return_objects=True)
        if isinstance(repaired, dict):
            return repaired
    except Exception:
        pass
    # Fallback regex repairs
    repaired = re.sub(r',(\s*[}\]])', r'\1', body)
    repaired = re.sub(r'(?<!\\)\\(?![\\/"bfnrtu])', r'\\\\', repaired)
    return json.loads(repaired, strict=False)


def call_gemini_brand(
    brand_name: str,
    products: list,
    full_brand_products: list,
    api_key: str,
    writing_dna: str,
    model_name: str = 'gemini-2.5-flash',
    use_grounding: bool = True,
    batch_index: int = 0,
    total_batches: int = 1,
    progress_cb=None,
) -> dict:
    """Call Gemini API for a single brand batch with retry logic.

    التغييرات الجوهرية:
    - 3 محاولات مع exponential backoff (2**attempt ثانية بين المحاولات)
    - عند 429/ResourceExhausted: انتظار إضافي 30 ثانية
    - تعطيل use_grounding يحدث فقط عند PERMISSION_DENIED أو INVALID_ARGUMENT
    - عند فشل كل المحاولات: ترجع {} مع رسالة خطأ واضحة (لا بيانات وهمية)
    """
    client = genai.Client(api_key=api_key)

    system_instruction = SYSTEM_INSTRUCTION_TEMPLATE.format(
        writing_dna=writing_dna,
        HTML_TEMPLATE_NEW=HTML_TEMPLATE_NEW,
        HTML_TEMPLATE_TESTER=HTML_TEMPLATE_TESTER,
        competitors_json=json.dumps(COMPETITOR_STORES, ensure_ascii=False),
    )

    base_perfumes = [p for p in full_brand_products if not is_tester(p.get('name', ''))]
    tester_products = [p for p in full_brand_products if is_tester(p.get('name', ''))]

    base_catalog_json = json.dumps(
        [{'id': p['id'], 'name': p['name'], 'price': p.get('price', 0),
          'image_url': (p.get('image_url') or '').split(',')[0].strip()}
         for p in base_perfumes],
        ensure_ascii=False, indent=2
    )

    tester_catalog_json = json.dumps(
        [{'id': p['id'], 'name': p['name'], 'price': p.get('price', 0)}
         for p in tester_products],
        ensure_ascii=False, indent=2
    )

    prompt = f"""أنت تعالج ماركة "{brand_name}" — الدفعة {batch_index + 1} من {total_batches}.

## قائمة العطور الأساسية لدينا (غير التساتر) — {len(base_perfumes)} عطر:
{base_catalog_json}

## التساتر الموجودة لدينا حالياً — {len(tester_products)} تستر:
{tester_catalog_json}

---

## المهمة 1: اكتشاف التساتر الناقصة
لكل عطر في قائمة "العطور الأساسية":
1. تحقق: هل يوجد في "التساتر الموجودة" تستر يحمل نفس الاسم (بأي صيغة)؟
   - إذا نعم → لا تفعل شيئاً (تخطّ)
   - إذا لا → تابع:
2. ابحث في المتاجر السعودية: هل يبيعون تستر لهذا العطر؟
   - إذا نعم → أضفه في testers_to_add مع:
     * name: اسم العطر + " تستر"
     * size_ml: حجم التستر الموجود عند المنافس
     * base_product_id: id العطر الأساسي من قائمتنا
     * image_url: انسخه حرفياً من حقل image_url للعطر الأساسي (الصورة الأولى فقط)
     * original_price: سعر العطر الأساسي من قائمتنا
     * new_price: مطبقاً قاعدة التسعير
     * source_store: اسم المتجر السعودي
     * new_description: قالب التستر مكتملاً
   - إذا لا → لا تقترح التستر

## المهمة 2: التساتر التي ليس لها عطر أساسي
لكل تستر في "التساتر الموجودة":
1. تحقق: هل يوجد في "العطور الأساسية" منتج بنفس الاسم (بدون كلمة تستر)؟
   - إذا نعم → تخطّ
   - إذا لا → أضف المنتج الأساسي في missing_products كعطر جديد يجب إضافته
     * ابحث عن صورة الزجاجة من المتاجر السعودية أو الموقع الرسمي للماركة
     * اكتب وصفاً بقالب العطور الجديدة

## المهمة 3: المنتجات الناقصة عند المنافسين
قارن قائمتنا الكاملة ({len(full_brand_products)} منتج) بما يبيعه المنافسون السعوديون لماركة "{brand_name}".
- 🔴 الأولوية القصوى: إصدارات 2024 و2025 و2026 — ابحث بالاسم الصريح مثل "Million Gold" و"Phantom Intense" و"Phantom Elixir" وما صدر حديثاً
- ركّز على: الأكثر مبيعاً، الأحجام الشائعة (50مل، 100مل، 200مل)، الإصدارات الجديدة
- لا تكتفِ بعطور 2022 و2023 — ابحث صراحةً عن "{brand_name} new release 2024 2025 2026" في المتاجر السعودية
- اقترح فقط المنتجات المتوفرة للشراء الآن مع ذكر المتجر المصدر
- لكل منتج مقترح: اكتب وصفاً كاملاً بقالب العطور الجديدة

## ⚠️ استراتيجية البحث الإلزامية — لا تتجاوزها
لكل عطر تبحث عنه، **افتح هذه المتاجر بالترتيب** وابحث فيها فعلياً:
1. https://www.noon.com/saudi-ar/ — ابحث: "[اسم العطر] {brand_name}"
2. https://en.ounass.com/saudi-arabia/ — ابحث: "{brand_name} perfume"
3. https://www.goldenscent.com/ — ابحث مباشرةً باسم العطر
4. https://niceonesa.com/ — ابحث مباشرةً باسم العطر
5. https://www.amazon.sa/ — ابحث: "{brand_name} [perfume name] عطر"

**قاعدة صارمة:** لا تكتفِ بمتجرَين. إذا لم تجد في الأول، انتقل للثاني والثالث.
**إذا وجدت المنتج في noon أو ounass — هذا اكتشاف ممتاز، سجّله.**

## ⚠️ تحذير نهائي قبل الإخراج
- قبل إرجاع JSON، راجع المصفوفات وتأكد:
  1. لا يوجد base_product_id مكرر داخل testers_to_add (واحد فقط لكل عطر).
  2. لا يوجد منتج مكرر داخل missing_products بأي صيغة (عربي/إنجليزي/أحجام مختلفة بنفس المنتج).
  3. حقل source_store منسوخ حرفياً من قائمة المتاجر — لا اختراع.
  4. كل عنصر في missing_products يحتوي image_url_1 و image_url_2 (الثاني قد يكون "" لكنه موجود).

**أعد JSON صارم فقط:**

{{
  "brand": "{brand_name}",
  "batch_index": {batch_index},
  "products_updated": [],
  "testers_to_add": [
    {{
      "base_product_id": "id من قائمتنا",
      "name": "اسم العطر تستر",
      "size_ml": 100,
      "original_price": 0,
      "new_price": 0,
      "image_url": "منسوخ حرفياً من العطر الأساسي",
      "source_store": "اسم المتجر السعودي",
      "tester_available_in_market": true,
      "new_description": "<p>...</p>",
      "seo_title": "عنوان أقل من 60 حرف",
      "seo_description": "وصف أقل من 155 حرف"
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
      "seo_description": "وصف أقل من 155 حرف"
    }}
  ]
}}"""

    # ─── Retry Loop (التغيير الجوهري) ───────────────────────────────────────
    MAX_RETRIES = 3
    last_error = None

    for attempt in range(MAX_RETRIES):
        try:
            # Configure each attempt — use_grounding قد يتغير بين المحاولات
            config_kwargs = dict(
                system_instruction=system_instruction,
                temperature=0.0,
                max_output_tokens=65536,
            )
            if use_grounding:
                config_kwargs['tools'] = [
                    genai_types.Tool(google_search=genai_types.GoogleSearch())
                ]
            else:
                config_kwargs['response_mime_type'] = 'application/json'

            config = genai_types.GenerateContentConfig(**config_kwargs)

            stream = client.models.generate_content_stream(
                model=model_name,
                contents=prompt,
                config=config,
            )

            text = ''
            last_chunk = None
            for chunk in stream:
                last_chunk = chunk
                t = ''
                try:
                    t = chunk.text or ''
                except Exception:
                    for cand in getattr(chunk, 'candidates', []) or []:
                        content = getattr(cand, 'content', None)
                        if not content:
                            continue
                        for part in getattr(content, 'parts', []) or []:
                            t += getattr(part, 'text', '') or ''
                if t:
                    text += t
                    if progress_cb:
                        try:
                            progress_cb(len(text))
                        except Exception:
                            pass

            finish = ''
            try:
                finish = str(last_chunk.candidates[0].finish_reason) if last_chunk else ''
            except Exception:
                pass

            if not text.strip():
                raise ValueError(f"Gemini أعاد رداً فارغاً (finish_reason={finish})")

            # نجاح → ارجع النتيجة المُحلَّلة
            return extract_json(text)

        except Exception as e:
            last_error = e
            err_str = str(e).lower()
            err_type = type(e).__name__.lower()

            # تشخيص نوع الخطأ
            is_rate_limit = (
                '429' in err_str
                or 'resourceexhausted' in err_type
                or 'resource_exhausted' in err_str
                or 'quota' in err_str
                or 'rate limit' in err_str
                or 'rate_limit' in err_str
            )
            is_config_error = (
                'permission_denied' in err_str
                or 'permissiondenied' in err_type
                or 'invalid_argument' in err_str
                or 'invalidargument' in err_type
            )

            # المحاولة الأخيرة → اخرج للسطر النهائي
            if attempt == MAX_RETRIES - 1:
                break

            # 1) Rate limit → 30 ثانية + exponential backoff
            if is_rate_limit:
                wait = 30 + (2 ** attempt)
                print(
                    f"[call_gemini_brand][{brand_name}] ⏳ Rate limit "
                    f"(محاولة {attempt + 1}/{MAX_RETRIES}). انتظار {wait}ث..."
                )
                time.sleep(wait)
                continue

            # 2) Config error + grounding مُفعَّل → عطّل grounding وأعد المحاولة
            #    هذه الحالة الوحيدة المسموح فيها بتعطيل البحث
            if is_config_error and use_grounding:
                print(
                    f"[call_gemini_brand][{brand_name}] ⚠️ خطأ تكوين Grounding "
                    f"(محاولة {attempt + 1}/{MAX_RETRIES}). تعطيله للمحاولة التالية فقط."
                )
                use_grounding = False
                time.sleep(2 ** attempt)
                continue

            # 3) أي خطأ آخر → exponential backoff فقط (Grounding يبقى مُفعَّلاً)
            wait = 2 ** attempt
            print(
                f"[call_gemini_brand][{brand_name}] 🔄 خطأ مؤقت "
                f"(محاولة {attempt + 1}/{MAX_RETRIES}): {str(e)[:150]}. انتظار {wait}ث..."
            )
            time.sleep(wait)

    # كل المحاولات فشلت → ارجع {} بدلاً من رفع استثناء أو إرجاع بيانات مخترعة
    print(
        f"[call_gemini_brand][{brand_name}] ❌ فشلت جميع المحاولات "
        f"({MAX_RETRIES}/{MAX_RETRIES}). آخر خطأ: {last_error}. "
        f"إرجاع نتيجة فارغة بدلاً من بيانات قد تكون مخترعة."
    )
    return {}


def _normalize_perfume_name(name: str) -> tuple:
    """تطبيع اسم العطر مع استخراج الحجم منفصلاً.

    التغيير: ترجع tuple (normalized_name, size) بدلاً من str.
    Example: "Dior Sauvage EDP 100ml" -> ("dior sauvage", "100ml")
    """
    if not name:
        return ('', '')

    s = str(name).lower().strip()

    # 1) توحيد الأرقام العربية إلى لاتينية
    arabic_to_latin = str.maketrans('٠١٢٣٤٥٦٧٨٩', '0123456789')
    s = s.translate(arabic_to_latin)

    # 2) إزالة التشكيل العربي
    s = re.sub(r'[ً-ٰٟ]', '', s)

    # 3) استخراج الحجم في متغير منفصل قبل أي تنظيف آخر
    size = ''
    size_match = re.search(r'(\d+)\s*(?:ml|مل)\b', s, flags=re.IGNORECASE)
    if size_match:
        size = f"{size_match.group(1)}ml"
        s = re.sub(r'\d+\s*(?:ml|مل)\b', ' ', s, flags=re.IGNORECASE)

    # 4) إزالة مصطلحات التركيز (من الأطول للأقصر لتجنّب القطع الجزئي)
    concentration_terms = [
        'eau de parfum', 'eau de toilette', 'eau de cologne',
        'parfum', 'perfume', 'cologne',
        'edp', 'edt', 'edc',
        'أو دو برفيوم', 'أو دو بارفيوم', 'أو دو بارفان',
        'أو دو تواليت', 'أو دو كولونيا',
        'برفيوم', 'بارفيوم', 'بارفان', 'تواليت', 'كولونيا',
    ]
    for term in sorted(concentration_terms, key=len, reverse=True):
        s = s.replace(term, ' ')

    # 5) معالجة الأرقام المكتوبة بالحروف (1 Million, 212, إلخ)
    _arabic_num_map = {
        'ون': '1', 'واحد': '1', 'وان': '1',
        'تو': '2', 'اثنين': '2', 'اثنان': '2',
        'ثري': '3', 'ثلاثة': '3', 'ثلاث': '3',
        'فور': '4', 'اربعة': '4', 'اربع': '4',
        'فايف': '5', 'خمسة': '5', 'خمس': '5',
        'سيكس': '6', 'ستة': '6',
        'سيفن': '7', 'سبعة': '7',
        'ايت': '8', 'ثمانية': '8',
        'ناين': '9', 'تسعة': '9', 'تسع': '9',
        'تن': '10', 'عشرة': '10', 'عشر': '10',
    }
    s = ' '.join(_arabic_num_map.get(w, w) for w in s.split())

    # 6) إزالة كلمات الضوضاء
    junk_words = {
        'للرجال', 'للنساء', 'رجالي', 'نسائي', 'النسائي', 'الرجالي',
        'عطر', 'تستر', 'tester', 'testr',
        'بخاخ', 'spray', 'للجنسين', 'unisex',
        'قديم', 'جديد', 'النسخه', 'النسخة',
        'intense', 'إنتنس', 'انتنس',
    }
    for j in junk_words:
        s = re.sub(rf'\b{re.escape(j)}\b', ' ', s)

    # كلمات مع أل التعريف
    clean_words = []
    for w in s.split():
        cw = w[2:] if w.startswith('ال') and len(w) > 3 else w
        if w in junk_words or cw in junk_words:
            continue
        clean_words.append(cw)
    s = ' '.join(clean_words)

    # 7) إزالة علامات الترقيم وتطبيع الفراغات
    s = re.sub(r'[^\w\u0600-\u06FF\s]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()

    return (s, size)


def filter_duplicates(result: dict, existing_products: list) -> dict:
    """شبكة أمان ضد تكرارات Gemini والاختراعات.

    التغييرات الجوهرية:
    - يستخدم SequenceMatcher (difflib) بعتبة 85% للمطابقة التقريبية
    - الحجم المُستخرج يجب أن يتطابق تماماً (أو يكون فارغاً عند الطرفين)
    - يطبع كل زوج محذوف لـ stdout: [DEDUP] '...' ≈ '...'
    - يستخدم difflib فقط من stdlib (بدون مكتبات خارجية)
    """
    SIMILARITY_THRESHOLD = 0.85

    # طبع كل المنتجات الموجودة في كاتالوجنا (norm + size)
    existing_normed = []
    for p in existing_products:
        n, sz = _normalize_perfume_name(p.get('name', ''))
        if n:
            existing_normed.append((n, sz, p.get('name', '')))

    def matches_existing(norm_name: str, size: str) -> bool:
        """هل المنتج المُقترح مطابق لشيء موجود في الكاتالوج؟"""
        if not norm_name:
            return True  # نُسقط الأسماء الفارغة دائماً
        for ex_n, ex_sz, _ in existing_normed:
            if not ex_n:
                continue
            # شرط الحجم: متطابق تماماً أو كلاهما فارغ
            same_size = (size == ex_sz) or (not size and not ex_sz)
            if not same_size:
                continue
            # شرط الاسم: تطابق نصي أو fuzzy ≥ 85%
            ratio = SequenceMatcher(None, norm_name, ex_n).ratio()
            if ratio >= SIMILARITY_THRESHOLD:
                return True
        return False

    def is_internal_duplicate(norm_name: str, size: str, seen: list) -> tuple:
        """هل هذا العنصر مكرر داخل المصفوفة الحالية؟"""
        if not norm_name:
            return (True, '')
        for s_n, s_sz, s_orig in seen:
            same_size = (size == s_sz) or (not size and not s_sz)
            if not same_size:
                continue
            ratio = SequenceMatcher(None, norm_name, s_n).ratio()
            if ratio >= SIMILARITY_THRESHOLD:
                return (True, s_orig)
        return (False, '')

    # ── 1. testers_to_add: dedup داخلي + مقارنة بالكاتالوج
    if 'testers_to_add' in result and isinstance(result['testers_to_add'], list):
        seen_base_ids = set()
        seen_norms = []  # list of (norm, size, original_name)
        kept = []
        for t in result['testers_to_add']:
            if not isinstance(t, dict):
                continue
            base_id = str(t.get('base_product_id', '') or '').strip()
            orig = t.get('name', '')
            n, sz = _normalize_perfume_name(orig)
            # السعر يكون من حقل size_ml — ندمجه مع الحجم لو كان مفقوداً
            if not sz and t.get('size_ml'):
                try:
                    sz = f"{int(t['size_ml'])}ml"
                except (ValueError, TypeError):
                    pass

            # دعدفة بالـ base_product_id (خط دفاع للتساتر عبر اللغات)
            if base_id and base_id in seen_base_ids:
                print(f"[DEDUP][tester base_id] '{orig}' (base_id={base_id} مكرر)")
                continue
            # دعدفة fuzzy داخلية
            is_dup, dup_orig = is_internal_duplicate(n, sz, seen_norms)
            if is_dup:
                print(f"[DEDUP][tester internal] '{orig}' ≈ '{dup_orig}'")
                continue
            # دعدفة ضد الكاتالوج
            if matches_existing(n, sz):
                print(f"[DEDUP][tester vs catalog] '{orig}' موجود مسبقاً")
                continue

            if base_id:
                seen_base_ids.add(base_id)
            seen_norms.append((n, sz, orig))
            kept.append(t)
        result['testers_to_add'] = kept

    # ── 2. missing_products: dedup داخلي + مقارنة بالكاتالوج + ضمان السكيما
    if 'missing_products' in result and isinstance(result['missing_products'], list):
        seen_norms = []
        kept = []
        for m in result['missing_products']:
            if not isinstance(m, dict):
                continue
            orig = m.get('name', '')
            n, sz = _normalize_perfume_name(orig)
            if not sz and m.get('size_ml'):
                try:
                    sz = f"{int(m['size_ml'])}ml"
                except (ValueError, TypeError):
                    pass

            is_dup, dup_orig = is_internal_duplicate(n, sz, seen_norms)
            if is_dup:
                print(f"[DEDUP][missing internal] '{orig}' ≈ '{dup_orig}'")
                continue
            if matches_existing(n, sz):
                print(f"[DEDUP][missing vs catalog] '{orig}' موجود مسبقاً في الكاتالوج")
                continue

            # ضمان حقول السكيما
            m.setdefault('image_url_1', '')
            m.setdefault('image_url_2', '')
            if m.get('image_url_2') is None:
                m['image_url_2'] = ''

            seen_norms.append((n, sz, orig))
            kept.append(m)
        result['missing_products'] = kept

    # ── 3. testers_updated (إن وُجد — نسخة قديمة)
    if 'testers_updated' in result and isinstance(result['testers_updated'], list):
        seen_norms = []
        kept = []
        for t in result['testers_updated']:
            if not isinstance(t, dict):
                continue
            orig = t.get('name', '')
            n, sz = _normalize_perfume_name(orig)
            is_dup, dup_orig = is_internal_duplicate(n, sz, seen_norms)
            if is_dup:
                print(f"[DEDUP][tester_updated] '{orig}' ≈ '{dup_orig}'")
                continue
            if t.get('is_new') and matches_existing(n, sz):
                continue
            seen_norms.append((n, sz, orig))
            kept.append(t)
        result['testers_updated'] = kept

    return result


def merge_batch_results(accum: dict, new: dict) -> dict:
    # If accum is empty, initialize it with empty arrays, DO NOT just return new
    if not accum:
        accum = {
            'brand': new.get('brand', ''),
            'products_updated': [],
            'testers_to_add': [],
            'orphan_testers': [],
            'missing_products': []
        }

    # 1. Hard filter for testers
    existing_ids = {str(t.get('base_product_id', '')) for t in accum.get('testers_to_add', [])}
    for t in new.get('testers_to_add', []):
        bid = str(t.get('base_product_id', ''))
        if bid and bid not in existing_ids:
            accum['testers_to_add'].append(t)
            existing_ids.add(bid)

    # 2. Hard filter for missing products (Internal Deduplication)
    existing_norms = {_normalize_perfume_name(m.get('name', ''))[0] for m in accum.get('missing_products', [])}
    for m in new.get('missing_products', []):
        norm = _normalize_perfume_name(m.get('name', ''))[0]
        if norm and norm not in existing_norms:
            accum['missing_products'].append(m)
            existing_norms.add(norm)

    # Also copy products_updated and orphan_testers if any
    if 'products_updated' in new:
        accum['products_updated'].extend(new['products_updated'])
    if 'orphan_testers' in new:
        accum['orphan_testers'].extend(new['orphan_testers'])

    return accum


def build_output_excel(
    result: dict,
    original_df: pd.DataFrame,
    template_bytes: bytes = None,  # اختياري الآن — لم يعد مستخدماً
) -> bytes:
    """يبني ملف Excel متوافق 100% مع منصة سلة من الصفر.

    التغييرات الجوهرية:
    - يستخدم قائمة SALLA_COLUMNS الثابتة كمرجع وحيد للأعمدة
    - يضمن الحقول الإلزامية: نوع المنتج, شحن, وزن, وحدة الوزن
    - يُسقط أي عمود غير موجود في SALLA_COLUMNS (مثل No. والكمية المتوفرة)
    - يضيف 'بيانات المنتج' في A1 و RTL على الورقة
    - تنظيف صارم لـ NaN/None/<NA> -> ""
    """
    # ─── المصدر الوحيد للحقيقة لأعمدة سلة ───────────────────────────────────
    SALLA_COLUMNS = [
        'النوع ', 'أسم المنتج', 'تصنيف المنتج', 'صورة المنتج',
        'وصف صورة المنتج', 'نوع المنتج', 'سعر المنتج', 'الوصف',
        'هل يتطلب شحن؟', 'رمز المنتج sku', 'سعر التكلفة', 'السعر المخفض',
        'تاريخ بداية التخفيض', 'تاريخ نهاية التخفيض',
        'اقصي كمية لكل عميل', 'إخفاء خيار تحديد الكمية',
        'اضافة صورة عند الطلب', 'الوزن', 'وحدة الوزن', 'الماركة',
        'العنوان الترويجي', 'تثبيت المنتج', 'الباركود', 'السعرات الحرارية',
        'MPN', 'GTIN', 'خاضع للضريبة ؟', 'سبب عدم الخضوع للضريبة',
        '[1] الاسم', '[1] النوع', '[1] القيمة', '[1] الصورة / اللون',
        '[2] الاسم', '[2] النوع', '[2] القيمة', '[2] الصورة / اللون',
        '[3] الاسم', '[3] النوع', '[3] القيمة', '[3] الصورة / اللون',
    ]

    # ─── helpers ────────────────────────────────────────────────────────────
    brand_col = get_brand_col(original_df)
    cat_col = find_col(original_df, 'category')
    img_col = find_col(original_df, 'images')
    brand_name = result.get('brand', '')

    def safe(v, default=''):
        """يحوّل أي قيمة إلى string نظيف بدون 'nan' / None / NA."""
        if v is None:
            return default
        try:
            if pd.isna(v):
                return default
        except (TypeError, ValueError):
            pass
        sv = str(v).strip()
        if sv.lower() in ('nan', 'none', '<na>', 'null'):
            return default
        return sv

    def get_base_row(base_id):
        if not base_id or 'No.' not in original_df.columns:
            return None
        match = original_df[original_df['No.'].astype(str) == str(base_id)]
        return match.iloc[0] if not match.empty else None

    def clean_category(cat):
        if not cat:
            return 'العطور'
        parts = [p.strip() for p in str(cat).split(',') if p.strip()]
        if not parts:
            return 'العطور'
        hierarchical = [p for p in parts if '>' in p]
        return hierarchical[0] if hierarchical else max(parts, key=len)

    default_category = 'العطور'
    if cat_col is not None and not original_df[cat_col].dropna().empty:
        default_category = clean_category(original_df[cat_col].dropna().mode().iloc[0])

    # ─── بناء الصفوف من نتيجة Gemini ────────────────────────────────────────
    rows = []

    # 1) التساتر الجديدة
    for t in result.get('testers_to_add', []):
        base_r = get_base_row(t.get('base_product_id'))

        # الصورة: من حقل t أولاً، ثم من المنتج الأساسي كـ fallback
        img = safe(t.get('image_url', ''))
        if not img and base_r is not None and img_col:
            raw = safe(base_r.get(img_col, ''))
            img = raw.split(',')[0].strip() if raw else ''

        # التصنيف: من المنتج الأساسي ثم default
        category = default_category
        if base_r is not None and cat_col:
            cv = safe(base_r.get(cat_col, ''))
            if cv:
                category = clean_category(cv)

        rows.append({
            'النوع ': 'منتج',
            'أسم المنتج': safe(t.get('name', '')),
            'تصنيف المنتج': category,
            'صورة المنتج': img,
            'نوع المنتج': 'منتج جاهز',
            'سعر المنتج': float(t.get('new_price', 0) or 0),
            'الوصف': safe(t.get('new_description', '')),
            'هل يتطلب شحن؟': 'نعم',
            'رمز المنتج sku': safe(t.get('sku', '')),
            'اقصي كمية لكل عميل': 0,
            'إخفاء خيار تحديد الكمية': 'لا',
            'الوزن': 1,
            'وحدة الوزن': 'kg',
            'الماركة': brand_name,
            'العنوان الترويجي': safe(t.get('seo_title', '')),
            'خاضع للضريبة ؟': 'نعم',
        })

    # 2) المنتجات الناقصة
    for m in result.get('missing_products', []):
        # دمج الصورتين (لو وُجدتا) بفاصلة
        img1 = safe(m.get('image_url_1', ''))
        img2 = safe(m.get('image_url_2', ''))
        images = ','.join(u for u in [img1, img2] if u)

        category_raw = safe(m.get('category', ''))
        category = clean_category(category_raw) if category_raw else default_category

        rows.append({
            'النوع ': 'منتج',
            'أسم المنتج': safe(m.get('name', '')),
            'تصنيف المنتج': category,
            'صورة المنتج': images,
            'نوع المنتج': 'منتج جاهز',
            'سعر المنتج': float(m.get('price', 0) or 0),
            'الوصف': safe(m.get('description', '')),
            'هل يتطلب شحن؟': 'نعم',
            'رمز المنتج sku': safe(m.get('sku', '')),
            'اقصي كمية لكل عميل': 0,
            'إخفاء خيار تحديد الكمية': 'لا',
            'الوزن': 1,
            'وحدة الوزن': 'kg',
            'الماركة': safe(m.get('brand', ''), brand_name),
            'العنوان الترويجي': safe(m.get('seo_title', '')),
            'خاضع للضريبة ؟': 'نعم',
        })

    # ─── تحويل لـ DataFrame وفرض ترتيب أعمدة سلة ────────────────────────────
    df = pd.DataFrame(rows) if rows else pd.DataFrame()

    # أضِف أي عمود مفقود من SALLA_COLUMNS كعمود فارغ
    for col in SALLA_COLUMNS:
        if col not in df.columns:
            df[col] = ''

    # اقطع على ترتيب SALLA_COLUMNS بالضبط (يُسقط أي عمود إضافي)
    df = df[SALLA_COLUMNS]

    # ─── تطبيق التنظيفات الإلزامية (Strict Salla Compliance) ────────────────
    if not df.empty:
        df['النوع '] = 'منتج'
        df['نوع المنتج'] = 'منتج جاهز'
        df['هل يتطلب شحن؟'] = 'نعم'
        df['الوزن'] = 1
        df['وحدة الوزن'] = 'kg'
        df['إخفاء خيار تحديد الكمية'] = df['إخفاء خيار تحديد الكمية'].replace('', 'لا')
        df['اقصي كمية لكل عميل'] = pd.to_numeric(
            df['اقصي كمية لكل عميل'], errors='coerce'
        ).fillna(0).astype(int)
        df['خاضع للضريبة ؟'] = df['خاضع للضريبة ؟'].replace('', 'نعم')

        # تنظيف شامل لكل أنواع الفراغات
        df = df.replace({
            pd.NA: '', None: '',
            'nan': '', 'NaN': '', 'NAN': '',
            'None': '', '<NA>': '', 'null': '', 'NULL': '',
        })
        df = df.fillna('')

    # ─── بناء ملف Excel من الصفر بأسلوب قالب سلة ────────────────────────────
    wb = Workbook()
    ws = wb.active
    ws.title = 'Salla Products Template Sheet'
    ws.sheet_view.rightToLeft = True  # RTL مطلوب لسلة

    n_cols = len(SALLA_COLUMNS)

    # الصف 1: عنوان "بيانات المنتج" مدموج عبر كل الأعمدة
    title_cell = ws.cell(row=1, column=1, value='بيانات المنتج')
    title_cell.font = Font(bold=True, size=14)
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    title_cell.fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)

    # الصف 2: أسماء الأعمدة
    header_font = Font(bold=True)
    header_fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
    for ci, header in enumerate(SALLA_COLUMNS, start=1):
        cell = ws.cell(row=2, column=ci, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # الصف 3+: بيانات المنتجات
    for ri, (_, row) in enumerate(df.iterrows(), start=3):
        for ci, col in enumerate(SALLA_COLUMNS, start=1):
            value = row[col]
            # ضمان نهائي: لا 'nan' كنص
            if isinstance(value, str) and value.lower() in ('nan', 'none', '<na>'):
                value = ''
            ws.cell(row=ri, column=ci, value=value)

    # عرض أعمدة معقول
    for ci in range(1, n_cols + 1):
        ws.column_dimensions[get_column_letter(ci)].width = 22

    # تجميد الصفين الأولين
    ws.freeze_panes = 'A3'

    # ─── حفظ في buffer ──────────────────────────────────────────────────────
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


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
        [
            'gemini-3-flash-preview',       # الأسرع من الجيل الجديد
            'gemini-3.1-pro-preview',       # الأدق — Gen 3
            'gemini-3-pro-preview',
            'gemini-flash-latest',          # alias لأحدث flash مستقر
            'gemini-pro-latest',
            'gemini-2.5-flash',
            'gemini-2.5-pro',
            'gemini-2.5-flash-lite',
            'gemini-2.0-flash',
        ],
        index=0,
        help="3-flash-preview: أسرع وأقوى | 3.1-pro-preview: أدق للبحث المعقد | 2.5-flash: مستقر ومضمون",
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
    n_testers = len(result.get('testers_to_add', []))
    n_orphans = len(result.get('orphan_testers', []))
    n_missing = len(result.get('missing_products', []))

    st.success(
        f"✅ اكتملت معالجة **{current_brand}** | "
        f"{n_testers} تستر جديد · {n_orphans} تستر بلا أساسي · {n_missing} منتج ناقص"
    )

    tabs = st.tabs([
        f"🏷️ التساتر الجديدة ({n_testers})",
        f"⚠️ تساتر بلا عطر أساسي ({n_orphans})",
        f"🔍 المنتجات الناقصة ({n_missing})",
    ])

    with tabs[0]:
        testers = result.get('testers_to_add', [])
        if testers:
            df_t = pd.DataFrame(testers)
            show_cols = [c for c in ['name', 'size_ml', 'original_price', 'new_price', 'source_store']
                         if c in df_t.columns]
            st.dataframe(df_t[show_cols] if show_cols else df_t,
                         use_container_width=True, hide_index=True)
            with st.expander("👁️ معاينة أول وصف HTML"):
                if testers[0].get('new_description'):
                    st.markdown(testers[0]['new_description'], unsafe_allow_html=True)
        else:
            st.info("✅ كل منتجاتك لديها تساتر — لا يوجد ناقص")

    with tabs[1]:
        orphans = result.get('orphan_testers', [])
        if orphans:
            st.warning("هذه التساتر موجودة في متجرك لكن ليس لها عطر أساسي:")
            st.dataframe(pd.DataFrame(orphans), use_container_width=True, hide_index=True)
        else:
            st.info("✅ كل التساتر لديها عطر أساسي مطابق")

    with tabs[2]:
        missing = result.get('missing_products', [])
        if missing:
            df_m = pd.DataFrame(missing)
            show_cols = [c for c in ['name', 'category', 'price', 'is_bestseller', 'source_store']
                         if c in df_m.columns]
            st.dataframe(df_m[show_cols] if show_cols else df_m,
                         use_container_width=True, hide_index=True)
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

writing_dna = extract_writing_dna(df)

brand_lbl.markdown("**الخطوة 1/3:** جاري تحليل أسلوب الكتابة وتجهيز بيانات المنتجات...")
brand_bar.progress(5)

# Build products payload (no description sent — saves tokens)
img_col_now = find_col(df, 'images')
products_payload = []
if name_col:
    for _, row in brand_df.iterrows():
        raw_img = str(row.get(img_col_now, '') or '') if img_col_now else ''
        products_payload.append({
            'id': str(row.get('No.', row.name)),
            'name': str(row.get(name_col, '')),
            'price': float(pd.to_numeric(row.get(price_col, 0), errors='coerce') or 0),
            'description': '',
            'image_url': raw_img.split(',')[0].strip(),
            'is_tester': is_tester(str(row.get(name_col, ''))),
        })

n = len(products_payload)

# ─── BATCHING ────────────────────────────────────────────────────────────────
BATCH_SIZE = 15
MAX_PARALLEL = 3
batches = [products_payload[i:i + BATCH_SIZE] for i in range(0, n, BATCH_SIZE)] or [[]]
total_batches = len(batches)

brand_bar.progress(10)
brand_lbl.markdown(
    f"**الخطوة 1/3:** {n} منتج → {total_batches} دفعة "
    f"(حجم {BATCH_SIZE} · {MAX_PARALLEL} متوازية)"
)

SAVE_DIR = ".mahwous_autosave"
os.makedirs(SAVE_DIR, exist_ok=True)
safe_brand_key = re.sub(r'[^\w]', '_', current_brand)
autosave_path = os.path.join(SAVE_DIR, f"{safe_brand_key}.json")

accumulated = {}
if os.path.exists(autosave_path):
    try:
        with open(autosave_path, 'r', encoding='utf-8') as f:
            accumulated = json.load(f)
    except Exception:
        accumulated = {}

completed_set = set(accumulated.get('_completed_batch_ids', []))

# ─── PARALLEL EXECUTION WITH LIVE STATUS ─────────────────────────────────────
# Capture session_state values BEFORE threading — workers can't access st.session_state
_api_key_val = st.session_state.api_key
_model_name_val = st.session_state.model_name

status_lock = threading.Lock()
batch_status = {
    i: {
        'state': 'done' if i in completed_set else 'pending',
        'chars': 0,
        'started_at': None,
        'finished_at': None,
        'error': '',
        'mode': 'grounding',
    }
    for i in range(total_batches)
}

def _run_one(b_idx):
    with status_lock:
        batch_status[b_idx]['state'] = 'running'
        batch_status[b_idx]['started_at'] = time.time()
        batch_status[b_idx]['mode'] = 'grounding'

    def cb(n_chars):
        with status_lock:
            batch_status[b_idx]['chars'] = n_chars

    common = dict(
        brand_name=current_brand,
        products=batches[b_idx],
        full_brand_products=products_payload,
        api_key=_api_key_val,
        writing_dna=writing_dna,
        model_name=_model_name_val,
        batch_index=b_idx,
        total_batches=total_batches,
        progress_cb=cb,
    )
    try:
        return call_gemini_brand(**common, use_grounding=True)
    except Exception as e1:
        msg = str(e1).lower()
        if any(x in msg for x in ['grounding', 'search', 'tool', 'billing', 'json', 'empty', 'فارغ']):
            with status_lock:
                batch_status[b_idx]['mode'] = 'no-grounding'
                batch_status[b_idx]['chars'] = 0
            return call_gemini_brand(**common, use_grounding=False)
        raise

status_panel = st.empty()

def _render_status():
    with status_lock:
        snap = {i: dict(s) for i, s in batch_status.items()}
    rows = []
    now = time.time()
    for i, s in snap.items():
        if s['state'] == 'pending':
            icon, info = '⏸️', 'في الانتظار'
        elif s['state'] == 'running':
            el = int(now - s['started_at']) if s['started_at'] else 0
            mode_lbl = '🌐 بحث' if s['mode'] == 'grounding' else '⚡ بدون بحث'
            info = f"{mode_lbl} · {el}ث · {s['chars']:,} حرف مستلم"
            icon = '🔄'
        elif s['state'] == 'done':
            dur = int((s['finished_at'] or now) - (s['started_at'] or now))
            icon, info = '✅', f'مكتمل في {dur}ث'
        else:
            icon, info = '❌', (s.get('error', '') or '')[:120]
        rows.append({
            'الدفعة': f"{i + 1}/{total_batches}",
            'الحالة': icon,
            'التفاصيل': info,
        })
    with status_panel.container():
        st.dataframe(
            pd.DataFrame(rows),
            use_container_width=True,
            hide_index=True,
        )

try:
    pending_ids = [i for i in range(total_batches) if i not in completed_set]
    results_by_idx = {}

    if pending_ids:
        with ThreadPoolExecutor(max_workers=MAX_PARALLEL) as ex:
            future_to_idx = {ex.submit(_run_one, i): i for i in pending_ids}

            while True:
                done_ids = []
                for fut, idx in list(future_to_idx.items()):
                    if fut.done() and idx not in results_by_idx:
                        try:
                            res = fut.result()
                            results_by_idx[idx] = res
                            with status_lock:
                                batch_status[idx]['state'] = 'done'
                                batch_status[idx]['finished_at'] = time.time()
                        except Exception as e:
                            with status_lock:
                                batch_status[idx]['state'] = 'error'
                                batch_status[idx]['error'] = str(e)
                                batch_status[idx]['finished_at'] = time.time()
                            results_by_idx[idx] = e
                        done_ids.append(idx)

                _render_status()
                done_count = sum(
                    1 for i in range(total_batches)
                    if i in completed_set or i in results_by_idx
                )
                brand_bar.progress(min(10 + int(done_count / total_batches * 65), 75))
                prod_bar.progress(done_count / max(total_batches, 1))
                prod_lbl.markdown(f"📦 {done_count}/{total_batches} دفعة مكتملة")

                if len(results_by_idx) == len(future_to_idx):
                    break
                time.sleep(0.7)

    # Collect — surface first error if any
    first_err = None
    for idx, res in results_by_idx.items():
        if isinstance(res, Exception):
            first_err = res
            continue
        merged = filter_duplicates(res, products_payload)
        accumulated = merge_batch_results(accumulated, merged)
        completed_set.add(idx)

    accumulated['_completed_batch_ids'] = sorted(completed_set)
    with open(autosave_path, 'w', encoding='utf-8') as f:
        json.dump(accumulated, f, ensure_ascii=False, indent=2)

    if first_err is not None and len(completed_set) < total_batches:
        raise first_err

    result = {k: v for k, v in accumulated.items() if not k.startswith('_')}

    brand_bar.progress(75)
    prod_bar.progress(0.7)
    brand_lbl.markdown("**الخطوة 3/3:** معالجة النتائج...")

    n_tst = len(result.get('testers_to_add', []))
    n_orph = len(result.get('orphan_testers', []))
    n_mis = len(result.get('missing_products', []))

    brand_bar.progress(100)
    prod_bar.progress(1.0)
    brand_lbl.markdown(f"✅ **اكتملت معالجة {current_brand}!**")
    prod_lbl.markdown(f"✅ {n_tst} تستر جديد · {n_orph} يتيم · {n_mis} ناقص")
    status_msg.success(
        f"🎯 اكتملت المعالجة — {n_tst} تستر جديد | {n_orph} تستر بلا أساسي | {n_mis} منتج ناقص"
    )

    st.session_state.current_result = result
    st.session_state.waiting_confirm = True
    st.session_state.processing = False
    time.sleep(0.5)
    st.rerun()

except Exception as e:
    err = str(e)
    brand_bar.progress(0)

    if 'api_key' in err.lower() or 'api key' in err.lower() or 'invalid' in err.lower():
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
