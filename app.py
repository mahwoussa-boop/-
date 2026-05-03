text = re.sub(r'\n?\s*```\s*$',          '', text, flags=re.MULTILINE)
    text = text.strip()

    start = text.find('{')
    end   = text.rfind('}')
    if start == -1 or end == -1:
        snippet = (text[:300] + '...') if len(text) > 300 else text
        raise ValueError(f"لم يُرجع Gemini JSON صالحاً ({len(text)} حرف): {snippet!r}")
    raw = text[start:end + 1]

    try:
        return json.loads(raw, strict=False)
    except (json.JSONDecodeError, ValueError):
        pass

    raw = re.sub(r'\\u(?![0-9a-fA-F]{4})', r'\\\\u', raw)
    raw = re.sub(r'\\(?!["\\/bfnrtu])', r'\\\\', raw)

    result, in_string, escape_next = [], False, False
    for ch in raw:
        if escape_next:
            result.append(ch); escape_next = False
        elif ch == '\\' and in_string:
            result.append(ch); escape_next = True
        elif ch == '"':
            in_string = not in_string; result.append(ch)
        elif in_string and ch == '\n':
            result.append('\\n')
        elif in_string and ch == '\r':
            result.append('\\r')
        elif in_string and ch == '\t':
            result.append('\\t')
        else:
            result.append(ch)
    cleaned = ''.join(result)
    cleaned = re.sub(r',(\s*[}\]])', r'\1', cleaned)

    try:
        return json.loads(cleaned, strict=False)
    except (json.JSONDecodeError, ValueError):
        pass

    cleaned2 = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', cleaned)
    try:
        return json.loads(cleaned2, strict=False)
    except (json.JSONDecodeError, ValueError):
        pass

    try:
        from json_repair import repair_json
        repaired = repair_json(cleaned2, return_objects=True)
        if isinstance(repaired, dict):
            return repaired
    except Exception:
        pass

    snippet = (cleaned2[:400] + '...') if len(cleaned2) > 400 else cleaned2
    raise ValueError(f"فشل تحليل JSON بعد كل المحاولات. المقتطف: {snippet!r}")


def test_gemini_key(api_key: str, model_name: str = 'gemini-2.5-flash') -> tuple[bool, str]:
    """يختبر مفتاح Gemini بطلب بسيط (5 توكن) للكشف المبكر عن المشاكل.

    Returns:
        (success, message) — message يحوي تشخيصاً عربياً مفصّلاً عند الفشل.
    """
    if not api_key or not api_key.strip():
        return False, "⚠️ لم يُدخل مفتاح API"

    if not api_key.startswith('AIza'):
        return False, (
            "⚠️ صيغة المفتاح غير صحيحة — مفاتيح Gemini تبدأ بـ 'AIza'. "
            "تأكد من نسخ المفتاح كاملاً من Google AI Studio."
        )

    try:
        client = genai.Client(api_key=api_key)
        config = genai_types.GenerateContentConfig(
            temperature=0.0, max_output_tokens=5,
        )
        resp = client.models.generate_content(
            model=model_name, contents='test', config=config,
        )
        if resp.text:
            return True, f"✅ المفتاح يعمل بنجاح مع {model_name}"
        return False, "⚠️ المفتاح صحيح لكن الرد فارغ — جرّب موديلاً آخر"
    except Exception as e:
        err = str(e).lower()

        # خطأ 403 PERMISSION_DENIED — المشروع محجوب
        if ('403' in err and 'denied access' in err) or 'permission_denied' in err:
            return False, (
                "❌ **المشروع محجوب من Google** (403 PERMISSION_DENIED)\n\n"
                "🔴 هذا الخطأ على مستوى حساب Google وليس في التطبيق.\n\n"
                "**الأسباب الشائعة:**\n"
                "• المفتاح تسرّب وحُجب تلقائياً (مثلاً ظهر على GitHub)\n"
                "• المشروع محجوب على مستوى Google Cloud\n"
                "• منطقتك تتطلب تفعيل Cloud Billing\n\n"
                "**الحل (بالترتيب):**\n"
                "1. اذهب إلى https://aistudio.google.com/apikey\n"
                "2. احذف المفتاح الحالي تماماً\n"
                "3. اضغط 'Create API Key' → 'Create in new project'\n"
                "4. استخدم المفتاح الجديد هنا\n\n"
                "إن استمرّ الخطأ، فعّل Billing من:\n"
                "https://console.cloud.google.com/billing"
            )

        # خطأ 401 / مفتاح غير صحيح
        if '401' in err or 'unauthenticated' in err or 'invalid api' in err or 'api_key_invalid' in err:
            return False, (
                "❌ **المفتاح غير صحيح أو منتهٍ**\n\n"
                "تحقق من:\n"
                "• نسخت المفتاح كاملاً بدون فراغات في الأطراف\n"
                "• المفتاح لم يُحذف من Google AI Studio\n"
                "• اذهب إلى https://aistudio.google.com/apikey وأنشئ مفتاحاً جديداً"
            )

        # خطأ تسريب
        if 'leaked' in err or 'reported as leaked' in err:
            return False, (
                "❌ **المفتاح مُبلَّغ عنه كمُسرَّب**\n\n"
                "Google رصدت هذا المفتاح في مكان عام (مثل GitHub) وحجبته.\n"
                "أنشئ مفتاحاً جديداً من https://aistudio.google.com/apikey\n"
                "وتأكد ألا تُضيفه أبداً لأي ملف يُرفع إلى مستودع عام."
            )

        # خطأ حصة
        if '429' in err or 'quota' in err or 'resource_exhausted' in err:
            return False, (
                "⚠️ **تجاوزت حصة الاستخدام**\n\n"
                "انتظر دقيقة وأعد المحاولة، أو جرّب موديل 'gemini-2.5-flash-lite' "
                "(حصته أكبر بـ 4 مرات)."
            )

        # خطأ شبكة
        if 'connection' in err or 'timeout' in err or 'network' in err:
            return False, f"⚠️ مشكلة اتصال بالإنترنت: {e}"

        # غير معروف
        return False, f"❌ خطأ غير متوقع: {e}"

# ═════════════════════════════════════════════════════════════════════════════
#  محرّك التطبيع والمقارنة التقريبية (Fuzzy Matching)
# ═════════════════════════════════════════════════════════════════════════════

_AR_VOWELS = set('اويىآأإٱةؤئءَُِّْ')

_AR_CONSONANTS = {
    'ب': 'b', 'ت': 't', 'ث': 't',
    'ج': 'g', 'ح': 'h', 'خ': 'k',
    'د': 'd', 'ذ': 'd',
    'ر': 'r', 'ز': 'z',
    'س': 's', 'ش': 's', 'ص': 's', 'ض': 'd',
    'ط': 't', 'ظ': 'z',
    'ع': '', 'غ': 'g',
    'ف': 'f', 'ق': 'k', 'ك': 'k',
    'ل': 'l', 'م': 'm', 'ن': 'n',
    'ه': 'h',
}

_LATIN_DIGRAPHS = [
    ('ph', 'f'), ('th', 't'), ('ch', 's'), ('sh', 's'),
    ('kh', 'k'), ('gh', 'g'), ('qu', 'k'),
]

_LATIN_REPLACE = [
    ('v', 'f'), ('j', 'g'), ('x', 'k'), ('q', 'k'),
]

_LATIN_VOWELS_DROP = set('aeiouy')

_LAT_PHRASES = [
    ('eau de parfum', 'edp'),
    ('eau de toilette', 'edt'),
    ('eau de cologne', 'edc'),
]

_JUNK_WORDS = {
    'للرجال', 'للنساء', 'رجالي', 'نسائي', 'النسائي', 'الرجالي',
    'عطر', 'العطر', 'تستر', 'تيستر', 'تستير', 'tester', 'testr',
    'مل', 'ml', 'بخاخ', 'spray', 'للجنسين', 'unisex',
    'قديم', 'جديد', 'الجديد', 'القديم', 'النسخه', 'النسخة',
    'edp', 'edt', 'edc',
    'pour', 'homme', 'femme', 'for', 'men', 'women', 'man', 'woman',
    'eau', 'de', 'la', 'le', 'دي', 'دو', 'لو',
    'الاصلي', 'original', 'authentic', 'اصلي',
}

_SIZE_RX = re.compile(
    r'(\d+(?:[.,]\d+)?)\s*(?:ml|مل|مللي|مللتر|milliliter)\b',
    re.IGNORECASE,
)


def _extract_size_ml(name: str) -> int:
    """يُرجع الحجم بالـ مل (0 إذا لم يُذكر)."""
    if not name:
        return 0
    m = _SIZE_RX.search(str(name))
    if not m:
        return 0
    try:
        return int(float(m.group(1).replace(',', '.')))
    except (ValueError, TypeError):
        return 0


def _normalize_perfume_name(name: str) -> str:
    """يُنتج «هيكلاً صوتياً» قابلاً للمقارنة عبر اللغات."""
    if not name:
        return ''
    s = str(name).lower().strip()
    s = re.sub(r'[\u064b-\u065f\u0670]', '', s)
    s = (s.replace('أ', 'ا').replace('إ', 'ا').replace('آ', 'ا').replace('ٱ', 'ا')
            .replace('ى', 'ي').replace('ة', 'ه'))
    s = _SIZE_RX.sub(' ', s)
    s = re.sub(r'\b\d+\b', ' ', s)
    for phrase, repl in _LAT_PHRASES:
        s = s.replace(phrase, repl)
    for pat, repl in _LATIN_DIGRAPHS:
        s = s.replace(pat, repl)
    for pat, repl in _LATIN_REPLACE:
        s = s.replace(pat, repl)

    out_chars = []
    for ch in s:
        if ch in _AR_VOWELS:
            continue
        elif ch in _AR_CONSONANTS:
            out_chars.append(_AR_CONSONANTS[ch])
        elif ch.isalnum() or ch.isspace():
            out_chars.append(ch)
    s = ''.join(out_chars)

    tokens = s.split()
    cleaned = []
    for tok in tokens:
        if tok.startswith('ال') and len(tok) > 3:
            tok = tok[2:]
        if not tok or tok in _JUNK_WORDS or len(tok) < 2:
            continue
        consonants = ''.join(c for c in tok if c not in _LATIN_VOWELS_DROP and c != 'h')
        if consonants:
            cleaned.append(consonants)

    skeleton = ''.join(cleaned)
    skeleton = re.sub(r'(.)\1+', r'\1', skeleton)
    return skeleton


def _name_similarity(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    if len(a) >= 4 and len(b) >= 4 and (a in b or b in a):
        ratio_len = min(len(a), len(b)) / max(len(a), len(b))
        if ratio_len >= 0.7:
            return 0.95
    return SequenceMatcher(None, a, b).ratio()


def filter_duplicates(result: dict, existing_products: list,
                      similarity_threshold: float = 0.85) -> dict:
    existing_keys = []
    for p in existing_products:
        nm = p.get('name', '')
        sk = _normalize_perfume_name(nm)
        sz = _extract_size_ml(nm)
        if sk:
            existing_keys.append((sk, sz))

    def _matches_existing(skeleton, size):
        if not skeleton:
            return True
        for ek, es in existing_keys:
            if size and es and size != es:
                continue
            if _name_similarity(skeleton, ek) >= similarity_threshold:
                return True
        return False

    def _is_internal_dup(skeleton, size, seen):
        if not skeleton:
            return True
        for sk, sz in seen:
            if size and sz and size != sz:
                continue
            if _name_similarity(skeleton, sk) >= similarity_threshold:
                return True
        return False

    # ⚠️ ملاحظة: التساتر تُدَقّق على عدم تكرار base_product_id فقط،
    # ولا تُدقَّق ضد existing_products لأنها بطبيعتها مشتقة من العطور الأساسية.
    testers = result.get('testers_to_add')
    if isinstance(testers, list):
        seen_base_ids: set = set()
        seen_keys: list = []
        deduped = []
        for t in testers:
            if not isinstance(t, dict):
                continue
            base_id = str(t.get('base_product_id', '') or '').strip()
            nm = t.get('name', '')
            sk = _normalize_perfume_name(nm)
            sz_from_name = _extract_size_ml(nm)
            try:
                sz_from_field = int(float(t.get('size_ml') or 0))
            except (ValueError, TypeError):
                sz_from_field = 0
            sz = sz_from_name or sz_from_field

            if base_id and base_id in seen_base_ids:
                continue
            if _is_internal_dup(sk, sz, seen_keys):
                continue

            # تأكد من وجود الحقول الجديدة
            t.setdefault('competitor_price', 0)
            t.setdefault('tester_available_in_market', False)
            t.setdefault('source_store', '')

            # وحّد صيغة الاسم: «تستر» في البداية وليس النهاية
            # (Gemini أحياناً تُلصقها في النهاية بصيغة خاطئة لسلة)
            if nm:
                normalized_name = build_tester_name(nm)
                if normalized_name != nm:
                    t['name'] = normalized_name

            if base_id:
                seen_base_ids.add(base_id)
            seen_keys.append((sk, sz))
            deduped.append(t)
        result['testers_to_add'] = deduped

    missing = result.get('missing_products')
    if isinstance(missing, list):
        seen_keys = []
        deduped = []
        for m in missing:
            if not isinstance(m, dict):
                continue
            nm = m.get('name', '')
            sk = _normalize_perfume_name(nm)
            sz_from_name = _extract_size_ml(nm)
            try:
                sz_from_field = int(float(m.get('size_ml') or 0))
            except (ValueError, TypeError):
                sz_from_field = 0
            sz = sz_from_name or sz_from_field

            if _is_internal_dup(sk, sz, seen_keys):
                continue
            if _matches_existing(sk, sz):
                continue

            m.setdefault('image_url_1', '')
            m.setdefault('image_url_2', '')
            if m.get('image_url_2') is None:
                m['image_url_2'] = ''

            seen_keys.append((sk, sz))
            deduped.append(m)
        result['missing_products'] = deduped

    upd = result.get('testers_updated')
    if isinstance(upd, list):
        seen_keys = []
        kept = []
        for t in upd:
            if not isinstance(t, dict):
                continue
            sk = _normalize_perfume_name(t.get('name', ''))
            sz = _extract_size_ml(t.get('name', ''))
            if _is_internal_dup(sk, sz, seen_keys):
                continue
            if t.get('is_new') and _matches_existing(sk, sz):
                continue
            seen_keys.append((sk, sz))
            kept.append(t)
        result['testers_updated'] = kept

    return result


# ═════════════════════════════════════════════════════════════════════════════
#  ⭐ شبكة الأمان: ضمان وجود تستر لكل عطر أساسي بدون تستر
# ═════════════════════════════════════════════════════════════════════════════

# regex لإزالة كلمات «الضوضاء» من الأسماء قبل التطبيع.
# ضرورية لأن `_normalize_perfume_name` يفحص `_JUNK_WORDS` *بعد* التحويل
# للحروف اللاتينية، لكن قائمة الكلمات تحتوي الأصل العربي — فلا يتطابق.
# نُزيلها هنا قبل التطبيع لتفادي اختلاف الهياكل بسبب كلمات حشو مثل
# "عطر" و "تستر" و "للرجال" إلخ.
_NOISE_STRIP_RX = re.compile(
    r'\b(?:tester|testr|test)\b'
    r'|تستر|تيستر|تستير'
    r'|العطر|عطر'
    r'|الرجالي|الرجالى|للرجال|رجالي|رجالى'
    r'|النسائي|النسائى|للنساء|نسائي|نسائى'
    r'|للجنسين|unisex'
    r'|الاصلي|الأصلي|اصلي|أصلي|original|authentic',
    re.IGNORECASE,
)

# alias للتوافق الخلفي
_TESTER_STRIP_RX = _NOISE_STRIP_RX


def _strip_tester_keyword(name: str) -> str:
    """يُزيل كلمات الحشو الشائعة (تستر/عطر/للرجال/الأصلي…) من الاسم قبل
    التطبيع. مهم جداً لأن `_normalize_perfume_name` لا يستطيع إزالة هذه
    الكلمات العربية بنفسه (لأنه يفحص قائمة `_JUNK_WORDS` بعد تحويل الأحرف
    إلى لاتينية، فلا يتطابق الأصل العربي).
    """
    if not name:
        return ''
    return _NOISE_STRIP_RX.sub(' ', str(name))


def ensure_all_testers_added(result: dict, products_payload: list) -> dict:
    """شبكة أمان نهائية: تضمن وجود تستر لكل عطر أساسي ليس له تستر،
    حتى لو فاتت Gemini أحدها. تستخدم قاعدة التسعير الداخلية إذا لم يكن
    هناك سعر منافس مرجعي.

    - `result`: قاموس النتائج المُجمّعة (بعد merge_batch_results)
    - `products_payload`: قائمة منتجات الماركة الكاملة
    """
    if not isinstance(result, dict):
        return result
    if not products_payload:
        return result

    # افصل العطور الأساسية عن التساتر الموجودة في القائمة الأصلية
    # ⚠️ نستثني الأطقم/المجموعات: لا نضيف تستر للأطقم.
    base_perfumes = [
        p for p in products_payload
        if isinstance(p, dict)
        and not is_tester(p.get('name', ''))
        and not is_set(p.get('name', ''))
    ]
    existing_tester_products = [
        p for p in products_payload
        if isinstance(p, dict) and is_tester(p.get('name', ''))
    ]

    # ⭐ قاعدة الحجم الأكبر: إذا كان نفس العطر متوفراً بأحجام متعددة،
    # نختار فقط الحجم الأكبر للحصول على تستر (لا تستر للأحجام الصغيرة).
    # نُجمّع العطور حسب هيكل الاسم بدون الحجم.
    size_groups: dict = {}  # {skeleton_no_size: [(idx, size, price, name), ...]}
    for bp in base_perfumes:
        bp_name = (bp.get('name', '') or '').strip()
        if not bp_name:
            continue
        sk_no_size = _normalize_perfume_name(_strip_tester_keyword(bp_name))
        sz = _extract_size_for_grouping(bp_name)
        size_groups.setdefault(sk_no_size, []).append(bp)

    # في كل مجموعة، احتفظ فقط بالحجم الأكبر
    largest_in_group = set()  # ids of base perfumes that should get a tester
    skipped_for_size = []
    for sk, items in size_groups.items():
        if not sk:  # مفقود الهيكل — تخطه
            continue
        if len(items) == 1:
            largest_in_group.add(id(items[0]))
            continue
        # احسب الأحجام واختر الأكبر
        sized = [(it, _extract_size_for_grouping(it.get('name', '') or '')) for it in items]
        max_size = max((s for _, s in sized), default=0)
        if max_size == 0:
            # لا أحجام معروفة — اختر الأول وتخطّى الباقي
            largest_in_group.add(id(items[0]))
            for it in items[1:]:
                skipped_for_size.append(it.get('name', ''))
            continue
        # احتفظ بكل من له الحجم الأكبر، تخطّ الأصغر
        for it, sz in sized:
            if sz == max_size:
                largest_in_group.add(id(it))
            else:
                skipped_for_size.append(it.get('name', ''))

    # مجموعة هياكل أسماء المنتجات التي لها تستر بالفعل في القائمة الأصلية
    has_tester_skeletons: set = set()
    for tp in existing_tester_products:
        raw = tp.get('name', '') or ''
        stripped = _strip_tester_keyword(raw)
        sk = _normalize_perfume_name(stripped)
        if sk:
            has_tester_skeletons.add(sk)

    # base_product_ids التي اقترحت لها Gemini تستر بالفعل
    existing_testers = result.get('testers_to_add', []) or []
    if not isinstance(existing_testers, list):
        existing_testers = []
    existing_base_ids = {
        str(t.get('base_product_id', '') or '').strip()
        for t in existing_testers
        if isinstance(t, dict) and str(t.get('base_product_id', '') or '').strip()
    }

    # هياكل أسماء التساتر التي اقترحتها Gemini (للحماية من التكرار حتى عند
    # غياب base_product_id)
    existing_tester_skeletons: set = set()
    for t in existing_testers:
        if not isinstance(t, dict):
            continue
        raw = t.get('name', '') or ''
        stripped = _strip_tester_keyword(raw)
        sk = _normalize_perfume_name(stripped)
        if sk:
            existing_tester_skeletons.add(sk)

    auto_added_names = []

    for bp in base_perfumes:
        bp_id = str(bp.get('id', '') or '').strip()
        bp_name = (bp.get('name', '') or '').strip()
        if not bp_name:
            continue

        # ⭐ تخطّ إذا لم يكن هذا الحجم الأكبر لمجموعته
        if id(bp) not in largest_in_group:
            continue

        bp_sk = _normalize_perfume_name(_strip_tester_keyword(bp_name))
        try:
            bp_price = float(bp.get('price', 0) or 0)
        except (TypeError, ValueError):
            bp_price = 0.0

        # 🖼️ احتفظ بقائمة الصور كاملة (مفصولة بفاصلة) — Salla يدعم صور متعددة
        raw_img = (bp.get('image_url') or '').strip()
        # نظّف ونزّل الفواصل
        if raw_img:
            img_list = [u.strip() for u in raw_img.split(',') if u.strip()]
            bp_img = ','.join(img_list)
        else:
            bp_img = ''

        # تخطّ إذا اقترحت Gemini له تستر بالفعل (عبر base_product_id)
        if bp_id and bp_id in existing_base_ids:
            continue
        # تخطّ إذا اقترحت Gemini له تستر بالفعل (عبر هيكل الاسم)
        if bp_sk and bp_sk in existing_tester_skeletons:
            continue
        # تخطّ إذا كان له تستر فعلاً في القائمة الأصلية
        if bp_sk and bp_sk in has_tester_skeletons:
            continue

        # أضف تستراً تلقائياً بقاعدة التسعير الداخلية
        tester_price = calc_tester_price(bp_price)
        size_ml = _extract_size_for_grouping(bp_name) or 100
        bp_brand = (bp.get('brand') or '').strip()  # قد يكون موجوداً في products_payload
        bp_category = (bp.get('category') or '').strip()

        # 📝 ابنِ وصفاً مُملوءاً بالكامل (بدون placeholders) كإفتراضي.
        # `enrich_auto_added_testers` لاحقاً تستبدل النوتات والعائلة بمعلومات من Gemini.
        clean_n = clean_perfume_display_name(bp_name)
        clean_b = clean_brand_display_name(bp_brand)
        gender_disp = detect_gender(bp_name, bp_category)
        complete_desc = fill_tester_template_complete(
            brand_name=bp_brand or '',  # يُحدَّث في build_output_excel
            perfume_name=bp_name,
            size_ml=size_ml,
            base_category=bp_category,
            enrichment=None,  # سيُعبَّأ لاحقاً
        )

        auto_tester = {
            'base_product_id': bp_id,
            'name': build_tester_name(bp_name),
            'size_ml': size_ml,
            'original_price': bp_price,
            'new_price': tester_price,
            'competitor_price': 0,
            'image_url': bp_img,  # ← قائمة كاملة بالصور
            'source_store': '',
            'tester_available_in_market': False,
            'new_description': complete_desc,  # ← مُملوء بدون placeholders
            'seo_title': build_tester_name(bp_name)[:60],
            'seo_description': (
                f"تستر {clean_n} من {clean_b} الأصلي 100% — نفس السائل "
                f"والثبات والفوحان للإصدار المغلف بسعر استثنائي."
                if clean_b else
                f"تستر {clean_n} الأصلي 100% — نفس السائل والثبات والفوحان "
                f"للإصدار المغلف بسعر استثنائي."
            )[:155],
            '_auto_added': True,
            '_base_category': bp_category,  # للاستخدام في تصنيف التستر لاحقاً
            '_base_perfume_name': bp_name,
        }
        existing_testers.append(auto_tester)
        if bp_id:
            existing_base_ids.add(bp_id)
        if bp_sk:
            existing_tester_skeletons.add(bp_sk)
        auto_added_names.append(bp_name)

    result['testers_to_add'] = existing_testers
    if auto_added_names:
        # سجل ما أُضيف تلقائياً للمراجعة لاحقاً
        prev = result.get('_auto_added_testers') or []
        if not isinstance(prev, list):
            prev = []
        result['_auto_added_testers'] = prev + auto_added_names
    if skipped_for_size:
        # سجل ما تم تخطيه بسبب وجود حجم أكبر
        result['_skipped_smaller_sizes'] = skipped_for_size

    return result


def merge_batch_results(accum: dict, new: dict) -> dict:
    if not accum:
        accum = {
            'brand': new.get('brand', ''),
            'products_updated': [],
            'testers_to_add': [],
            'orphan_testers': [],
            'missing_products': []
        }
    existing_ids = {str(t.get('base_product_id', '')) for t in accum.get('testers_to_add', [])}
    for t in new.get('testers_to_add', []):
        bid = str(t.get('base_product_id', ''))
        if bid and bid not in existing_ids:
            accum['testers_to_add'].append(t)
            existing_ids.add(bid)
    existing_norms = {_normalize_perfume_name(m.get('name', '')) for m in accum.get('missing_products', [])}
    for m in new.get('missing_products', []):
        norm = _normalize_perfume_name(m.get('name', ''))
        if norm and norm not in existing_norms:
            accum['missing_products'].append(m)
            existing_norms.add(norm)
    if 'products_updated' in new:
        accum['products_updated'].extend(new['products_updated'])
    if 'orphan_testers' in new:
        accum['orphan_testers'].extend(new['orphan_testers'])
    return accum


# ═════════════════════════════════════════════════════════════════════════════
#  call_gemini_brand — إعادة محاولة ذكية، البحث محفوظ دائماً
# ═════════════════════════════════════════════════════════════════════════════

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
    max_retries: int = 4,
    base_backoff_seconds: float = 6.0,
) -> dict:
    """يستدعي Gemini لمعالجة دفعة من ماركة واحدة.

    - إعادة محاولة بانتظار تصاعدي للأخطاء العابرة (6→18→54→162ث)
    - رفع فوري للأخطاء الصلبة (مفتاح API، فوترة، أمان)
    - عند استنفاد المحاولات → RuntimeError واضح
    - البحث (Grounding) محفوظ دائماً — لا يُعطّل عند الفشل
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

## المهمة 1: إضافة التستر لكل عطر أساسي بلا تستر (إلزامي)
لكل عطر في قائمة "العطور الأساسية":
1. تحقق: هل يوجد في "التساتر الموجودة" تستر يحمل نفس الاسم (بأي صيغة عربية أو إنجليزية)؟
   - إذا نعم → لا تفعل شيئاً (تخطّ).
   - إذا لا → تابع الخطوة 2 (إلزامي):

2. ابحث في المتاجر السعودية للحصول على **سعر مرجعي** للتستر، ثم **أضف التستر دائماً** في `testers_to_add` مع:
     * `name`: اسم العطر + " تستر"
     * `size_ml`: نفس حجم العطر الأساسي ما لم يُذكر حجم تستر مختلف عند منافس
     * `base_product_id`: id العطر الأساسي من قائمتنا (إلزامي ودقيق)
     * `image_url`: انسخه حرفياً من حقل image_url للعطر الأساسي (الصورة الأولى فقط) — لا تبحث عن صورة جديدة
     * `original_price`: سعر العطر الأساسي من قائمتنا
     * `new_price`: مطبقاً قاعدة التسعير الداخلية
       (أقل من 1000 → ناقص 70؛ 1000 فأكثر → ناقص 150)
     * `competitor_price`: سعر التستر عند المنافس (مرجعي فقط)، أو 0 إذا لم تجد
     * `tester_available_in_market`: true إذا وجدته عند منافس سعودي، false إذا لم تجد
     * `source_store`: اسم المتجر السعودي إن وُجد، أو "" إذا لم تجد
     * `new_description`: قالب التستر مكتملاً
     * `seo_title` و `seo_description`: قصيران ومحسّنان

   - 🔴 **لا تتجاوز هذه الخطوة أبداً** — كل عطر بلا تستر **يجب** يحصل على تستر، حتى لو لم تجد التستر عند أي منافس. ضع `competitor_price=0` و `tester_available_in_market=false` و `source_store=""` وأضفه.

## المهمة 2: التساتر التي ليس لها عطر أساسي (Orphan Testers)
لكل تستر في "التساتر الموجودة":
1. تحقق: هل يوجد في "العطور الأساسية" منتج بنفس الاسم (بدون كلمة تستر)؟
   - إذا نعم → تخطّ.
   - إذا لا → أضف المنتج الأساسي في `missing_products` كعطر جديد يجب إضافته:
     * ابحث عن صورة الزجاجة من المتاجر السعودية أو الموقع الرسمي للماركة
     * اكتب وصفاً بقالب العطور الجديدة
     * واذكر التستر اليتيم في `orphan_testers` للعرض في الواجهة.

## المهمة 3: المنتجات الناقصة عند المنافسين
قارن قائمتنا الكاملة ({len(full_brand_products)} منتج) بما يبيعه المنافسون السعوديون لماركة "{brand_name}".
- 🔴 الأولوية القصوى: إصدارات 2024 و2025 و2026 — ابحث بالاسم الصريح مثل "Million Gold" و"Phantom Intense" و"Phantom Elixir" وما صدر حديثاً
- ركّز على: الأكثر مبيعاً، الأحجام الشائعة (50مل، 100مل، 200مل)، الإصدارات الجديدة
- لا تكتفِ بعطور 2022 و2023 — ابحث صراحةً عن "{brand_name} new release 2024 2025 2026" في المتاجر السعودية
- اقترح فقط المنتجات المتوفرة للشراء الآن مع ذكر المتجر المصدر
- لكل منتج مقترح: اكتب وصفاً كاملاً بقالب العطور الجديدة

## ⚠️ استراتيجية البحث الإلزامية — لا تتجاوزها
لكل عطر تبحث عنه، **افتح هذه المتاجر بالترتيب** وابحث فيها فعلياً:
1. https://www.noon.com/saudi-ar/ — ابحث: "[اسم العطر] {brand_name} tester"
2. https://en.ounass.com/saudi-arabia/ — ابحث: "{brand_name} perfume tester"
3. https://www.goldenscent.com/ — ابحث مباشرةً باسم العطر + tester
4. https://niceonesa.com/ — ابحث مباشرةً باسم العطر + tester
5. https://www.amazon.sa/ — ابحث: "{brand_name} [perfume name] tester"

**قاعدة صارمة:** لا تكتفِ بمتجرَين. إذا لم تجد في الأول، انتقل للثاني والثالث.
**لكن: حتى لو لم تجد التستر في أي مكان، أضفه إلى testers_to_add بـ competitor_price=0.**

## ⚠️ تحذير نهائي قبل الإخراج
- قبل إرجاع JSON، راجع المصفوفات وتأكد:
  1. **كل عطر أساسي بلا تستر له عنصر مقابل في testers_to_add** (لا استثناءات).
  2. لا يوجد base_product_id مكرر داخل testers_to_add (واحد فقط لكل عطر).
  3. لا يوجد منتج مكرر داخل missing_products بأي صيغة (عربي/إنجليزي/أحجام مختلفة بنفس المنتج).
  4. كل عنصر في testers_to_add يحتوي `competitor_price` و `tester_available_in_market` (حتى لو 0 و false).
  5. كل عنصر في missing_products يحتوي image_url_1 و image_url_2 (الثاني قد يكون "" لكنه موجود).

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
      "competitor_price": 0,
      "image_url": "منسوخ حرفياً من العطر الأساسي",
      "source_store": "اسم المتجر السعودي أو \\"\\"",
      "tester_available_in_market": false,
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

    config_kwargs = dict(
        system_instruction=system_instruction,
        temperature=0.0,
        max_output_tokens=65536,
    )
    if use_grounding:
        config_kwargs['tools'] = [genai_types.Tool(google_search=genai_types.GoogleSearch())]
    else:
        config_kwargs['response_mime_type'] = 'application/json'

    config = genai_types.GenerateContentConfig(**config_kwargs)

    last_err = None

    for attempt in range(1, max_retries + 1):
        try:
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
            safety = ''
            try:
                finish = str(last_chunk.candidates[0].finish_reason) if last_chunk else ''
                safety = (
                    str(getattr(last_chunk.candidates[0], 'safety_ratings', ''))[:200]
                    if last_chunk else ''
                )
            except Exception:
                pass

            if not text.strip():
                if 'SAFETY' in finish:
                    raise RuntimeError(f"حُجبت الاستجابة بفلتر الأمان (safety={safety})")
                hint = ' — قلّل BATCH_SIZE.' if 'MAX_TOKENS' in finish else ''
                raise ValueError(f"Gemini أعاد رداً فارغاً (finish_reason={finish}){hint}")

            try:
                return extract_json(text)
            except (ValueError, json.JSONDecodeError) as e:
                raise ValueError(f"{e} | finish_reason={finish}") from e

        except Exception as e:
            last_err = e
            err_low = str(e).lower()

            HARD_ERRORS = (
                'api key', 'invalid api', 'api_key', 'unauthenticated',
                'authentication', 'permission_denied', 'permission denied',
                'billing', 'safety', 'حُجبت', 'blocked',
            )
            if any(x in err_low for x in HARD_ERRORS):
                raise RuntimeError(f"❌ خطأ صلب في Gemini (لا يُعاد المحاولة): {e}") from e

            # ⚠️ إذا تجاوزت الحصة اليومية (RPD)، لا فائدة من إعادة المحاولة —
            # الانتظار حتى منتصف الليل بتوقيت المحيط الهادئ ضروري.
            DAILY_QUOTA_INDICATORS = (
                'exceeded your current quota', 'exceeded your quota',
                'requests per day', 'rpd', 'daily quota', 'per_day',
                'generate_requests_per_model_per_day',
            )
            if '429' in err_low and any(x in err_low for x in DAILY_QUOTA_INDICATORS):
                raise RuntimeError(
                    f"❌ تجاوزت الحصة اليومية لـ Gemini (RPD).\n"
                    f"الحلول: (1) انتظر حتى 10:00 صباحاً بتوقيت السعودية لإعادة الحصة. "
                    f"(2) جرّب نموذج 'gemini-2.5-flash-lite' (حصة أكبر بـ 4 مرات). "
                    f"(3) فعّل Cloud Billing على حسابك (تيار 1) — يرفع الحد إلى 1,500 RPD.\n"
                    f"تفاصيل الخطأ: {e}"
                ) from e

            if attempt >= max_retries:
                break

            backoff = base_backoff_seconds * (3 ** (attempt - 1))
            if any(x in err_low for x in ('429', 'quota', 'rate', 'resource_exhausted',
                                          'too many', 'unavailable', '503', '500')):
                backoff *= 2
            backoff = min(backoff, 180.0)

            time.sleep(backoff)

    raise RuntimeError(
        f"فشل الاتصال بـ Gemini بعد {max_retries} محاولات للدفعة "
        f"{batch_index + 1}/{total_batches} من ماركة «{brand_name}». "
        f"البحث الإلزامي محفوظ — لم يتم التنازل عنه. آخر خطأ: {last_err}"
    )


# ═════════════════════════════════════════════════════════════════════════════
#  build_output_excel — توافق صارم مع منصة سلة
# ═════════════════════════════════════════════════════════════════════════════

SALLA_MANDATORY = {
    'نوع المنتج':         'منتج جاهز',
    'النوع':              'منتج',
    'هل يتطلب شحن؟':      'نعم',
    'يتطلب شحن؟':         'نعم',
    'يتطلب شحن':          'نعم',
    'الكمية':             10,
    'الكمية المتوفرة':    10,
    'أقصى كمية لكل عميل': 2,
    'الوزن':              1,
    'وحدة الوزن':         'kg',
    'إخفاء خيار التوصيل': 'لا',
    'حالة المنتج':        'نشط',
}

_NULL_LIKE = {'nan', 'NaN', 'NAN', 'None', 'NONE', 'none',
              '<NA>', '<na>', 'NaT', 'nat', 'null', 'NULL', 'undefined'}


def _clean_cell(v):
    """يُحوّل أي قيمة شبيهة بـ NaN إلى سلسلة فارغة ''."""
    if v is None:
        return ''
    try:
        if pd.isna(v):
            return ''
    except (TypeError, ValueError):
        pass
    sv = str(v).strip()
    if sv in _NULL_LIKE or sv == '':
        return ''
    return v if not isinstance(v, str) else sv


def _norm_hdr(s) -> str:
    """تطبيع رؤوس الأعمدة للمقارنة."""
    s = str(s).strip()
    s = (s.replace('أ', 'ا').replace('إ', 'ا').replace('آ', 'ا')
            .replace('ى', 'ي').replace('ة', 'ه'))
    s = re.sub(r'\s+', ' ', s)
    return s


def build_output_excel(result: dict, original_df: pd.DataFrame, template_bytes: bytes) -> bytes:
    """يبني ملف Excel متوافقاً 100% مع قالب سلة.

    يحتوي **فقط** على:
      • testers_to_add  → التساتر الجديدة المُقترحة (من Gemini أو شبكة الأمان)
      • missing_products → المنتجات الناقصة (موجودة عند المنافسين، غير موجودة لديك)

    لا يحتوي على:
      • products_updated → المنتجات الحالية في متجرك (لا تُكرّر إضافتها)
    """
    brand_col = get_brand_col(original_df)
    name_col  = find_col(original_df, 'name')
    price_col = find_col(original_df, 'price')
    desc_col  = find_col(original_df, 'description')
    cat_col   = find_col(original_df, 'category')
    img_col   = find_col(original_df, 'images')

    brand_name = result.get('brand', '')
    all_cols = list(original_df.columns)

    # 🛡️ مرشّح دفاعي: هياكل أسماء كل المنتجات الموجودة فعلاً في القائمة الأصلية
    # (سواء كانت عطور أساسية أو تساتر). نُسقط أي صف من المخرجات يطابق هياكل
    # موجودة، حتى لو فات ذلك على Gemini أو filter_duplicates.
    existing_skeletons: set = set()
    if name_col and name_col in original_df.columns:
        for raw in original_df[name_col].dropna().astype(str):
            sk = _normalize_perfume_name(_strip_tester_keyword(raw))
            if sk:
                existing_skeletons.add(sk)

    def get_safe_row(base_id):
        if not base_id or 'No.' not in original_df.columns:
            return None
        match = original_df[original_df['No.'].astype(str) == str(base_id)]
        return match.iloc[0] if not match.empty else None

    def apply_salla_mandatory(nr: dict) -> dict:
        for c in all_cols:
            cs = str(c).strip()
            ns = _norm_hdr(cs)
            matched = False
            for key, val in SALLA_MANDATORY.items():
                if cs == key or ns == _norm_hdr(key):
                    nr[c] = val
                    matched = True
                    break
            if matched:
                continue
            if 'نوع المنتج' in cs:
                nr[c] = 'منتج جاهز'
            elif cs == 'النوع':
                nr[c] = 'منتج'
            elif 'يتطلب شحن' in cs:
                nr[c] = 'نعم'
            # ⚠️ مهم: «إخفاء خيار تحديد الكمية» يحتوي كلمة «الكمية» — لذا يجب
            # فحصه قبل قاعدة الكمية حتى لا يأخذ القيمة 10 ويرفضه سلة.
            elif 'إخفاء خيار' in cs or 'اخفاء خيار' in ns:
                nr[c] = 'لا'
            elif 'اقصي كميه' in ns or 'اقصى كميه' in ns:
                nr[c] = 2
            elif ('الكمية' in cs or 'الكميه' in ns) and 'اقصي' not in ns and 'اقصى' not in ns:
                nr[c] = 10
            elif cs == 'الوزن':
                nr[c] = 1
            elif 'وحدة الوزن' in cs or 'وحده الوزن' in ns:
                nr[c] = 'kg'
            elif 'الماركة' in cs and brand_col and c == brand_col:
                nr[c] = brand_name
        return nr

    def _clean_category(cat: str) -> str:
        if not cat:
            return cat
        parts = [p.strip() for p in str(cat).split(',') if p.strip()]
        if not parts:
            return cat
        hierarchical = [p for p in parts if '>' in p]
        return hierarchical[0] if hierarchical else max(parts, key=len)

    rows = []
    skipped_existing = 0  # عدّاد للصفوف التي أُسقطت لأنها موجودة فعلاً
    skipped_sets = 0      # عدّاد للأطقم المتجاهَلة

    for t in result.get('testers_to_add', []):
        t_name = str(t.get('name', '') or '')

        # 🛑 لا تُضِف تستراً لطقم/مجموعة (إن مرّ من Gemini بالخطأ)
        if is_set(t_name):
            skipped_sets += 1
            continue

        # 🛡️ حماية نهائية: تخطّ التساتر التي تطابق منتجاً موجوداً (Gemini ربما
        # اقترحت تستراً لمنتج له تستر مماثل في القائمة الأصلية).
        t_sk = _normalize_perfume_name(_strip_tester_keyword(t_name))
        if t_sk and t_sk in existing_skeletons:
            # نتخطى فقط إذا كان التستر *الأصلي* موجوداً — لا نريد أن نفقد
            # التساتر الجديدة لعطور موجودة بدون تساتر. لذا نتحقق أن المنتج
            # الموجود فعلاً تستر، وليس عطراً أساسياً.
            existing_is_tester = False
            if name_col and name_col in original_df.columns:
                for raw in original_df[name_col].dropna().astype(str):
                    raw_sk = _normalize_perfume_name(_strip_tester_keyword(raw))
                    if raw_sk == t_sk and is_tester(raw):
                        existing_is_tester = True
                        break
            if existing_is_tester:
                skipped_existing += 1
                continue

        nr = {c: '' for c in all_cols}
        base_r = get_safe_row(t.get('base_product_id'))

        if name_col:  nr[name_col]  = _clean_cell(t.get('name', ''))
        if price_col: nr[price_col] = _clean_cell(t.get('new_price', 0)) or 0
        # 📝 نظّف الـ HTML من الفراغات الزائدة لمنع أسطر فارغة في عرض سلة
        if desc_col:
            raw_desc = _clean_cell(t.get('new_description', ''))
            # إذا كان الوصف مازال يحوي placeholders (لم يُملأ بعد)، املأ ما يمكن برمجياً
            desc_str = str(raw_desc) if raw_desc else ''
            if desc_str and ('[اسم الماركة]' in desc_str or '[اسم العطر]' in desc_str):
                # الفولباك: استبدل placeholders البسيطة على الأقل
                bp_cat_for_fill = ''
                if base_r is not None and cat_col:
                    bp_cat_for_fill = str(_clean_cell(base_r.get(cat_col, '')))
                base_name = t.get('_base_perfume_name') or t.get('name', '')
                desc_str = fill_tester_template_basics(
                    desc_str,
                    brand_name=brand_name,
                    perfume_name=base_name,
                    size_ml=int(t.get('size_ml') or 100),
                    base_category=bp_cat_for_fill,
                )
            nr[desc_col] = minify_html(desc_str) if desc_str else ''
        if brand_col: nr[brand_col] = brand_name

        # 🏷️ تصنيف التستر: يُحدَّد دائماً من تصنيفات التستر المخصصة (وليس
        # تصنيف العطر الأساسي). نستخدم تصنيف العطر الأساسي فقط كمدخل
        # لتحديد ما إذا كان رجالي/نسائي/نيش.
        if cat_col:
            base_cat_val = ''
            if base_r is not None:
                base_cat_val = str(_clean_cell(base_r.get(cat_col, '')))
            if not base_cat_val:
                # خذ من _base_category إن وُجد (من شبكة الأمان)
                base_cat_val = str(t.get('_base_category', '') or '')
            t_name = t.get('name', '') or t.get('_base_perfume_name', '')
            nr[cat_col] = map_to_tester_category(base_cat_val, t_name)

        if img_col:
            # 🖼️ احتفظ بقائمة الصور الكاملة مفصولة بفواصل (Salla يدعم متعددة)
            img = _clean_cell(t.get('image_url', ''))
            if not img and base_r is not None:
                raw_img = _clean_cell(base_r.get(img_col, ''))
                if raw_img:
                    # نظّف القائمة من الفراغات والقيم الفارغة
                    img_list = [u.strip() for u in str(raw_img).split(',') if u.strip()]
                    img = ','.join(img_list)
            nr[img_col] = img

        nr = apply_salla_mandatory(nr)
        rows.append(pd.Series(nr))

    for m in result.get('missing_products', []):
        m_name = str(m.get('name', '') or '')

        # 🛑 لا تُضِف طقماً/مجموعة من missing_products
        if is_set(m_name):
            skipped_sets += 1
            continue

        # 🛡️ حماية نهائية: تخطّ المنتجات الناقصة المطابقة لمنتج موجود فعلاً.
        m_sk = _normalize_perfume_name(_strip_tester_keyword(m_name))
        if m_sk and m_sk in existing_skeletons:
            skipped_existing += 1
            continue

        nr = {c: '' for c in all_cols}

        if name_col:  nr[name_col]  = _clean_cell(m.get('name', ''))
        if price_col: nr[price_col] = _clean_cell(m.get('price', 0)) or 0
        if desc_col:
            raw_desc = _clean_cell(m.get('description', ''))
            nr[desc_col] = minify_html(str(raw_desc)) if raw_desc else ''
        if brand_col: nr[brand_col] = _clean_cell(m.get('brand', '')) or brand_name

        cat_val = _clean_cell(m.get('category', ''))
        if cat_col:
            if cat_val:
                nr[cat_col] = _clean_category(cat_val)
            elif not original_df[cat_col].dropna().empty:
                nr[cat_col] = _clean_category(original_df[cat_col].dropna().mode().iloc[0])
            else:
                nr[cat_col] = 'العطور'

        if img_col:
            img1 = _clean_cell(m.get('image_url_1', ''))
            img2 = _clean_cell(m.get('image_url_2', ''))
            imgs = [u for u in (img1, img2) if u]
            nr[img_col] = ','.join(str(u) for u in imgs)

        nr = apply_salla_mandatory(nr)
        rows.append(pd.Series(nr))

    output_df = pd.DataFrame(rows) if rows else pd.DataFrame(columns=all_cols)
    output_df = output_df.fillna('')
    for null_token in _NULL_LIKE:
        output_df = output_df.replace(null_token, '')
    # pandas ≥ 2.1 أزال DataFrame.applymap لصالح DataFrame.map
    if hasattr(output_df, 'map') and callable(getattr(pd.DataFrame, 'map', None)):
        try:
            output_df = output_df.map(_clean_cell)
        except (TypeError, AttributeError):
            output_df = output_df.applymap(_clean_cell)
    else:
        output_df = output_df.applymap(_clean_cell)

    wb = load_workbook(io.BytesIO(template_bytes))

    active_title = wb.active.title
    for sheet_name in list(wb.sheetnames):
        if sheet_name != active_title:
            wb.remove(wb[sheet_name])

    ws = wb.active

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

    semantic_for_template = {
        'تصنيف المنتج': 'category', 'فئة المنتج': 'category', 'فئة': 'category',
        'التصنيف': 'category',
        'اسم المنتج': 'name', 'الاسم': 'name',
        'سعر المنتج': 'price', 'السعر': 'price', 'السعر العادي': 'price',
        'الوصف': 'description', 'وصف المنتج': 'description',
        'صورة المنتج': 'images', 'الصور': 'images', 'صور المنتج': 'images',
        'الماركة': 'brand', 'العلامة التجارية': 'brand',
    }
    semantic_to_dfcol = {
        'category': cat_col, 'name': name_col, 'price': price_col,
        'description': desc_col, 'images': img_col, 'brand': brand_col,
    }

    col_map: dict = {}
    for t_idx, t_hdr in enumerate(template_headers):
        if not t_hdr:
            continue
        t_str = str(t_hdr).strip()
        t_norm = _norm_hdr(t_str)
        matched = None
        for df_col in output_df.columns:
            if t_str == str(df_col).strip() or t_norm == _norm_hdr(str(df_col)):
                matched = df_col
                break
        if matched is None:
            sem = semantic_for_template.get(t_str) or semantic_for_template.get(t_norm)
            if sem and sem in semantic_to_dfcol and semantic_to_dfcol[sem] in output_df.columns:
                matched = semantic_to_dfcol[sem]
        if matched is not None:
            col_map[t_idx + 1] = matched

    direct_template_values: dict = {}
    for t_idx, t_hdr in enumerate(template_headers):
        if not t_hdr or (t_idx + 1) in col_map:
            continue
        cs = str(t_hdr).strip()
        ns = _norm_hdr(cs)
        matched_val = None
        for k, v in SALLA_MANDATORY.items():
            if cs == k or ns == _norm_hdr(k):
                matched_val = v
                break
        if matched_val is not None:
            direct_template_values[t_idx + 1] = matched_val
        elif 'نوع المنتج' in cs:
            direct_template_values[t_idx + 1] = 'منتج جاهز'
        elif cs == 'النوع':
            direct_template_values[t_idx + 1] = 'منتج'
        elif 'يتطلب شحن' in cs:
            direct_template_values[t_idx + 1] = 'نعم'
        # ⚠️ مهم: «إخفاء خيار تحديد الكمية» يحتوي كلمة «الكمية» — لذا يجب
        # فحصه قبل قاعدة الكمية حتى لا يأخذ القيمة 10 ويرفضه سلة.
        elif 'إخفاء خيار' in cs or 'اخفاء خيار' in ns:
            direct_template_values[t_idx + 1] = 'لا'
        elif 'اقصي كميه' in ns or 'اقصى كميه' in ns:
            direct_template_values[t_idx + 1] = 2
        elif ('الكمية' in cs or 'الكميه' in ns) and 'اقصي' not in ns and 'اقصى' not in ns:
            direct_template_values[t_idx + 1] = 10
        elif cs == 'الوزن':
            direct_template_values[t_idx + 1] = 1
        elif 'وحدة الوزن' in cs or 'وحده الوزن' in ns:
            direct_template_values[t_idx + 1] = 'kg'
        elif 'الماركة' in cs:
            direct_template_values[t_idx + 1] = brand_name

    last_written = data_start - 1
    for r_idx, (_, row) in enumerate(output_df.iterrows()):
        excel_row = data_start + r_idx
        last_written = excel_row
        for t_col, df_col in col_map.items():
            val = _clean_cell(row.get(df_col, ''))
            ws.cell(row=excel_row, column=t_col, value=val)
        for t_col, val in direct_template_values.items():
            ws.cell(row=excel_row, column=t_col, value=val)

    WIPE_UNTIL = max(ws.max_row + 1, last_written + 200)
    for r in range(last_written + 1, WIPE_UNTIL):
        for c in range(1, ws.max_column + 1):
            ws.cell(row=r, column=c, value=None)

    for r in range(data_start, last_written + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            if cell.value is None:
                continue
            sv = str(cell.value).strip()
            if sv in _NULL_LIKE:
                cell.value = ''

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ─── INIT STATE ──────────────────────────────────────────────────────────────

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

    # ⭐ زر اختبار المفتاح — يكشف خطأ 403 قبل بدء المعالجة الطويلة
    if api_key:
        if st.button("🔬 اختبار المفتاح", use_container_width=True):
            with st.spinner("جاري الاختبار..."):
                ok, msg = test_gemini_key(api_key, st.session_state.get('model_name', 'gemini-2.5-flash'))
                if ok:
                    st.success(msg)
                else:
                    st.error(msg)

    # 📋 قائمة نماذج Gemini (محدّثة آخر تحديث: مايو 2026 من توثيق Google الرسمي)
    # https://ai.google.dev/gemini-api/docs/models
    #
    # 🌟 الموصى به افتراضياً: gemini-2.5-flash
    #    - مستقر ومضمون (ليس Preview)
    #    - متاح في الـ Free Tier (10 RPM, 250 RPD)
    #    - متوازن: 90% من جودة Pro بسعر/سرعة Flash
    #
    # ⚡ للسرعة والحجم العالي: gemini-2.5-flash-lite (15 RPM, 1000 RPD)
    # 🧠 للدقة العالية والبحث المعقد: gemini-2.5-pro (5 RPM, 100 RPD)
    # 🚀 الأحدث (Preview، يتطلب Tier 1 المدفوع): gemini-3-flash, gemini-3.1-pro-preview
    model_name = st.selectbox(
        "🤖 النموذج",
        [
            # ─── Stable / Production-ready (موصى بها) ───
            'gemini-2.5-flash',            # ⭐ الأفضل توازناً (الافتراضي)
            'gemini-2.5-flash-lite',       # أسرع وأرخص لكن أقل دقة
            'gemini-2.5-pro',              # أعلى دقة، حصة أقل (5 RPM)
            # ─── Latest aliases (تتحدّث تلقائياً) ───
            'gemini-flash-latest',         # alias لأحدث Flash مستقر
            'gemini-pro-latest',           # alias لأحدث Pro مستقر
            # ─── Preview (يتطلب Cloud Billing) ───
            'gemini-3-flash',              # أحدث جيل، سرعة عالية
            'gemini-3.1-pro-preview',      # أدق نموذج (Preview)
            'gemini-3-flash-lite',         # الأرخص في الجيل الجديد
        ],
        index=0,  # 2.5-flash هو الافتراضي (الأكثر استقراراً)
        help=(
            "⭐ 2.5-flash: الموصى به (متوازن، مستقر، Free Tier)\n"
            "⚡ 2.5-flash-lite: أسرع، حصة أكبر (1000 RPD)\n"
            "🧠 2.5-pro: أعلى دقة، حصة أقل (100 RPD)\n"
            "🚀 3-flash / 3.1-pro: أحدث جيل (يتطلب Cloud Billing)"
        ),
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
st.caption("سياسة جديدة: تستر إلزامي لكل عطر بلا تستر · بحث عن سعر المنافس · سدّ الفجوات")

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
        column_config={'متوسط السعر': st.column_config.NumberColumn(format="%.0f ريال")}
    )

st.divider()

if current_idx >= total_brands:
    st.balloons()
    st.success("🎉 تمت معالجة جميع الماركات بنجاح!")
    results_json = json.dumps(st.session_state.brand_results, ensure_ascii=False, indent=2)
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

if st.session_state.waiting_confirm and st.session_state.current_result:
    result = st.session_state.current_result
    n_testers = len(result.get('testers_to_add', []))
    n_orphans = len(result.get('orphan_testers', []))
    n_missing = len(result.get('missing_products', []))

    # عدّ التساتر التي وُجدت عند منافس مقابل التي أُضيفت تلقائياً
    testers_list = result.get('testers_to_add', []) or []
    n_with_competitor = sum(1 for t in testers_list if t.get('tester_available_in_market'))
    n_auto_added = sum(1 for t in testers_list if t.get('_auto_added'))

    st.success(
        f"✅ اكتملت معالجة **{current_brand}** | "
        f"{n_testers} تستر جديد ({n_with_competitor} موجود عند منافس · "
        f"{n_auto_added} تلقائي بلا منافس) · "
        f"{n_orphans} تستر بلا أساسي · {n_missing} منتج ناقص"
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
            show_cols = [c for c in [
                'name', 'size_ml', 'original_price', 'new_price',
                'competitor_price', 'tester_available_in_market', 'source_store'
            ] if c in df_t.columns]
            st.dataframe(
                df_t[show_cols] if show_cols else df_t,
                use_container_width=True,
                hide_index=True,
                column_config={
                    'name': 'اسم التستر',
                    'size_ml': st.column_config.NumberColumn('الحجم (مل)', format="%d"),
                    'original_price': st.column_config.NumberColumn('السعر الأصلي', format="%.0f ريال"),
                    'new_price': st.column_config.NumberColumn('سعرنا للتستر', format="%.0f ريال"),
                    'competitor_price': st.column_config.NumberColumn('سعر المنافس (مرجعي)', format="%.0f ريال"),
                    'tester_available_in_market': st.column_config.CheckboxColumn('متوفر عند منافس؟'),
                    'source_store': 'المتجر المرجعي',
                }
            )
            st.caption(
                f"ℹ️ {n_with_competitor} من التساتر وُجدت عند منافس سعودي (سعرها مرجعي فقط) · "
                f"{n_auto_added} أُضيفت تلقائياً بسياسة المتجر بدون منافس مرجعي."
            )
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

if not st.session_state.processing:
    if not st.session_state.api_key:
        st.warning("⚠️ أدخل Gemini API Key في الشريط الجانبي أولاً")
        st.stop()

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

brand_lbl  = st.empty()
brand_bar  = st.progress(0)
prod_lbl   = st.empty()
prod_bar   = st.progress(0)
status_msg = st.empty()

writing_dna = extract_writing_dna(df)

brand_lbl.markdown("**الخطوة 1/3:** جاري تحليل أسلوب الكتابة وتجهيز بيانات المنتجات...")
brand_bar.progress(5)

img_col_now = find_col(df, 'images')
cat_col_now = find_col(df, 'category')
brand_col_now = get_brand_col(df)
products_payload = []
if name_col:
    for _, row in brand_df.iterrows():
        raw_img = str(row.get(img_col_now, '') or '') if img_col_now else ''
        # احتفظ بكل الصور كقائمة مفصولة بفواصل (Salla يدعم متعددة)
        img_list = [u.strip() for u in raw_img.split(',') if u.strip()]
        full_img = ','.join(img_list)

        cat_val = str(row.get(cat_col_now, '') or '') if cat_col_now else ''
        brand_val = str(row.get(brand_col_now, '') or '') if brand_col_now else ''

        products_payload.append({
            'id': str(row.get('No.', row.name)),
            'name': str(row.get(name_col, '')),
            'price': float(pd.to_numeric(row.get(price_col, 0), errors='coerce') or 0),
            'description': '',
            'image_url': full_img,  # ← قائمة كاملة الآن (وليس الصورة الأولى فقط)
            'category': cat_val,
            'brand': brand_val,
            'is_tester': is_tester(str(row.get(name_col, ''))),
        })

n = len(products_payload)

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
    """يُشغّل دفعة واحدة. البحث (Grounding) محفوظ دائماً —
    call_gemini_brand تُعيد المحاولة داخلياً عند الأخطاء العابرة."""
    with status_lock:
        batch_status[b_idx]['state'] = 'running'
        batch_status[b_idx]['started_at'] = time.time()
        batch_status[b_idx]['mode'] = 'grounding'

    def cb(n_chars):
        with status_lock:
            batch_status[b_idx]['chars'] = n_chars

    return call_gemini_brand(
        brand_name=current_brand,
        products=batches[b_idx],
        full_brand_products=products_payload,
        api_key=_api_key_val,
        writing_dna=writing_dna,
        model_name=_model_name_val,
        use_grounding=True,
        batch_index=b_idx,
        total_batches=total_batches,
        progress_cb=cb,
    )


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
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


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

    # ⭐ شبكة الأمان النهائية — تضمن وجود تستر لكل عطر أساسي بلا تستر
    # حتى لو فاتت Gemini أحدها، نُضيفه تلقائياً بسعر داخلي.
    accumulated = ensure_all_testers_added(accumulated, products_payload)

    # احفظ مرة أخرى بعد التطبيق التلقائي
    with open(autosave_path, 'w', encoding='utf-8') as f:
        json.dump(accumulated, f, ensure_ascii=False, indent=2)

    # 🌟 إثراء التساتر التلقائية — استدعاء Gemini لتعبئة النوتات والتفاصيل
    # المتخصصة (افتتاحية، قلب، قاعدة، عائلة عطرية، أسئلة شائعة) لكل تستر.
    auto_count = sum(
        1 for t in (accumulated.get('testers_to_add') or [])
        if isinstance(t, dict) and t.get('_auto_added')
    )
    if auto_count > 0:
        brand_lbl.markdown(
            f"**الخطوة 3/3:** إثراء أوصاف {auto_count} تستر تلقائي بمعلومات حقيقية..."
        )
        brand_bar.progress(80)

        def _enrich_progress(done, total):
            try:
                pct = 80 + int(15 * done / max(total, 1))
                brand_bar.progress(min(pct, 95))
                prod_lbl.markdown(f"📝 {done}/{total} وصف مُعبّأ بالمعلومات الحقيقية")
            except Exception:
                pass

        try:
            accumulated = enrich_auto_added_testers(
                accumulated,
                brand_name=current_brand,
                api_key=api_key,
                model_name=model_name,
                batch_size=8,
                progress_cb=_enrich_progress,
            )
            with open(autosave_path, 'w', encoding='utf-8') as f:
                json.dump(accumulated, f, ensure_ascii=False, indent=2)
        except Exception as enrich_err:
            # فشل الإثراء — نستمر بالقالب المُملوء برمجياً
            st.warning(
                f"⚠️ تعذّر إثراء التساتر التلقائية بمعلومات إضافية "
                f"({enrich_err}). سيتم استخدام الوصف الأساسي فقط."
            )

    result = {k: v for k, v in accumulated.items() if not k.startswith('_')}
    # احتفظ بمعلومة الإضافات التلقائية للعرض
    if accumulated.get('_auto_added_testers'):
        result['_auto_added_testers'] = accumulated['_auto_added_testers']

    brand_bar.progress(75)
    prod_bar.progress(0.7)
    brand_lbl.markdown("**الخطوة 3/3:** معالجة النتائج وتطبيق سياسة التستر الإلزامي...")

    n_tst = len(result.get('testers_to_add', []))
    n_orph = len(result.get('orphan_testers', []))
    n_mis = len(result.get('missing_products', []))
    n_auto = len(result.get('_auto_added_testers', []) or [])

    brand_bar.progress(100)
    prod_bar.progress(1.0)
    brand_lbl.markdown(f"✅ **اكتملت معالجة {current_brand}!**")
    prod_lbl.markdown(
        f"✅ {n_tst} تستر جديد (منهم {n_auto} مُضافين تلقائياً) · "
        f"{n_orph} يتيم · {n_mis} ناقص"
    )
    status_msg.success(
        f"🎯 اكتملت المعالجة — {n_tst} تستر جديد "
        f"({n_auto} تلقائي بسياسة المتجر) | "
        f"{n_orph} تستر بلا أساسي | {n_mis} منتج ناقص"
    )

    st.session_state.current_result = result
    st.session_state.waiting_confirm = True
    st.session_state.processing = False
    time.sleep(0.5)
    st.rerun()

except Exception as e:
    err = str(e)
    brand_bar.progress(0)

    err_low = err.lower()

    if ('403' in err_low and 'denied access' in err_low) or 'permission_denied' in err_low:
        status_msg.error("❌ **المشروع محجوب من Google** (403 PERMISSION_DENIED)")
        with st.expander("🔍 كيف أحل هذه المشكلة؟", expanded=True):
            st.markdown("""
            **هذا الخطأ على مستوى حساب Google — ليس في التطبيق.**

            **الأسباب الشائعة:**
            - المفتاح تسرّب وحُجب تلقائياً (مثلاً ظهر على GitHub)
            - المشروع محجوب على مستوى Google Cloud
            - منطقتك تتطلب تفعيل Cloud Billing

            **الحل (جرّب بالترتيب):**
            1. افتح [Google AI Studio API Keys](https://aistudio.google.com/apikey)
            2. **احذف** المفتاح الحالي تماماً
            3. اضغط **Create API Key** → **Create in new project**
            4. انسخ المفتاح الجديد إلى الشريط الجانبي
            5. اضغط 🔬 اختبار المفتاح للتأكد قبل المعالجة

            إن استمرّ الخطأ بعد المفتاح الجديد:
            - فعّل Billing على المشروع من [Google Cloud Console](https://console.cloud.google.com/billing) (مجاني ضمن Free Tier)
            - تأكد ألا تُضيف المفتاح لأي ملف يُرفع إلى GitHub
            """)
        st.session_state.processing = False

    elif 'leaked' in err_low or 'reported as leaked' in err_low:
        status_msg.error("❌ **المفتاح مُبلَّغ عنه كمُسرَّب** — Google حجبته تلقائياً")
        st.info("أنشئ مفتاحاً جديداً من https://aistudio.google.com/apikey")
        st.session_state.processing = False

    elif 'api_key' in err_low or 'api key' in err_low or 'unauthenticated' in err_low or '401' in err_low:
        status_msg.error("❌ **مفتاح Gemini غير صحيح** — تحقق من نسخه كاملاً بدون فراغات")
        st.session_state.processing = False

    elif 'quota' in err_low or 'rate' in err_low or '429' in err or 'resource_exhausted' in err_low:
        status_msg.error("❌ **تجاوزت حد الاستخدام** — انتظر دقيقة أو جرّب موديل gemini-2.5-flash-lite")
        st.session_state.processing = False

    elif 'safety' in err_low or 'حُجبت' in err:
        status_msg.error("❌ **حُجب الرد بفلتر الأمان** — راجع المنتجات لكلمات حساسة")
        st.session_state.processing = False

    else:
        status_msg.error(f"❌ خطأ غير متوقع: {err}")
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
