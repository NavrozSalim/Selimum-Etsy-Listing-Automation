"""
Etsy SEO Tag Engine
Implements the 20 strict Etsy SEO Tag Generation rules.
Validates character limits (<= 20 chars), bans measurement/quantity terms,
diversifies SEO angles, and strictly enforces the last 2 tags as 'party decor', 'party supplies'.
"""

import re
from typing import List, Tuple, Dict, Any, Optional

BANNED_MEASUREMENT_PATTERNS = [
    r'\b\d+(\.\d+)?\s*(inch|inches|in|yd|yds|yard|yards|ft|feet|foot|cm|mm|m|meter|meters)\b',
    r'\b(inch|inches|yard|yards|feet|foot|centimeter|centimeters|cm|mm|meter|meters)\b',
    r'\b(size|sizes|dimension|dimensions)\b',
    r'\b\d+\s*(pc|pcs|piece|pieces|pk|pack|packs|ct|count|counts)\b',
    r'\b(piece|pieces|quantity|piece counts|package counts|counts)\b',
    r'\b\d+x\d+\b',
    r'\b\d+["\']\b',
]

COLOR_KEYWORDS = [
    "burgundy", "sage green", "emerald green", "forest green", "mint green", "olive green",
    "dusty blue", "navy blue", "royal blue", "baby blue", "sky blue", "teal", "turquoise",
    "blush pink", "dusty rose", "rose gold", "hot pink", "baby pink", "coral", "peach",
    "champagne", "ivory", "cream", "white", "black", "gold", "silver", "bronze", "copper",
    "lavender", "lilac", "purple", "plum", "mauve", "mustard yellow", "yellow", "orange",
    "terracotta", "rust", "burnt orange", "brown", "beige", "tan", "red", "wine red"
]

MATERIAL_KEYWORDS = [
    "satin", "velvet", "silk", "chiffon", "organza", "burlap", "linen", "cotton",
    "grosgrain", "lace", "tulle", "glitter", "metallic", "foil", "latex", "vinyl",
    "acrylic", "wooden", "wood", "kraft paper", "cardstock", "sequin", "wire", "canvas"
]

OCCASION_MAPPINGS = {
    "christmas": "christmas party",
    "halloween": "halloween party",
    "birthday": "birthday party",
    "valentines": "valentines party",
    "valentine": "valentines party",
    "fathers day": "fathers day party",
    "mothers day": "mothers day party",
    "new year": "new year party",
    "graduation": "graduation party",
    "bachelorette": "bachelorette party",
}

EVENT_DECOR_MAPPINGS = {
    "wedding": "wedding decor",
    "baby shower": "baby shower decor",
    "bridal shower": "bridal shower decor",
    "engagement": "engagement decor",
    "graduation": "graduation decor",
    "anniversary": "anniversary decor",
    "nursery": "nursery decor",
    "reception": "wedding decor",
}

RIBBON_ANGLES = [
    "bow making", "gift wrapping", "floral arrangement", "bouquet ribbon",
    "craft ribbon", "wreath ribbon", "tree ribbon", "hair bow ribbon",
    "sewing trim", "diy crafts"
]

BALLOON_ANGLES = [
    "balloon garland", "balloon arch", "party balloons", "backdrop decor",
    "photo backdrop", "birthday balloons", "balloon bouquet", "event balloons"
]

BANNER_ANGLES = [
    "party banner", "garland banner", "backdrop banner", "party garland",
    "wall hanging", "photo booth banner", "celebration banner"
]


def clean_text(text: str) -> str:
    """Lowercases and cleans extra whitespace."""
    return re.sub(r'\s+', ' ', text.lower()).strip()


def contains_banned_terms(tag: str) -> bool:
    """Checks if tag contains measurement or quantity keywords."""
    tag_lower = clean_text(tag)
    for pattern in BANNED_MEASUREMENT_PATTERNS:
        if re.search(pattern, tag_lower):
            return True
    return False


def sanitize_tag(tag: str) -> str:
    """Cleans a single tag, removes punctuation, trims to 20 chars."""
    tag = clean_text(tag)
    # Remove special punctuation except spaces and letters
    tag = re.sub(r'[^a-zA-Z0-9\s-]', '', tag)
    tag = re.sub(r'\s+', ' ', tag).strip()
    if len(tag) > 20:
        # Try trimming words from the end
        words = tag.split()
        while words and len(" ".join(words)) > 20:
            words.pop()
        tag = " ".join(words)
    return tag[:20].strip()


def validate_tags(tags: List[str]) -> Tuple[bool, List[str]]:
    """
    Validates a list of tags against the 20 rules.
    Returns (is_valid, list_of_error_reasons).
    """
    errors = []
    if len(tags) != 13:
        errors.append(f"Expected exactly 13 tags, got {len(tags)}")

    for i, t in enumerate(tags, start=1):
        if len(t) > 20:
            errors.append(f"Tag #{i} '{t}' exceeds 20 characters ({len(t)} chars)")
        if contains_banned_terms(t):
            errors.append(f"Tag #{i} '{t}' contains prohibited measurement or quantity keyword")

    # Check last 2 tags
    if len(tags) >= 13:
        if tags[-2].lower().strip() != "party decor":
            errors.append(f"Tag #12 must be 'party decor', got '{tags[-2]}'")
        if tags[-1].lower().strip() != "party supplies":
            errors.append(f"Tag #13 must be 'party supplies', got '{tags[-1]}'")

    # Check duplicates
    lower_tags = [t.lower().strip() for t in tags]
    if len(lower_tags) != len(set(lower_tags)):
        errors.append("Tags contain duplicates or repetitive entries")

    return (len(errors) == 0, errors)


def generate_tags_rule_based(title: str, description: str = "") -> List[str]:
    """
    Algorithmic SEO tag generator strictly adhering to the 20 rules.
    Generates 11 distinct high-value angles + 2 fixed tags ('party decor', 'party supplies').
    """
    text = clean_text(f"{title} {description}")

    # Remove all measurement/quantity numbers from search consideration
    cleaned_title = title
    for pat in BANNED_MEASUREMENT_PATTERNS:
        cleaned_title = re.sub(pat, " ", cleaned_title, flags=re.IGNORECASE)
    cleaned_text = clean_text(cleaned_title)

    candidates: List[str] = []
    seen = set()

    def add_candidate(cand: str):
        cand = sanitize_tag(cand)
        if not cand or len(cand) > 20 or len(cand) < 3:
            return
        if cand in seen or cand in ["party decor", "party supplies"]:
            return
        if contains_banned_terms(cand):
            return
        seen.add(cand)
        candidates.append(cand)

    # 1. Detect Category / Core Product Type
    is_ribbon = bool(re.search(r'\b(ribbon|trim|satin ribbon|velvet ribbon|grosgrain)\b', text))
    is_balloon = bool(re.search(r'\b(balloon|balloons|garland arch|arch kit)\b', text))
    is_banner = bool(re.search(r'\b(banner|garland|bunting|backdrop|pennant)\b', text))

    # Extract detected colors
    detected_colors = []
    for c in COLOR_KEYWORDS:
        if re.search(r'\b' + re.escape(c) + r'\b', text):
            detected_colors.append(c)

    # Extract detected materials
    detected_materials = []
    for m in MATERIAL_KEYWORDS:
        if re.search(r'\b' + re.escape(m) + r'\b', text):
            detected_materials.append(m)

    # Extract detected events / occasions
    detected_events = []
    for evt, tag in EVENT_DECOR_MAPPINGS.items():
        if re.search(r'\b' + re.escape(evt) + r'\b', text):
            detected_events.append(tag)

    detected_occasions = []
    for occ, tag in OCCASION_MAPPINGS.items():
        if re.search(r'\b' + re.escape(occ) + r'\b', text):
            detected_occasions.append(tag)

    primary_color = detected_colors[0] if detected_colors else None
    secondary_color = detected_colors[1] if len(detected_colors) > 1 else None
    primary_material = detected_materials[0] if detected_materials else None

    # --- Build Targeted Angles ---
    if is_ribbon:
        # Angle 1: Color + Ribbon or Material + Ribbon
        if primary_color and primary_material:
            cand = f"{primary_color} {primary_material}"
            if len(cand) <= 20:
                add_candidate(cand)
        if primary_color:
            cand = f"{primary_color} ribbon"
            if len(cand) <= 20:
                add_candidate(cand)
        if primary_material:
            cand = f"{primary_material} ribbon"
            if len(cand) <= 20:
                add_candidate(cand)

        # Angle 2: Craft uses (bow making, gift wrapping, floral wrap)
        add_candidate("bow making")
        add_candidate("gift wrapping")
        add_candidate("craft ribbon")
        add_candidate("floral arrangement")
        add_candidate("bouquet ribbon")
        add_candidate("diy craft supply")
        add_candidate("wreath supplies")

    elif is_balloon:
        if primary_color:
            cand = f"{primary_color} balloons"
            if len(cand) <= 20:
                add_candidate(cand)
        add_candidate("balloon garland")
        add_candidate("balloon arch")
        add_candidate("party balloons")
        add_candidate("backdrop decor")
        add_candidate("photo backdrop")
        add_candidate("balloon bouquet")
        add_candidate("event balloons")

    elif is_banner:
        if primary_color:
            cand = f"{primary_color} banner"
            if len(cand) <= 20:
                add_candidate(cand)
        add_candidate("party banner")
        add_candidate("garland banner")
        add_candidate("backdrop banner")
        add_candidate("party garland")
        add_candidate("photo booth banner")
        add_candidate("celebration banner")

    # Add Event / Decor Angles
    for evt_tag in detected_events:
        add_candidate(evt_tag)

    for occ_tag in detected_occasions:
        add_candidate(occ_tag)

    # General phrases extracted from title (filtering out banned numbers)
    # Extract noun chunks / 2-3 word combinations
    clean_words = [w for w in re.findall(r'[a-zA-Z]+', cleaned_text) if len(w) > 2]
    for i in range(len(clean_words) - 1):
        bigram = f"{clean_words[i]} {clean_words[i+1]}"
        if len(bigram) <= 20:
            add_candidate(bigram)
        if i < len(clean_words) - 2:
            trigram = f"{clean_words[i]} {clean_words[i+1]} {clean_words[i+2]}"
            if len(trigram) <= 20:
                add_candidate(trigram)

    # Default versatile high-performing tags if we still need more to reach 11
    backup_pool = [
        "event decoration", "celebration decor", "handmade supply",
        "craft supplies", "diy party", "holiday decor", "tablescape decor",
        "home celebration", "rustic decor", "boho decor", "modern party"
    ]

    for b in backup_pool:
        if len(candidates) >= 11:
            break
        add_candidate(b)

    # Pick exactly 11 distinct candidates
    first_11 = candidates[:11]
    while len(first_11) < 11:
        # Fallback filler
        for b in ["festive decor", "gathering decor", "diy decorations", "ceremony decor"]:
            if b not in first_11 and len(first_11) < 11:
                first_11.append(b)

    # Append strictly required tags #12 and #13
    final_tags = first_11 + ["party decor", "party supplies"]

    return final_tags


def generate_tags_with_openai(title: str, api_key: str, model: str = "gpt-4o-mini") -> Optional[List[str]]:
    """
    Generates tags using OpenAI API following the 20 rules,
    then rigorously validates and sanitizes them.
    """
    try:
        from openai import OpenAI
        client = OpenAI(api_key=api_key)

        prompt = f"""Generate Etsy SEO tags for the product title I provide.

Follow ALL of these rules strictly:

1. Generate exactly 13 Etsy tags.
2. Put all 13 tags in ONE code block only.
3. Separate every tag with a comma.
4. The first 11 tags must be highly relevant, SEO-friendly product tags.
5. The last 2 tags MUST always be exactly:
   party decor
   party supplies
6. Every tag must be 20 characters or fewer. Count characters carefully, including spaces.
7. NEVER use measurements or quantity-related keywords in tags. Do not include:
   - inches / inch
   - yards / yard
   - feet / foot
   - centimeters / cm
   - size
   - dimensions
   - quantity or piece counts
   - package counts
8. Do not simply copy the product title into tags. Select the strongest and most useful Etsy search phrases from the product.
9. Avoid repetitive tags. The first 11 tags should use different relevant SEO angles whenever possible, such as:
   - color
   - material
   - product type
   - style
   - theme
   - application
   - decoration use
   - craft use
   - event
   - occasion
   - audience/use case
10. Keep all tags directly relevant to the product. Do not add unrelated keywords just because they may have high search volume.
11. For ribbon products, consider different angles such as: ribbon color, ribbon type, material, bow making, gift wrapping, floral arrangements, bouquet use, craft projects, wedding/event use, decorative use.
12. For balloon products, consider: balloon color, balloon type, balloon garland, balloon arch, backdrop, party theme, birthday/event, decoration use, specific occasion.
13. For banners and garlands, consider: banner type, garland type, theme, event, occasion, decoration style, party use.
14. For occasions, use "party" when it naturally fits (e.g., christmas party, halloween party, birthday party, valentines party).
15. For events, use "decor" when appropriate (e.g., wedding decor, baby shower decor, bridal shower decor, engagement decor).
16. Do not use vague or unnecessarily short occasion tags when a stronger phrase is possible.
17. Do not make the first 11 tags too similar or repetitive.
18. Make sure EVERY tag is 20 characters or fewer.
19. Do not add explanations, descriptions, suggestions, headings, or promotional text.
20. Output ONLY the final 13 comma-separated Etsy tags inside ONE code block.

Product title:
{title}
"""

        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": "You are an expert Etsy SEO specialist following strict tag rules."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.3,
        )

        content = response.choices[0].message.content.strip()
        # Extract from code block if present
        match = re.search(r'```(?:[a-zA-Z]*)\n?(.*?)\n?```', content, re.DOTALL)
        raw_text = match.group(1).strip() if match else content.strip()

        # Parse comma separated tags
        parsed_tags = [sanitize_tag(t) for t in raw_text.split(",") if t.strip()]

        # Clean and ensure 13 tags with correct ending
        cleaned_tags = []
        for t in parsed_tags:
            if not contains_banned_terms(t) and len(t) <= 20 and t not in cleaned_tags:
                cleaned_tags.append(t)

        # Enforce last two
        if "party decor" in cleaned_tags:
            cleaned_tags.remove("party decor")
        if "party supplies" in cleaned_tags:
            cleaned_tags.remove("party supplies")

        first_11 = cleaned_tags[:11]
        if len(first_11) < 11:
            # Supplement from rule-based engine if needed
            rule_fallback = generate_tags_rule_based(title)
            for fb in rule_fallback[:11]:
                if fb not in first_11 and len(first_11) < 11:
                    first_11.append(fb)

        final_tags = first_11 + ["party decor", "party supplies"]
        return final_tags
    except Exception as e:
        print(f"[TagEngine] OpenAI tag generation failed: {e}. Falling back to rule-based engine.")
        return None


def generate_etsy_tags(title: str, description: str = "", api_key: Optional[str] = None) -> List[str]:
    """
    Unified entry point. Tries AI if API key is present, otherwise uses rule-based generator.
    Guarantees exactly 13 tags, <= 20 chars, no banned terms, ending in 'party decor', 'party supplies'.
    """
    if api_key:
        tags = generate_tags_with_openai(title, api_key)
        if tags and len(tags) == 13:
            return tags

    # Rule-based fallback
    return generate_tags_rule_based(title, description)


def format_tags_output(tags: List[str]) -> str:
    """Formats tags strictly as ONE code block with comma separation."""
    return f"```\n{', '.join(tags)}\n```"
