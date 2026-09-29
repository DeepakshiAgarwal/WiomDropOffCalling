"""
Voice-of-Customer (VOC) analysis: categorizes the free-text comment column
('What EXACTLY are they waiting for (their words)', CALL SHEET column V) for
the top reason codes, mirroring the "X. <Reason> - What is the customer VOC?"
tables already hand-built in 'DAILY SUMMARY -APP Drop offs ' rows 151-217.

Unlike the rest of the dashboard, this is NOT a deterministic spreadsheet
formula being replicated -- the original table's sub-categories were produced
by a human reading the raw comments. This module reproduces the same
sub-category taxonomy with keyword/regex rules (first matching rule wins, in
priority order) so it can be refreshed mechanically from fresh data. It is a
best-effort text classifier, not an exact match to any formula -- treat the
counts as approximate, and re-tune the rules in REASON_RULES below if a
category's proportions drift a lot from what a manual read would show.
"""
import re

# Column letters (openpyxl 1-based via column_index_from_string)
PRIMARY_REASON_COL = "T"
COMMENT_COL = "V"  # "What EXACTLY are they waiting for (their words)"

# Order matters: for each reason, rules are tried top-to-bottom, first match wins.
REASON_RULES = {
    "R25 Non-serviceable area": [
        ("Non-serviceable shown in app/phone", re.compile(r"\b(app|phone)\b.*\b(show|shown|dikha)", re.I)),
        ("Non-serviceable shown in app/phone", re.compile(r"\bshow(ing)?\b.*\b(app|phone)\b", re.I)),
        ("Area / location not serviceable", re.compile(r".*")),  # catch-all
    ],
    "R23 Confusion Regarding the Booking Process and Plans": [
        ("Booking already completed", re.compile(r"booking (kar liya|complete|completed)|completed the booking", re.I)),
        ("Not interested / no requirement", re.compile(r"interest(ed)? nahi|nahi chahte|not interested", re.I)),
        ("Temporary constraint / callback", re.compile(r"\b(busy|callback|call ?back|out of station|office)\b", re.I)),
        ("Location / serviceability clarification", re.compile(r"location|serviceable|showing as serviceable", re.I)),
        ("Booking intent - date/time committed",
         re.compile(r"\b(aaj|kal|shaam|evening|morning|tarik|tareekh|date|monday|tuesday|wednesday|thursday|friday|saturday|sunday|\d+\s*(din|day|week|month|hafte|mahine|minute|min))\b.*\bbooking\b", re.I)),
        ("Booking intent - date/time committed",
         re.compile(r"\bbooking\b.*\b(aaj|kal|shaam|evening|tarik|tareekh|\d+\s*(din|day|week|month|hafte|mahine|minute|min))\b", re.I)),
        ("Plan / recharge / charges information", re.compile(r"recharge|charges|amount|plan|jankari|information|security", re.I)),
        ("Booking process / guidance", re.compile(r"guide|process|samajh nahi|confusion", re.I)),
        ("Other / unclear / decision pending", re.compile(r".*")),
    ],
    "R01 Purana Connection Chal Raha hai": [
        ("Waiting for current recharge / plan to end", re.compile(r"recharge (khatam|khtam|khatm)|recharge.{0,15}(baad|end)", re.I)),
        ("Not interested / just checking", re.compile(r"check kar|need nahi|interest(ed)? nahi|zarurat nahi", re.I)),
        ("Existing competitor / other ISP active",
         re.compile(r"airtel|jio|excitel|local (company|brand)|another (brand|company)|other (brand|company|isp)|different local", re.I)),
        ("Price / plan / feature concern", re.compile(r"speed|price|plan|ott|channel|feature", re.I)),
        ("Will book later / undecided", re.compile(r"next (month|week)|baad me|later|month|week|sochunga|aage", re.I)),
        ("Already a Wiom user / additional connection", re.compile(r"already.{0,10}wiom|dobara.{0,10}wiom", re.I)),
        ("Booking / setup / process issue", re.compile(r"setup|technician|process issue|booking nahi karna", re.I)),
        ("Other / unclear", re.compile(r".*")),
    ],
    "R09 Bas dekh rahe the, abhi jaldi nahi": [
        ("App downloaded / visited by mistake", re.compile(r"mistake|galti se|advertisement|advatisment|by mistake", re.I)),
        ("Comparing plans / providers", re.compile(r"(2|3|do|teen) (company|companies)|comparing|dusri company.{0,10}dekh", re.I)),
        ("No current requirement / not interested", re.compile(r"interest(ed)? nahi|jarurat nahi|need nahi|nahi lagwana", re.I)),
        ("Will book later / future need",
         re.compile(r"\d+\s*(din|day|week|month|hafte|mahine|saal)\b.{0,15}(baad|later)|next (month|week)|month ke (end|start)|future", re.I)),
        ("Just checking / exploring", re.compile(r"check kar|dekh rah|bas.{0,10}(check|dekh)|sirf check", re.I)),
        ("Other / unclear", re.compile(r".*")),
    ],
    "R10 Payment failed - UPI/card/bank issue": [
        ("Waiting for refund / retry", re.compile(r"refund", re.I)),
        ("Payment / UPI / app technical issue",
         re.compile(r"\b(issue|failed|fail|dikat|dikkat|band ho gaya|server issue|scanner nahi)\b", re.I)),
        ("Does not use UPI / wants offline payment", re.compile(r"upi|online payment", re.I)),
        ("Plan / payment information needed", re.compile(r"jankari|information|charges|plan", re.I)),
        ("Will complete booking shortly",
         re.compile(r"\b(aaj|kal|abhi|\d+\s*(min|minute|din|day))\b.{0,15}booking|booking.{0,15}\b(aaj|kal|\d+\s*(min|minute|din|day))\b", re.I)),
        ("Other / unclear", re.compile(r".*")),
    ],
    "R24 Others": [
        ("Cancellation / refund issue", re.compile(r"cancel|refund", re.I)),
        ("Not interested / no requirement", re.compile(r"not interested|interest(ed)? nahi|nahi chahta|booking nahi kar(ni|wani)|connection nahi lena", re.I)),
        ("Booking intent / will book",
         re.compile(r"\b(aaj|kal|shaam|evening|tarik|tareekh|ready)\b.{0,20}booking|booking.{0,20}\b(aaj|kal|shaam|tarik|tareekh|ready)\b", re.I)),
        ("Booking / app / process issue", re.compile(r"\botp\b|app (download|delete)|process|guide|issue", re.I)),
        ("Plan / product / service query", re.compile(r"recharge|plan|charges|speed|security|smart tv|ott", re.I)),
        ("Temporary constraint / unavailable", re.compile(r"busy|not.{0,5}location|office|available nahi|out of station", re.I)),
        ("Other / unclear", re.compile(r".*")),
    ],
}

DISPLAY_TITLE = {
    "R25 Non-serviceable area": "Non-serviceable area",
    "R23 Confusion Regarding the Booking Process and Plans": "Confusion regarding booking process",
    "R01 Purana Connection Chal Raha hai": "Purana connection chal raha hai",
    "R09 Bas dekh rahe the, abhi jaldi nahi": "Just exploring, no immediate requirement",
    "R10 Payment failed - UPI/card/bank issue": "Payment failed - UPI/card/bank issue",
    "R24 Others": "Others",
}

MAX_SAMPLES = 3
MAX_SAMPLE_LEN = 140


def _clean_sample(text):
    t = re.sub(r"\s+", " ", str(text)).strip()
    if len(t) > MAX_SAMPLE_LEN:
        t = t[:MAX_SAMPLE_LEN].rsplit(" ", 1)[0] + "…"
    return t


def classify_voc(df, col):
    """df: the CALL SHEET dataframe. col(letter) -> Series helper from build_dashboard.py."""
    T = col(df, PRIMARY_REASON_COL).astype(str).str.strip()
    V = col(df, COMMENT_COL)

    groups = []
    for reason_label, rules in REASON_RULES.items():
        mask = (T == reason_label) & V.notna()
        texts = V[mask].astype(str).str.strip()
        texts = texts[texts != ""]
        total = len(texts)
        if total == 0:
            continue

        buckets = {}
        for t in texts:
            for cat_name, pattern in rules:
                if pattern.search(t):
                    buckets.setdefault(cat_name, []).append(t)
                    break

        cats = []
        for cat_name, items in buckets.items():
            samples = [_clean_sample(s) for s in items[:MAX_SAMPLES]]
            cats.append(dict(
                name=cat_name,
                count=len(items),
                pct=round(100 * len(items) / total, 1),
                samples=samples,
            ))
        cats.sort(key=lambda c: -c["count"])

        top = cats[0]
        takeaway = f"\"{top['name']}\" is the single largest theme, at {top['pct']}% of {total} comments read for this reason."

        groups.append(dict(
            reason_label=reason_label,
            title=DISPLAY_TITLE.get(reason_label, reason_label),
            total=total,
            categories=cats,
            takeaway=takeaway,
        ))

    return groups
