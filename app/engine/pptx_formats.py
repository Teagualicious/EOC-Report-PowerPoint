"""Value formatting for PowerPoint insertion: dates, text case, numbers.

These helpers make inserted values visually match the placeholder text
they replace (date style detection, case matching, CamelCase splitting).
"""

import json
import logging
import os
import re
from datetime import datetime

from config.paths import DICT_PATH

logger = logging.getLogger(__name__)


def _load_date_formats():
    """Load the date_formats section of the metric dictionary."""
    if os.path.exists(DICT_PATH):
        try:
            with open(DICT_PATH, "r") as f:
                return json.load(f).get("date_formats", {})
        except (OSError, json.JSONDecodeError):
            logger.warning("Could not read date formats from %s", DICT_PATH,
                           exc_info=True)
    return {}


def _detect_date_format(text, date_formats=None):
    """Detect date format style from existing text. Returns style name or None.

    ``date_formats`` may be injected (tests); defaults to the dictionary file.
    """
    clean = text.replace("[", "").replace("]", "").strip()

    if date_formats is None:
        date_formats = _load_date_formats()
    for style, info in date_formats.items():
        for pattern in info.get("regex_patterns", []):
            # Try both original text (with brackets) and cleaned text
            if re.search(pattern, text.strip(), re.IGNORECASE) or \
               re.search(pattern, clean, re.IGNORECASE):
                return style

    # Fallback hardcoded patterns
    patterns = [
        (r'(?:Jan(?:uary)?|Feb(?:ruary|uary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+\d{1,2}(?:st|nd|rd|th)?,?\s+\d{4}',
         "long_ordinal"),
        (r'\d{1,2}(?:st|nd|rd|th),?\s+\d{4}', "long_ordinal"),
        (r'\d{2}/\d{2}/\d{4}', "us_slash"),
        (r'\d{1,2}/\d{1,2}/\d{4}', "us_slash_short"),
        (r'\d{4}-\d{2}-\d{2}', "iso"),
        (r'\d{2}-\d{2}-\d{4}', "us_dash"),
    ]

    for pattern, style in patterns:
        if re.search(pattern, clean, re.IGNORECASE):
            return style
    return None


def _format_date(date_str, target_style):
    """Convert a date string (YYYY-MM-DD) to the target style."""
    try:
        dt = datetime.strptime(date_str.strip(), "%Y-%m-%d")
    except (ValueError, AttributeError):
        return date_str  # Can't parse, return as-is

    if target_style == "long_ordinal":
        day = dt.day
        suffix = "th" if 11 <= day <= 13 else {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
        return dt.strftime(f"%B {day}{suffix}, %Y")
    elif target_style == "short":
        return dt.strftime(f"%b {dt.day}, %Y")
    elif target_style == "short_month":
        return f"{dt.strftime('%b')} {dt.day}, {dt.year}"
    elif target_style == "us_slash":
        return dt.strftime("%m/%d/%Y")
    elif target_style == "us_slash_short":
        return f"{dt.month}/{dt.day}/{dt.year}"
    elif target_style == "iso":
        return dt.strftime("%Y-%m-%d")
    elif target_style == "us_dash":
        return dt.strftime("%m-%d-%Y")
    elif target_style == "month_only":
        return dt.strftime("%B")
    elif target_style == "month_year":
        return dt.strftime("%B %Y")
    elif target_style == "short_month_year":
        return dt.strftime("%b %Y")
    elif target_style == "long_mdy":
        day = dt.day
        suffix = "th" if 11 <= day <= 13 else {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
        return f"{dt.strftime('%B')} {day}{suffix}, {dt.year}"
    return date_str


def _split_camel_case(text):
    """Split CamelCase into separate words. 'AcmeAppliance' → 'Acme Appliance'."""
    # Insert space before uppercase letters that follow lowercase
    result = re.sub(r'([a-z])([A-Z])', r'\1 \2', text)
    # Insert space before uppercase letters that are followed by lowercase (for acronyms)
    result = re.sub(r'([A-Z]+)([A-Z][a-z])', r'\1 \2', result)
    return result


def _match_text_style(value, existing_text):
    """Match the text style (case, spacing) of the existing text."""
    if not existing_text or not isinstance(value, str) or not value:
        return value

    # Split CamelCase if the value has no spaces but should
    if " " not in value and any(c.isupper() for c in value[1:]):
        value = _split_camel_case(value)

    # Match case style of existing text
    existing_clean = existing_text.strip()
    if not existing_clean:
        return value

    if existing_clean.isupper():
        return value.upper()
    elif existing_clean.islower():
        return value.lower()
    elif existing_clean == existing_clean.title():
        return value.title()
    # If existing is mixed case, keep the value as-is
    return value


def _clean_float_noise(value):
    """Collapse pandas .sum() float noise (1.0000000000000009 → 1)."""
    if isinstance(value, float) and abs(value - round(value)) < 0.0001:
        return int(round(value))
    return value


def _coerce_number(value):
    """Convert numpy scalars (int64, float64, ...) to native Python types.

    pandas aggregations return numpy scalars; numpy.int64 does NOT subclass
    Python int, so every ``isinstance(v, (int, float))`` check silently
    fails and values fall through to raw-string formatting. This was the
    root cause of format settings not applying anywhere.
    """
    if isinstance(value, (int, float, str, bool)) or value is None:
        return value
    item = getattr(value, "item", None)
    if callable(item):
        try:
            return value.item()
        except Exception:
            return value
    return value


DATE_STYLE_CHOICES = [
    ("long_ordinal", "May 1st, 2026"),
    ("us_slash", "05/01/2026"),
    ("us_slash_short", "5/1/2026"),
    ("us_dash", "05-01-2026"),
    ("iso", "2026-05-01"),
    ("short_month", "May 1, 2026"),
    ("month_year", "May 2026"),
    ("month_only", "May"),
    ("custom", "Custom (strftime)"),
]


def format_date_with_style(value, style, custom_fmt=None):
    """Format an ISO date string ('YYYY-MM-DD') or range
    ('YYYY-MM-DD - YYYY-MM-DD') using a named style or a custom strftime
    pattern. Non-date values pass through unchanged."""
    if not isinstance(value, str):
        return str(value)

    def _one(date_str):
        if style == "custom" and custom_fmt:
            try:
                dt = datetime.strptime(date_str.strip(), "%Y-%m-%d")
                return dt.strftime(custom_fmt)
            except (ValueError, AttributeError):
                return date_str
        return _format_date(date_str, style)

    m = re.match(r'(\d{4}-\d{2}-\d{2})\s*[-–—]\s*(\d{4}-\d{2}-\d{2})', value.strip())
    if m:
        return f"{_one(m.group(1))} – {_one(m.group(2))}"
    if re.match(r'^\d{4}-\d{2}-\d{2}$', value.strip()):
        return _one(value.strip())
    return value


def format_with_details(value, details):
    """Format a numeric value using detailed settings from the format popup.

    ``details``: {"format": str, "decimals": int, "commas": bool,
                  "prefix": str, "suffix": str}
    Falls back to sensible default formatting when details are missing or
    the value is not numeric.
    """
    value = _coerce_number(value)

    # Date formatting applies to string values — handle before the
    # numeric-only path
    if details and details.get("format") == "date":
        return format_date_with_style(value, details.get("date_style", "iso"),
                                      details.get("custom_strftime"))

    if not isinstance(value, (int, float)) or isinstance(value, bool):
        return str(value)

    if not details:
        value = _clean_float_noise(value)
        if isinstance(value, float) and value == int(value):
            return f"{int(value):,}"
        return f"{value:,}" if isinstance(value, int) else f"{value:,.2f}"

    fmt = details.get("format", "text")
    dec = details.get("decimals", 0)
    use_commas = details.get("commas", True)
    pre = details.get("prefix", "")
    suf = details.get("suffix", "")

    if fmt == "percentage":
        return f"{value:.{dec}%}" if value < 1 else f"{value:.{dec}f}%"
    elif fmt == "currency":
        fmt_str = f",.{dec}f" if use_commas else f".{dec}f"
        return f"${value:{fmt_str}}"
    elif fmt == "number":
        fmt_str = f",.{dec}f" if use_commas else f".{dec}f"
        return f"{value:{fmt_str}}"
    elif fmt == "custom":
        fmt_str = f",.{dec}f" if use_commas else f".{dec}f"
        return f"{pre}{value:{fmt_str}}{suf}"
    else:
        value = _clean_float_noise(value)
        if isinstance(value, float) and value == int(value):
            return f"{int(value):,}"
        # Non-whole floats: round to 2dp — never leak full float repr
        return f"{value:,}" if isinstance(value, int) else f"{value:,.2f}"


def _format_value(value, fmt, existing_text=None):
    """Format a value for display. Auto-matches date format, text case, and number style."""
    # numpy scalars (from pandas aggregation) fail isinstance(int/float)
    # checks — convert to native Python first
    value = _coerce_number(value)
    # Check if value looks like a date (YYYY-MM-DD or YYYY-MM-DD - YYYY-MM-DD)
    if isinstance(value, str) and existing_text:
        # Date range: "2026-05-01 - 2026-05-31"
        range_match = re.match(r'(\d{4}-\d{2}-\d{2})\s*-\s*(\d{4}-\d{2}-\d{2})', value)
        if range_match:
            target_style = _detect_date_format(existing_text)
            if target_style:
                d1 = _format_date(range_match.group(1), target_style)
                d2 = _format_date(range_match.group(2), target_style)
                if " – " in existing_text or " — " in existing_text:
                    sep = " – " if " – " in existing_text else " — "
                elif " - " in existing_text:
                    sep = " - "
                elif " to " in existing_text.lower():
                    sep = " to "
                else:
                    sep = " – "
                return f"{d1}{sep}{d2}"

        # Single date: "2026-05-01"
        single_match = re.match(r'^\d{4}-\d{2}-\d{2}$', value.strip())
        if single_match:
            target_style = _detect_date_format(existing_text)
            if target_style:
                return _format_date(value.strip(), target_style)

        # Text formatting — match case and spacing of existing text
        return _match_text_style(value, existing_text)

    if isinstance(value, (int, float)):
        import numpy as np
        value = _clean_float_noise(value)
        # Treat whole floats as integers (356590.0 → 356,590)
        is_whole = isinstance(value, float) and value == int(value) and not np.isnan(value)
        if fmt == "currency":
            if is_whole:
                return f"${int(value):,}"
            return f"${value:,.2f}" if isinstance(value, float) else f"${value:,}"
        elif fmt == "percentage":
            return f"{value:.1%}" if value < 1 else f"{value:.1f}%"
        elif fmt == "number":
            if is_whole:
                return f"{int(value):,}"
            return f"{value:,}" if isinstance(value, int) else f"{value:,.2f}"
        else:
            # Default "text" format — still format numbers nicely
            if is_whole:
                return f"{int(value):,}"
            # Non-whole floats: 2dp — never leak full float repr
            return f"{value:,}" if isinstance(value, int) else f"{value:,.2f}"
    return str(value)
