"""Deck Engine's single analyst-facing light theme."""

LIGHT_THEME = {
    "bg": "#F3F6FA", "fg": "#172435", "card": "#FFFFFF",
    "card_fg": "#172435", "input_bg": "#FFFFFF", "input_fg": "#172435",
    "muted": "#66758A", "accent": "#0876E8", "danger": "#C53B35",
    "success": "#1F8A55", "secondary": "#E8EEF5", "secondary_fg": "#23405C",
    "list_bg": "#FFFFFF", "border": "#D8E1EC", "insert": "#172435",
    "highlight": "#EAF2FB", "warning": "#B86A00",
    "pill_warn_bg": "#FFF1D7", "pill_warn_fg": "#9A5700",
    "brand": "#003057", "brand_fg": "#FFFFFF", "brand_muted": "#BFD5EA",
    "surface_alt": "#F8FAFD", "hover": "#0768CC", "focus": "#B9D8FA",
    "pill_ok_bg": "#E1F5EA", "pill_ok_fg": "#176A43",
    "pill_info_bg": "#E5F0FF", "pill_info_fg": "#075EBC",
}

THEMES = {"light": LIGHT_THEME}


def get_theme(_name="light"):
    """Return the only supported theme; legacy dark settings are ignored."""
    return LIGHT_THEME
