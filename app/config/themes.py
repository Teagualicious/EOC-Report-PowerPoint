"""UI color themes for the desktop application."""

THEMES = {
    "light": {
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
    },
    "dark": {
        "bg": "#091827", "fg": "#EAF1F8", "card": "#10263B",
        "card_fg": "#EAF1F8", "input_bg": "#0C2134", "input_fg": "#EAF1F8",
        "muted": "#91A8BC", "accent": "#3793F4", "danger": "#E05A55",
        "success": "#45A778", "secondary": "#18344D", "secondary_fg": "#D6E7F5",
        "list_bg": "#0C2134", "border": "#24435D", "insert": "#FFFFFF",
        "highlight": "#163B5B", "warning": "#D99535",
        "pill_warn_bg": "#3A2B10", "pill_warn_fg": "#F2BC68",
        "brand": "#061421", "brand_fg": "#FFFFFF", "brand_muted": "#9DB9D1",
        "surface_alt": "#0C2032", "hover": "#4BA0F6", "focus": "#285E8A",
        "pill_ok_bg": "#153B2A", "pill_ok_fg": "#72D2A1",
        "pill_info_bg": "#123451", "pill_info_fg": "#79B9F8",
    },
}


def get_theme(name):
    """Return the theme dict for a name, defaulting to light."""
    return THEMES.get(name, THEMES["light"])
