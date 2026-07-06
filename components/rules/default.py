from ..common import clean_ingredient_name, SHORT_NAMES, normalize_comp
from .base import BaseRule
from .lixil import LixilRule
from .om import OmRule
from .zamil import ZamilRule
from .scavi import ScaviRule

_CUSTOM_RULES = {
    "lixil": LixilRule(),
    "om": OmRule(),
    "osv": OmRule(),
    "odsv": OmRule(),
    "zamil": ZamilRule(),
    "scavi": ScaviRule(),
}

def get_rule_for_customer(customer_name):
    key = normalize_comp(customer_name)
    return _CUSTOM_RULES.get(key, BaseRule())

def get_shopping_short_name(nvl_name, customer_name):
    # Try customer specific rule first
    rule = get_rule_for_customer(customer_name)
    custom_val = rule.get_ingredient_short_name(nvl_name)
    if custom_val:
        return custom_val

    # Generic short name rules
    n = str(nvl_name).lower().strip()
    n = clean_ingredient_name(n)

    if "củ cải" in n:
        return "củ cải"
    if "má gà" in n or "thịt gà" in n:
        return "gà"
    if "tôm" in n:
        return "tôm"
    if "cá sống" in n:
        return "cá sống"
    if "mọc" in n:
        return "mọc"
    if "xay" in n or "bằm" in n:
        return "thịt xay"
    if "đậu hũ chiên" in n:
        return "đậu hũ chiên"
    if "đậu hũ trắng" in n:
        return "đậu hũ trắng"
    if "nấm mèo" in n:
        return "nấm mèo"
    if "miến" in n:
        return "miến"
    if "đậu xanh" in n:
        return "đậu xanh"
    if "bào ngư" in n:
        return "nấm bào ngư"
    if "bông cải" in n:
        return "bông cải"
    if "vịt" in n:
        return "vịt"
    if "khoai môn" in n:
        return "khoai môn"
    if "chả cá" in n:
        return "chả cá"
    if "xương heo" in n or "xương hầm" in n:
        return "xương hầm"
    if "kim chi" in n:
        return "kim chi"
    if "hoành thánh" in n:
        return "hoành thánh"
    if "bò gàu" in n:
        return "bò gàu"
    if "chả lụa" in n:
        return "chả lụa"
    if "chả bò" in n:
        return "chả bò"
    if "cà chua" in n:
        return "cà chua"
    if "mực nang" in n or "mực" in n:
        return "mực"
    if "bò tái" in n:
        return "bò tái"
    if "bò nạm" in n:
        return "bò nạm"
    if "bò viên" in n:
        return "bò viên"

    return SHORT_NAMES.get(n, n)

def clean_site_name(site_text, customer_name):
    rule = get_rule_for_customer(customer_name)
    custom_val = rule.clean_site_name(site_text)
    if custom_val is not None:
        return custom_val

    # Generic site cleaning
    s = str(site_text).strip().lower()
    s_no_space = s.replace(" ", "")
    if "fab1" in s_no_space:
        return "fab1"
    if "fab2" in s_no_space:
        return "fab2"
    if "vpc" in s_no_space:
        return "vpc"
    if "til" in s_no_space:
        return "til"

    # Remove customer name prefix
    comp_key = normalize_comp(customer_name)
    if comp_key:
        for prefix in [f"{comp_key} - ", f"{comp_key}-", f"{customer_name.lower()} - ", f"{customer_name.lower()}-"]:
            if s.startswith(prefix):
                s = s[len(prefix):]
    return s
