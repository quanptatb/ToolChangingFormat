from .base import BaseRule

class LixilRule(BaseRule):
    def get_ingredient_short_name(self, nvl_name):
        n = str(nvl_name).lower().strip()
        if any(x in n for x in ["nạc vai", "heo vai", "vai tươi", "vai heo"]):
            return "vai"
        return None

    def clean_site_name(self, site_text):
        s = str(site_text).strip().lower()
        s_no_space = s.replace(" ", "")
        if "fab1" in s_no_space:
            return "FAB1"
        if "fab2" in s_no_space:
            return "FAB2"
        if "vpc" in s_no_space:
            return "VPC"
        if "til" in s_no_space:
            return "TIL"
        return s
