from .base import BaseRule

class ScaviRule(BaseRule):
    def clean_site_name(self, site_text):
        s = str(site_text).strip().lower()
        s_no_space = s.replace(" ", "")
        if "vpc" in s_no_space:
            return "vpc"
        if "k9" in s_no_space:
            return "k9"
        
        # Remove prefix
        for prefix in ["scavi - ", "scavi-"]:
            if s.startswith(prefix):
                s = s[len(prefix):]
        return s
