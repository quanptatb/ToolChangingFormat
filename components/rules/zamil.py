from .base import BaseRule

class ZamilRule(BaseRule):
    def get_ingredient_short_name(self, nvl_name):
        n = str(nvl_name).lower().strip()
        if any(x in n for x in ["nạc vai", "heo vai", "vai tươi", "vai heo"]):
            return "vai"
        if "trứng" in n:
            return "trứng gà"
        return None
