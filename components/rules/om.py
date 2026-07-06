from .base import BaseRule


class OmRule(BaseRule):
    def get_loai_mon_val(self, r_dict):
        # OM/ODSV layout follows the meal structure rather than only Loai mon.
        return str(r_dict.get("co_cau") or "").strip()
