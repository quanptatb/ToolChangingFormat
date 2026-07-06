from .base import BaseRule

class OdsvRule(BaseRule):
    def get_loai_mon_val(self, r_dict):
        # ODSV displays LoaiMonAn based on meal data (co_cau)
        return str(r_dict.get("co_cau") or "").strip()
