class BaseRule:
    def get_ingredient_short_name(self, nvl_name):
        """Return custom short name for ingredient, or None to use default."""
        return None
        
    def clean_site_name(self, site_text):
        """Return custom cleaned site name, or None to use default."""
        return None
        
    def get_loai_mon_val(self, r_dict):
        """Return custom Loại món ăn value, or default."""
        return str(r_dict.get("loai_mon") or "").strip()
