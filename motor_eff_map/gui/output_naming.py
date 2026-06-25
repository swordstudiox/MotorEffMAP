import os
import re


class OutputNamingMixin:
    def sanitize_filename_component(self, value):
        text = str(value or "").strip()
        text = re.sub(r"[^\w\u4e00-\u9fff.-]+", "_", text)
        text = re.sub(r"_+", "_", text).strip("._ ")
        return text or "未命名"

    def build_output_stem(self, suffix=""):
        source_file = self.current_results.get("file", "") if hasattr(self, "current_results") else ""
        source_stem = os.path.splitext(os.path.basename(source_file))[0]
        sheet = self.current_results.get("sheet", "") if hasattr(self, "current_results") else ""
        veh_code = self.config_dict.get("VehicleCode", "")

        udc_val = "0"
        if self.logic is not None and self.logic.u_dc is not None:
            udc_val = str(int(round(self.logic.u_dc.mean())))

        direction = self.current_results.get("direction", "") if hasattr(self, "current_results") else ""
        state = self.current_results.get("state", "") if hasattr(self, "current_results") else ""
        if state == "电动":
            state = "驱动"

        parts = [
            source_stem,
            sheet,
            f"{veh_code}-{udc_val}V-{direction}{state}",
            suffix,
        ]
        return "_".join(self.sanitize_filename_component(p) for p in parts if p)

