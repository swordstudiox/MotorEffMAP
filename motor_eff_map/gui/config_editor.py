import configparser
import os

from PySide6.QtWidgets import QComboBox, QLineEdit


class ConfigEditorMixin:
    def parse_ini_file(self, path):
        if not os.path.exists(path):
            return configparser.ConfigParser(), {}

        content = ""
        encoding_used = "utf-8-sig"
        for encoding in ("utf-8-sig", "gb18030"):
            try:
                with open(path, "r", encoding=encoding) as f:
                    content = f.read()
                encoding_used = encoding
                break
            except UnicodeDecodeError:
                continue
            except Exception as e:
                raise Exception(f"读取配置文件失败。({e})")
        else:
            raise Exception("读取配置文件失败。未知编码，请保存为 UTF-8 或 GB18030。")

        self.current_encoding = encoding_used

        has_section = False
        for line in content.splitlines():
            line = line.strip()
            if line.startswith("[") and line.endswith("]"):
                has_section = True
                break
            if line and not line.startswith("#") and not line.startswith(";"):
                break

        if not has_section:
            content = "[DEFAULT]\n" + content

        parser = configparser.ConfigParser()
        parser.optionxform = str
        parser.read_string(content)

        flat_dict = {}
        if "DEFAULT" in parser:
            for key, val in parser["DEFAULT"].items():
                val_clean = val.split(";")[0].split("#")[0].strip()
                flat_dict[key] = val_clean

        for section in parser.sections():
            for key, val in parser.items(section):
                val_clean = val.split(";")[0].split("#")[0].strip()
                flat_dict[key] = val_clean

        self.ensure_default_config_values(parser, flat_dict)
        return parser, flat_dict

    def ensure_default_config_values(self, parser, flat_dict):
        default_section = parser["DEFAULT"]
        for key, default_value in self.DEFAULT_CONFIG_VALUES.items():
            if key not in flat_dict:
                flat_dict[key] = default_value
                default_section[key] = default_value

    def get_config_section_keys(self, parser, section):
        if section == "DEFAULT":
            return parser.defaults().keys()
        return parser._sections.get(section, {}).keys()

    def get_config_display_label(self, key):
        description = self.CONFIG_LABELS.get(key)
        if not description:
            return f"{key}:"
        return f"{key}（{description}）:"

    def create_config_editor(self, key, value):
        if key in self.SWITCH_CONFIG_KEYS:
            combo = QComboBox()
            combo.addItem("开启", "1")
            combo.addItem("关闭", "0")
            combo.setCurrentIndex(0 if str(value).strip() == "1" else 1)
            combo.setStyleSheet("""
                QComboBox {
                    font-family: 'Microsoft YaHei', 'Segoe UI', sans-serif;
                    font-size: 10pt;
                    padding: 4px;
                    border: 1px solid #aaa;
                    border-radius: 3px;
                    background-color: white;
                }
                QComboBox:focus {
                    border: 1px solid #3399ff;
                }
            """)
            return combo

        edit = QLineEdit(value)
        edit.setStyleSheet("""
            QLineEdit {
                font-family: 'Microsoft YaHei', 'Segoe UI', sans-serif;
                font-size: 10pt;
                padding: 4px;
                border: 1px solid #aaa;
                border-radius: 3px;
                background-color: white;
            }
            QLineEdit:focus {
                border: 1px solid #3399ff;
            }
        """)
        return edit

    def get_config_editor_value(self, editor):
        if isinstance(editor, QComboBox):
            return editor.currentData()
        return editor.text()

    def write_ini_file(self, file_path, data):
        try:
            with open(file_path, "r", encoding=self.current_encoding) as f:
                lines = f.readlines()
        except FileNotFoundError:
            lines = []

        new_lines = []
        written_keys = set()

        for line in lines:
            stripped = line.strip()
            if not stripped or stripped.startswith(";") or stripped.startswith("#") or stripped.startswith("["):
                new_lines.append(line)
                continue

            if "=" in line:
                parts = line.split("=", 1)
                key = parts[0].strip()

                if key in data:
                    new_val = data[key]

                    comment_part = ""
                    if "#" in parts[1]:
                        comment_part = " #" + parts[1].split("#", 1)[1]
                    elif ";" in parts[1]:
                        comment_part = " ;" + parts[1].split(";", 1)[1]

                    value_text = f" {new_val}" if str(new_val) else ""
                    new_lines.append(f"{key} ={value_text}{comment_part}\n")
                    written_keys.add(key)
                else:
                    new_lines.append(line)
            else:
                new_lines.append(line)

        missing_keys = [key for key in data if key not in written_keys]
        if missing_keys:
            if new_lines and new_lines[-1].strip():
                new_lines.append("\n")
            new_lines.append("# 由程序补充的新增配置项\n")
            for key in missing_keys:
                new_val = data[key]
                value_text = f" {new_val}" if str(new_val) else ""
                new_lines.append(f"{key} ={value_text}\n")

        self.current_encoding = "utf-8-sig"
        with open(file_path, "w", encoding=self.current_encoding, newline="\n") as f:
            f.writelines(new_lines)

