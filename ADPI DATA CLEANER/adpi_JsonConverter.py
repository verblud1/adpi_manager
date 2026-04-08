import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import json
import re
import os
import sys
from datetime import datetime

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class RegistryProDatabase(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Registry Pro Database Engine")
        self.geometry("740x600")
        ctk.set_appearance_mode("dark")
        
        self.label = ctk.CTkLabel(self, text="Конвертер АДПИ (v4.6 Auto-Header)", font=("Arial", 22, "bold"))
        self.label.pack(pady=20)

        self.btn = ctk.CTkButton(self, text="Выбрать файл и создать базу JSON", 
                                 command=self.process, height=50, width=400, corner_radius=10)
        self.btn.pack(pady=10)

        self.info_box = ctk.CTkTextbox(self, height=350, width=680, font=("Consolas", 12))
        self.info_box.pack(pady=20)
        self.info_box.insert("0.0", "Система готова. Ожидание файла...\n" + "-"*50 + "\n")

    def find_headers_and_map(self, df_raw):
        """ Ищет строку заголовков и сопоставляет колонки """
        patterns = {
            "id": ["№", "n", "id", "номер"],
            "name": ["фио", "резидент", "заявитель", "фамилия"],
            "children": ["детей", "ребен"],
            "address": ["адрес", "проживания", "установка"],
            "device_count": ["кол-во", "количество"],
            "model": ["марка", "модель"],
            "install_date": ["дата установки"],
            "check_date": ["дата проверки", "работоспособности"],
            "notes_tech": ["комментарии", "исправен", "неисправен"],
            "notes_social": ["соц", "обслуживании", "сопровождении", "о семье"]
        }

        # Проверяем первые 10 строк на наличие заголовков
        for i in range(min(10, len(df_raw))):
            row_values = [str(val).lower().strip() for val in df_raw.iloc[i].values]
            mapping = {}
            
            for key, keywords in patterns.items():
                for idx, cell_text in enumerate(row_values):
                    if any(k in cell_text for k in keywords):
                        mapping[key] = idx
                        break
            
            # Если нашли хотя бы ФИО и Адрес, значит это строка заголовков
            if "name" in mapping and "address" in mapping:
                return i, mapping
        return None, None

    def extract_dates(self, val):
        if pd.isna(val): return []
        if isinstance(val, (datetime, pd.Timestamp)): return [val.strftime("%d.%m.%Y")]
        return re.findall(r'(\d{2}\.\d{2}\.\d{4})', str(val))

    def parse_address_pro(self, addr_str):
        addr_str = str(addr_str).strip()
        data = {"region": "Вышневолоцкий городской округ", "settlement_type": "город", "settlement_name": "Вышний Волочек",
                "street": "н/д", "house": "", "apartment": "", "is_village": False, "raw_address": addr_str}
        
        v_pref = r'(?:пгт\.|пос\.|поселок|дер\.|д\.|с\.|село|с/п|снт|п\.)'
        v_match = re.search(v_pref + r'\s?([А-ЯЁ][а-яё\-]+(?:\s[А-ЯЁ][а-яё\-]+)?)', addr_str, re.IGNORECASE)
        known = r'(Есеновичи|Афимьино|Красномайский|Терелесовский|Белый Омут|Зеленогорский|Деревково|Красная Заря|Пригородный|Солнечный|Горняк)'
        spec = re.search(known, addr_str, re.IGNORECASE)

        if v_match or spec:
            name = (v_match.group(1) if v_match else spec.group(1)).strip()
            if name.lower() not in ["вышний волочек", "в.волочек"]:
                data["is_village"], data["settlement_name"], data["settlement_type"] = True, name, "пос./дер."

        h_match = re.search(r'д\.\s?(\d+[а-яА-ЯёЁ]? (?:/\s?\d+)?|\d+/\d+|\d+[а-яА-ЯёЁ]?)', addr_str)
        if h_match: data["house"] = h_match.group(1).replace(" ", "").lower().strip()

        st_match = re.search(r'(?:ул\.|пер\.|проезд|пр-т|наб\.|тупик)\s?([^,]+)', addr_str, re.IGNORECASE)
        if not st_match:
            fallback = r'(?<!г\.\s)(?<!город\s)(?<!Вышний\s)([А-ЯЁ][а-яё\-]+(?:\s[А-ЯЁ][а-яё\-]+)?)(?=\s?,?\s?д\.)'
            st_match = re.search(fallback, addr_str)
        if st_match:
            st = re.sub(r'^(ул\.|г\.|д\.)\s?', '', st_match.group(1 if st_match.lastindex else 0).strip(), flags=re.IGNORECASE)
            st = re.sub(r'[\s,]+д\.?$', '', st).strip()
            if st.lower() not in ["волочек", "вышний", "в.волочек"]: data["street"] = st

        apt_match = re.search(r'кв\.\s?(\d+)', addr_str)
        if apt_match: data["apartment"] = apt_match.group(1).strip()
        if data["is_village"] and data["street"] == "н/д": data["street"] = data["settlement_name"]
        return data

    def clean_numeric(self, val):
        if pd.isna(val): return 0
        if isinstance(val, (int, float)): return int(val)
        res = re.findall(r'\d+', str(val))
        return int(res[0]) if res else 0

    def analyze_status(self, row_values, resident_name=""):
        text_blob = " ".join([str(v) for v in row_values if pd.notna(v)]) + " " + str(resident_name)
        blob_lower = text_blob.lower()
        tags, status, move_data = [], "active", None
        m_match = re.search(r'(?:ул\.|улица)\s*([А-Яа-яЁё\s\d\-]+),\s*(?:д\.|дом)\s*([\d/]+[а-я]?)', text_blob, re.IGNORECASE)
        if m_match:
            move_data = {"street": m_match.group(1).strip(), "house": m_match.group(2).strip()}
            apt = re.search(r'кв\.\s?(\d+)', text_blob[m_match.end():], re.IGNORECASE)
            move_data["apartment"] = apt.group(1) if apt else ""
        if any(x in blob_lower for x in ["не прожив", "переехал", "нет на месте"]): tags.append("not_living_here"); status = "inactive"
        if "отказ" in blob_lower: tags.append("access_refused"); status = "refused"
        if "исправен" in blob_lower: tags.append("functional")
        if any(x in blob_lower for x in ["неисправен", "замен", "сломан"]): tags.append("maintenance_required")
        if any(x in blob_lower for x in ["утратила статус", "не многодетн"]): tags.append("status_lost"); status = "archived"
        return status, tags, move_data

    def deduplicate_database(self, raw_data):
        merged = {}
        for entry in raw_data:
            name_key = entry["resident_name"].strip().lower()
            if name_key not in merged:
                merged[name_key] = entry
            else:
                existing = merged[name_key]
                all_checks = existing["check_history"] + entry["check_history"]
                seen_dates = set()
                unique_checks = []
                for c in all_checks:
                    if c["date"] not in seen_dates:
                        unique_checks.append(c); seen_dates.add(c["date"])
                
                unique_checks.sort(key=lambda x: datetime.strptime(x["date"], "%d.%m.%Y"), reverse=True)
                existing["check_history"] = unique_checks

                latest_new = datetime.strptime(entry["check_history"][0]["date"], "%d.%m.%Y") if entry["check_history"] else datetime.min
                latest_ext = datetime.strptime(existing["check_history"][0]["date"], "%d.%m.%Y") if existing["check_history"] else datetime.min

                if latest_new >= latest_ext:
                    if entry["location"]["raw_address"] != existing["location"]["raw_address"]:
                        existing["location_history"].append({
                            "old_address": existing["location"]["raw_address"],
                            "date_archived": datetime.now().strftime("%d.%m.%Y"),
                            "reason": "deduplication_merge"
                        })
                    existing["location"] = entry["location"]
                    existing["status"] = entry["status"]
                    existing["device"] = entry["device"]
                
                existing["tags"] = list(set(existing["tags"] + entry["tags"]))
                merged[name_key] = existing
        return list(merged.values())

    def process(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not path: return
        try:
            # Читаем без заголовков (header=None), чтобы самим найти нужную строку
            df_raw = pd.read_excel(path, engine='openpyxl', header=None)
            
            header_row_idx, col_map = self.find_headers_and_map(df_raw)
            
            if header_row_idx is None:
                messagebox.showerror("Ошибка", "Не удалось найти строку с заголовками (ФИО, Адрес).")
                return

            # Данные начинаются со следующей строки после заголовков
            data_df = df_raw.iloc[header_row_idx + 1:]
            
            raw_database = []
            current_family = None

            for _, row in data_df.iterrows():
                id_idx = col_map.get("id", 0)
                val_id = str(row.iloc[id_idx]).strip().replace('.0','')
                
                if val_id.isdigit():
                    if current_family: raw_database.append(current_family)
                    
                    m_fields = [row.iloc[col_map[k]] for k in ["notes_tech", "notes_social"] if k in col_map]
                    stat, tags, move_info = self.analyze_status(m_fields, row.iloc[col_map["name"]])
                    
                    current_family = {
                        "id": int(val_id),
                        "resident_name": str(row.iloc[col_map["name"]]).strip(),
                        "children_count": self.clean_numeric(row.iloc[col_map["children"]]) if "children" in col_map else 0,
                        "location": self.parse_address_pro(row.iloc[col_map["address"]]),
                        "location_history": [],
                        "device": {
                            "model": str(row.iloc[col_map["model"]]).strip() if "model" in col_map else "АДПИ",
                            "count": self.clean_numeric(row.iloc[col_map["device_count"]]) if "device_count" in col_map else 0,
                            "install_date": self.extract_dates(row.iloc[col_map["install_date"]])[0] if "install_date" in col_map and self.extract_dates(row.iloc[col_map["install_date"]]) else ""
                        },
                        "check_history": [{"date": d} for d in self.extract_dates(row.iloc[col_map["check_date"]])] if "check_date" in col_map else [],
                        "status": stat, "tags": tags,
                        "notes": {
                            "social": str(row.iloc[col_map["notes_social"]]).strip() if "notes_social" in col_map else "",
                            "tech": str(row.iloc[col_map["notes_tech"]]).strip() if "notes_tech" in col_map else "",
                            "misc": ""
                        }
                    }
                elif current_family:
                    for val in row.values:
                        if pd.notna(val):
                            for d in self.extract_dates(val):
                                if d not in [h["date"] for h in current_family["check_history"]]:
                                    current_family["check_history"].append({"date": d})
            
            if current_family: raw_database.append(current_family)
            final_database = self.deduplicate_database(raw_database)

            output_path = os.path.splitext(path)[0] + "_FINAL.json"
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(final_database, f, ensure_ascii=False, indent=4)

            self.info_box.insert("end", f"Успех!\nНайдено семей: {len(final_database)}\n")
            messagebox.showinfo("Успех", f"Обработка завершена.")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Критическая ошибка: {str(e)}")

if __name__ == "__main__":
    app = RegistryProDatabase()
    app.mainloop()