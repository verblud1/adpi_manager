import customtkinter as ctk
from tkinter import filedialog, messagebox
import json
import os
from datetime import datetime

class DeepJSONMerger(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Registry Pro | Глубокая Синхронизация")
        self.geometry("900x700")
        ctk.set_appearance_mode("dark")

        self.old_json_path = None
        self.new_json_path = None

        # UI
        self.label = ctk.CTkLabel(self, text="Полная синхронизация полей JSON", font=("Arial", 22, "bold"))
        self.label.pack(pady=20)

        self.frame = ctk.CTkFrame(self)
        self.frame.pack(pady=10, padx=20, fill="x")

        self.btn_old = ctk.CTkButton(self.frame, text="1. Старая база (Архив)", command=self.load_old, fg_color="#4A4A4A")
        self.btn_old.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

        self.btn_new = ctk.CTkButton(self.frame, text="2. Новый файл (Приоритет)", command=self.load_new, fg_color="#1f538d")
        self.btn_new.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        self.frame.grid_columnconfigure((0, 1), weight=1)

        self.process_btn = ctk.CTkButton(self, text="НАЧАТЬ ПОЛНОЕ СРАВНЕНИЕ", command=self.sync_logic, 
                                         height=50, state="disabled", fg_color="#28a745")
        self.process_btn.pack(pady=20)

        self.info_box = ctk.CTkTextbox(self, height=400, width=860, font=("Consolas", 12))
        self.info_box.pack(pady=10, padx=20)

    def load_old(self):
        path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if path:
            self.old_json_path = path
            self.log(f"[OK] Архив загружен: {os.path.basename(path)}")
            self.check_ready()

    def load_new(self):
        path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if path:
            self.new_json_path = path
            self.log(f"[OK] Актуальный файл загружен: {os.path.basename(path)}")
            self.check_ready()

    def check_ready(self):
        if self.old_json_path and self.new_json_path:
            self.process_btn.configure(state="normal")

    def log(self, text):
        self.info_box.insert("end", f"{text}\n")
        self.info_box.see("end")

    def parse_dt(self, d_str):
        try: return datetime.strptime(d_str, "%d.%m.%Y")
        except: return datetime.min

    def sync_logic(self):
        try:
            with open(self.old_json_path, 'r', encoding='utf-8') as f:
                old_data = json.load(f)
            with open(self.new_json_path, 'r', encoding='utf-8') as f:
                new_data = json.load(f)

            old_dict = {item["resident_name"].strip().lower(): item for item in old_data}
            final_db = []
            removed_count = 0

            self.log("-" * 30 + "\nАНАЛИЗ ИЗМЕНЕНИЙ:")

            for new_item in new_data:
                name_key = new_item["resident_name"].strip().lower()
                
                if name_key in old_dict:
                    old_item = old_dict[name_key]
                    
                    # 1. Слияние Check History (уникальные даты + сортировка)
                    existing_dates = {c["date"] for c in new_item.get("check_history", [])}
                    for c in old_item.get("check_history", []):
                        if c["date"] not in existing_dates:
                            new_item["check_history"].append(c)
                    new_item["check_history"].sort(key=lambda x: self.parse_dt(x["date"]))

                    # 2. Слияние Tags (метки из обоих файлов)
                    new_item["tags"] = list(set(new_item.get("tags", []) + old_item.get("tags", [])))

                    # 3. Проверка изменения адреса (сохранение в историю)
                    old_addr = old_item["location"].get("raw_address", "")
                    new_addr = new_item["location"].get("raw_address", "")
                    if old_addr != new_addr:
                        self.log(f"[!] Смена адреса у {new_item['resident_name']}")
                        history = old_item.get("location_history", [])
                        history.append({
                            "old_address": old_addr,
                            "date_archived": datetime.now().strftime("%d.%m.%Y"),
                            "reason": "sync_update"
                        })
                        new_item["location_history"] = history

                    # 4. Обработка заметок (если в новом пусто, берем из старого)
                    for note_type in ["social", "tech"]:
                        if not new_item["notes"].get(note_type) and old_item["notes"].get(note_type):
                            new_item["notes"][note_type] = old_item["notes"][note_type]

                final_db.append(new_item)

            # Вычисляем удаленных
            new_names = {i["resident_name"].strip().lower() for i in new_data}
            for old_name, obj in old_dict.items():
                if old_name not in new_names:
                    self.log(f"[-] УДАЛЕНО: {obj['resident_name']}")
                    removed_count += 1

            # Сохранение
            save_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")])
            if save_path:
                with open(save_path, 'w', encoding='utf-8') as f:
                    json.dump(final_db, f, ensure_ascii=False, indent=4)
                messagebox.showinfo("Успех", f"Синхронизация завершена.\nСохранено: {len(final_db)}\nУдалено: {removed_count}")

        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

if __name__ == "__main__":
    DeepJSONMerger().mainloop()