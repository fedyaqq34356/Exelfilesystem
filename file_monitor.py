# file_monitor.py
import time
from pathlib import Path
from datetime import datetime
import config
from database import Database
from excel_handler import ExcelHandler

class FileMonitor:

    
    def __init__(self):
        self.db = Database()
        self.excel = ExcelHandler()
        self.last_check = datetime.now()
        

        self.pending_files = {}

    def check_folders(self):

        new_apps = []
        

        folders = {
            "ФІНДИРЕКТОР": config.get_path("findirector_folder"),
            "ДИРЕКТОР": config.get_path("director_folder")
        }

        current_time = time.time()

        for approver, folder in folders.items():

            if not folder:
                continue
            
            folder_path = Path(folder)
            if not folder_path.exists():
                print(f"⚠️ Папка не існує: {folder}")
                continue


            for pattern in ["*.xlsm", "*.xlsx"]:
                for file in folder_path.glob(pattern):
                    fp = str(file.resolve())


                    if self.excel.is_file_locked(fp):

                        continue


                    if self.db.is_file_processed(fp):

                        self.pending_files.pop(fp, None)
                        continue


                    if fp not in self.pending_files:
                        print(f"🔍 Виявлено новий файл: {file.name} [{approver}]")
                        self.pending_files[fp] = current_time
                        continue


                    time_since_first_seen = current_time - self.pending_files[fp]
                    
                    if time_since_first_seen < config.FILE_SETTLE_TIME:

                        continue


                    print(f"📄 Обробка файлу: {file.name} [{approver}]")
                    

                    valid, error = self.excel.validate_file(fp)
                    if not valid:
                        print(f"❌ Файл не валідний: {error}")

                        self.pending_files.pop(fp, None)
                        continue


                    data = self.excel.read_application(fp)
                    
                    if data:

                        data["intended_approver"] = approver
                        

                        new_apps.append(data)
                        

                        self.db.add_processed_file(fp, approver)
                        

                        self.db.log_action(
                            data["file_name"], 
                            "DETECTED", 
                            "monitor", 
                            f"Сума: {data['сума']}, Погоджує: {approver}"
                        )
                        
                        print(f"✅ Заявка додана: {data['file_name']} ({approver})")
                    else:
                        print(f"⚠️ Не вдалося прочитати файл: {file.name}")
                    

                    self.pending_files.pop(fp, None)


        self._cleanup_pending_files()


        self.last_check = datetime.now()

        return new_apps

    def _cleanup_pending_files(self):

        files_to_remove = []
        
        for fp in self.pending_files.keys():
            if not Path(fp).exists():
                files_to_remove.append(fp)
        
        for fp in files_to_remove:
            print(f"🗑️ Видалено з pending (файл зник): {Path(fp).name}")
            self.pending_files.pop(fp, None)

    def get_monitoring_stats(self):

        stats = {
            "last_check": self.last_check.strftime("%H:%M:%S"),
            "pending_files": len(self.pending_files),
            "folders": {}
        }
        

        for key, path in config.PATHS.items():
            if key.endswith("_folder") and path:
                folder_path = Path(path)
                exists = folder_path.exists()
                
                if exists:

                    xlsm_count = len(list(folder_path.glob("*.xlsm")))
                    xlsx_count = len(list(folder_path.glob("*.xlsx")))
                    total_count = xlsm_count + xlsx_count
                else:
                    total_count = 0
                

                folder_name = key.replace("_folder", "").replace("_", " ").title()
                
                stats["folders"][folder_name] = {
                    "exists": exists,
                    "files": total_count,
                    "path": path
                }
        
        return stats

    def force_check_file(self, file_path):

        fp = str(Path(file_path).resolve())
        

        self.pending_files.pop(fp, None)
        

        
        print(f"🔄 Примусова перевірка: {Path(fp).name}")
        

        data = self.excel.read_application(fp)
        return data

    def get_pending_files_info(self):

        info = []
        current_time = time.time()
        
        for fp, first_seen in self.pending_files.items():
            time_waiting = current_time - first_seen
            info.append({
                "file": Path(fp).name,
                "waiting": f"{time_waiting:.1f}s",
                "ready_in": f"{max(0, config.FILE_SETTLE_TIME - time_waiting):.1f}s"
            })
        
        return info