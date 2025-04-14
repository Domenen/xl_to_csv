import pandas as pd
from pathlib import Path
from cons import XL_FILE, CSV_FILE, DIR_FILES
import sys

def get_project_root() -> Path:
    """Возвращает абсолютный путь к корневой директории проекта"""
    return Path(__file__).parent.absolute().parent

def excel_to_csv(input_file: str, output_file: str, sheet_name: str = 0) -> None:
    """
    Конвертирует Excel-файл в CSV с автоматическим созданием директорий
    
    :param input_file: имя входного Excel-файла
    :param output_file: имя выходного CSV-файла
    :param sheet_name: имя или номер листа в Excel
    """
    try:
        root_dir = get_project_root()
        files_dir = root_dir / DIR_FILES
        
        input_path = files_dir / input_file
        output_path = files_dir / output_file
        
        if not input_path.exists():
            raise FileNotFoundError(f"Входной файл не найден: {input_path}")
            
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        df = pd.read_excel(input_path, sheet_name=sheet_name)
        df.to_csv(output_path, index=False, encoding='utf-8')
        
        print(f"Файл успешно конвертирован: {output_path}")
        
    except FileNotFoundError as e:
        print(f"Ошибка: {e}", file=sys.stderr)
    except pd.errors.EmptyDataError:
        print("Ошибка: Файл Excel пуст или поврежден", file=sys.stderr)
    except Exception as e:
        print(f"Неожиданная ошибка при конвертации: {e}", file=sys.stderr)

if __name__ == "__main__":
    excel_to_csv(XL_FILE, CSV_FILE)