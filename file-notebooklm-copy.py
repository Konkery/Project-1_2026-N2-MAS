import os
import sys
import shutil
import re
import argparse
from pathlib import Path
from collections import defaultdict

def parse_arguments():
    parser = argparse.ArgumentParser(
        description="Скрипт подготовки контекста проекта для Google NotebookLM."
    )
    parser.add_argument(
        '--delete',
        nargs='+',
        default=[],
        help="Исключить типы файлов из обработки и копирования. Допустимые значения: json, m, md (например: --delete json m)"
    )
    args = parser.parse_args()
    
    # Нормализуем переданные расширения (приводим к нижнему регистру и убираем точки)
    excluded = {val.lower().lstrip('.') for val in args.delete}
    
    valid_types = {'json', 'm', 'md'}
    invalid = excluded - valid_types
    if invalid:
        print(f"[ОШИБКА АРГУМЕНТОВ] Недопустимые типы файлов для --delete: {', '.join(invalid)}")
        print(f"Допустимы только: {', '.join(valid_types)}")
        sys.exit(1)
        
    return excluded

def main():
    # 0. Разбор аргументов командной строки
    excluded_types = parse_arguments()

    # 1. Определение путей относительно расположения скрипта (корень проекта)
    BASE_DIR = Path(__file__).resolve().parent
    TEMP_DIR = BASE_DIR / "03__Проект" / "temp"

    # 2. Инициализация регулярных выражений
    re_readme = re.compile(r'^README', re.IGNORECASE)
    re_temp_md = re.compile(r'(^|[-_]+)temp([-_]+|$)', re.IGNORECASE)
    re_prefix = re.compile(r'^\d+__')
    re_versioned = re.compile(r'^(.*?)[ _-]+rev\.(\d+)[ _-]+v(\d+)$', re.IGNORECASE)

    # Директории, которые скрипт полностью игнорирует
    IGNORED_DIRS = {'00_raw', '.git', '__pycache__', '.venv', '.idea'}

    staged_files = {}  # dict: Итоговое_Имя -> Исходный_Путь
    versioned_files = defaultdict(list)

    # 3. Функция контроля дубликатов (Fail-Fast)
    def add_to_staged(dest_filename, src_path):
        if dest_filename in staged_files:
            print("\n[ФАТАЛЬНАЯ ОШИБКА КОЛЛИЗИИ ИМЕН]")
            print(f"В проекте обнаружено два файла, которые претендуют на одно имя: '{dest_filename}'")
            print(f"Файл 1: {staged_files[dest_filename]}")
            print(f"Файл 2: {src_path}")
            print(">>> Работа скрипта остановлена. Пожалуйста, устраните дубликат или переименуйте файл.")
            sys.exit(1)
        staged_files[dest_filename] = src_path

    # Информационный вывод о параметрах запуска
    if excluded_types:
        print(f"[ПАРАМЕТРЫ] Исключены из копирования типы: {', '.join(sorted(excluded_types))}")

    # 4. Предварительная очистка целевой директории
    if TEMP_DIR.exists():
        print(f"[ОЧИСТКА] Очищаю директорию назначения: {TEMP_DIR}")
        for file_path in TEMP_DIR.glob('*'):
            if file_path.is_file() and file_path.suffix.lower() in ['.md', '.txt']:
                file_path.unlink()
    else:
        print(f"[ИНФО] Директория назначения не найдена. Создаю: {TEMP_DIR}")
        TEMP_DIR.mkdir(parents=True, exist_ok=True)

    # 5. Рекурсивный обход проекта
    print("[СКАНИРОВАНИЕ] Анализирую структуру проекта...")
    
    def scan_directory(current_dir):
        try:
            for item in current_dir.iterdir():
                # Обработка папок
                if item.is_dir():
                    if item.name in IGNORED_DIRS:
                        continue
                    if item.resolve() == TEMP_DIR.resolve():
                        continue
                    scan_directory(item)
                
                # Обработка файлов
                elif item.is_file():
                    ext = item.suffix.lower()
                    stem = item.stem

                    # --- А. Документация Markdown ---
                    if ext == '.md' and 'md' not in excluded_types:
                        if re_temp_md.search(stem):
                            continue
                        
                        if re_readme.match(stem):
                            file_size = item.stat().st_size
                            if file_size < 500:
                                print(f"[ПРОПУСК ЧЕРНОВИКА] {item.name} ({file_size} байт < 500 байт)")
                                continue
                            add_to_staged(item.name, item)

                    # --- Б. Конфигурации JSON ---
                    elif ext == '.json' and 'json' not in excluded_types:
                        clean_stem = re_prefix.sub('', stem).strip()
                        match = re_versioned.match(clean_stem)
                        dest_name = f"{clean_stem}.txt"
                        
                        if match:
                            basename = match.group(1).strip()
                            rev = int(match.group(2))
                            version = int(match.group(3))
                            versioned_files[(ext, basename)].append((rev, version, dest_name, item))
                        else:
                            add_to_staged(dest_name, item)

                    # --- В. Скрипты Power Query (.m) ---
                    elif ext == '.m' and 'm' not in excluded_types:
                        clean_stem = re_prefix.sub('', stem).strip()
                        match = re_versioned.match(clean_stem)
                        dest_name = f"{clean_stem}.txt"
                        
                        if match:
                            basename = match.group(1).strip()
                            rev = int(match.group(2))
                            version = int(match.group(3))
                            versioned_files[(ext, basename)].append((rev, version, dest_name, item))
                        else:
                            add_to_staged(dest_name, item)
                            
        except PermissionError:
            pass

    scan_directory(BASE_DIR)

    # 6. Разрешение версий (Выбор старшей ревизии)
    for (ext, basename), versions in versioned_files.items():
        # Сортировка: сначала по ревизии [0] (убывание), затем по версии [1] (убывание)
        versions.sort(key=lambda x: (x[0], x[1]), reverse=True)
        
        best_rev, best_ver, best_dest_name, best_src_path = versions[0]
        
        # Информирование о пропущенных младших версиях
        for old_rev, old_ver, old_dest, old_src in versions[1:]:
            print(f"[ПРОПУСК СТАРОЙ ВЕРСИИ] {old_src.name} (выбрана старшая: {best_src_path.name})")
            
        add_to_staged(best_dest_name, best_src_path)

    # 7. Финализация и копирование
    if not staged_files:
        print("\n[ИНФО] Подходящих файлов с учетом фильтров не найдено.")
        sys.exit(0)

    print(f"\n[КОПИРОВАНИЕ] Отобрано {len(staged_files)} актуальных файлов. Копирую в temp...")
    for dest_name, src_path in staged_files.items():
        dest_path = TEMP_DIR / dest_name
        shutil.copy2(src_path, dest_path)
        
    print(f"[УСПЕХ] Готово! Контекст сформирован в: {TEMP_DIR}")

if __name__ == "__main__":
    main()