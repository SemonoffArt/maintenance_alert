#!/usr/bin/env python3
"""
Тестовый скрипт для проверки работы пересчета формул Excel.
Используется для диагностики проблем с xlwings и openpyxl.
"""

import sys
from pathlib import Path

# Добавляем путь к основному модулю
sys.path.insert(0, str(Path(__file__).parent))

from maintenance_alert import (
    recalculate_excel_formulas, 
    XLWINGS_AVAILABLE, 
    OPENPYXL_AVAILABLE, 
    EXCEL_FILE
)

def test_libraries():
    """Тестирует доступность необходимых библиотек"""
    print("=" * 60)
    print("🔍 ТЕСТИРОВАНИЕ ДОСТУПНОСТИ БИБЛИОТЕК")
    print("=" * 60)
    
    print(f"xlwings доступен: {'✅ Да' if XLWINGS_AVAILABLE else '❌ Нет'}")
    print(f"openpyxl доступен: {'✅ Да' if OPENPYXL_AVAILABLE else '❌ Нет'}")
    
    if not XLWINGS_AVAILABLE and not OPENPYXL_AVAILABLE:
        print("\n❌ Ни одна из необходимых библиотек недоступна!")
        print("💡 Установите: pip install xlwings openpyxl")
        return False
    
    return True

def test_excel_file_access():
    """Тестирует доступность Excel файла"""
    print("\n" + "=" * 60)
    print("📁 ТЕСТИРОВАНИЕ ДОСТУПА К EXCEL ФАЙЛУ")
    print("=" * 60)
    
    print(f"Путь к файлу: {EXCEL_FILE}")
    
    if EXCEL_FILE.exists():
        print("✅ Файл найден")
        print(f"Размер файла: {EXCEL_FILE.stat().st_size:,} байт")
        return True
    else:
        print("❌ Файл не найден!")
        print("💡 Убедитесь, что файл 'Обслуживание ПК и шкафов АСУТП.xlsx' находится в папке скрипта")
        return False

def test_formula_recalculation():
    """Тестирует функцию пересчета формул"""
    print("\n" + "=" * 60)
    print("🔄 ТЕСТИРОВАНИЕ ПЕРЕСЧЕТА ФОРМУЛ")
    print("=" * 60)
    
    if not EXCEL_FILE.exists():
        print("❌ Не удается провести тест - файл Excel не найден")
        return False
    
    try:
        success = recalculate_excel_formulas(EXCEL_FILE)
        if success:
            print("✅ Тест пересчета формул выполнен успешно")
            return True
        else:
            print("⚠️ Функция пересчета вернула False")
            return False
    except Exception as e:
        print(f"❌ Ошибка при тестировании пересчета: {e}")
        return False

def main():
    """Основная функция тестирования"""
    print("🧪 ТЕСТ СИСТЕМЫ ПЕРЕСЧЕТА ФОРМУЛ EXCEL")
    
    # Тест 1: Доступность библиотек
    libraries_ok = test_libraries()
    
    # Тест 2: Доступность Excel файла
    file_ok = test_excel_file_access()
    
    # Тест 3: Пересчет формул (только если файл доступен)
    recalc_ok = False
    if file_ok:
        recalc_ok = test_formula_recalculation()
    
    # Итоговый отчет
    print("\n" + "=" * 60)
    print("📊 ИТОГОВЫЙ ОТЧЕТ")
    print("=" * 60)
    print(f"Библиотеки: {'✅' if libraries_ok else '❌'}")
    print(f"Файл Excel: {'✅' if file_ok else '❌'}")
    print(f"Пересчет формул: {'✅' if recalc_ok else '❌'}")
    
    if libraries_ok and file_ok and recalc_ok:
        print("\n🎉 ВСЕ ТЕСТЫ ПРОШЛИ УСПЕШНО!")
        print("Система готова к работе с автоматическим пересчетом формул Excel.")
    else:
        print("\n⚠️ ОБНАРУЖЕНЫ ПРОБЛЕМЫ")
        if not libraries_ok:
            print("- Установите недостающие библиотеки: pip install xlwings openpyxl")
        if not file_ok:
            print("- Убедитесь, что Excel файл находится в правильном месте")
        if not recalc_ok and file_ok:
            print("- Проверьте, что файл Excel не заблокирован другой программой")

if __name__ == "__main__":
    main()