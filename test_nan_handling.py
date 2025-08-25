import pandas as pd
import numpy as np

def format_field_value(value) -> str:
    """
    Форматирует значение поля, обрабатывая NaN значения.
    
    Args:
        value: Значение поля
    
    Returns:
        Отформатированная строка или пустая строка для NaN
    """
    if pd.isna(value):
        return ""
    return str(value)

# Тестируем различные случаи
test_values = [
    "Выполнить очистку",  # Нормальное значение
    "",                   # Пустая строка
    None,                 # None
    np.nan,              # NaN
    "Проверить состояние" # Другое нормальное значение
]

print("Тестирование обработки NaN значений:")
print("=" * 50)

for i, value in enumerate(test_values):
    result = format_field_value(value)
    print(f"Test {i+1}: {repr(value)} -> '{result}'")
    
    # Проверяем, что для NaN значений возвращается пустая строка
    if pd.isna(value):
        assert result == "", f"Expected empty string for NaN value, got '{result}'"
        print("  ✅ NaN правильно обработан как пустая строка")
    else:
        # Для не-NaN значений просто проверяем, что результат - это строка
        assert isinstance(result, str), f"Expected string result, got {type(result)}"
        print("  ✅ Обычное значение правильно преобразовано в строку"))

print("\n🎉 Все тесты пройдены успешно!")
print("NaN значения в поле 'Выполнить' будут отображаться как пустые строки вместо 'nan'")