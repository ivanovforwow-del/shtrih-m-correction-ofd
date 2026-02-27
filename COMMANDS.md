# Часто используемые команды проекта коррекции чеков ККТ

## Установка зависимостей

### Установка всех зависимостей
```bash
pip install pandas openpyxl requests pywin32
```

### Установка pywin32 (если возникли проблемы)
```bash
pip install pywin32
python Scripts/pywin32_postinstall.py -install
```

## Подготовка данных

### Подготовка данных из Excel
```bash
python prepare_data.py
```

### Парсинг данных о товарах из чеков
```bash
python parse_ofd_receipts.py
```

## Запуск коррекции чеков

### Запуск в тестовом режиме (обрабатывает только 1 чек)
```bash
python correction_final.py
```
> **Примечание:** Убедитесь, что в файле `correction_final.py` установлено `MODE = 'test'`

### Запуск в полном режиме (обрабатывает все чеки)
```bash
python correction_final.py
```
> **Примечание:** Убедитесь, что в файле `correction_final.py` установлено `MODE = 'prod'`

## Отладка и тестирование

### Запуск альтернативных версий парсера
```bash
python parse_ofd_receipts_v2.py
python parse_ofd_receipts_v3.py
python parse_ofd_receipts_v4.py
```

### Запуск отладочных скриптов
```bash
python debug_live_fetch.py
python debug_parser.py
python analyze_html.py
```

## Управление состоянием

### Проверка уже обработанных чеков
```bash
type processed.json
```
или
```bash
cat processed.json
```

### Сброс состояния (удаление списка обработанных чеков)
```bash
del processed.json
```
или
```bash
rm processed.json
```

### Просмотр лога операций
```bash
type correction_process.log
```
или
```bash
cat correction_process.log
```

## Проверка файлов данных

### Просмотр первых строк основного CSV файла
```bash
# В PowerShell
Get-Content list.csv -Head 10

# В cmd
type list.csv
```

### Просмотр первых строк CSV файла с товарами
```bash
# В PowerShell
Get-Content receipts_data.csv -Head 10

# В cmd
type receipts_data.csv
```

## Управление настройками

### Изменение режима работы (в файле correction_final.py)
- Для тестового режима: `MODE = 'test'`
- Для полного режима: `MODE = 'prod'`

### Изменение порта подключения к ККТ (в файле correction_final.py)
- `COM_PORT = 3` (или другой номер порта)

### Изменение ставки НДС (в файле correction_final.py)
- `VAT_RATE = 22` (или другое значение)

## Управление файлами

### Создание резервной копии важных файлов
```bash
copy list.csv list.csv.backup
copy receipts_data.csv receipts_data.csv.backup
copy processed.json processed.json.backup
```

### Проверка наличия файлов
```bash
dir *.csv
dir *.json
dir *.log
```

## Диагностика проблем

### Проверка установленных пакетов
```bash
pip list | findstr -i "pandas openpyxl requests pywin32"
```

### Проверка версии Python
```bash
python --version
```

### Запуск с детальным выводом
```bash
python -u correction_final.py
```

## Работа с Excel файлами

### Проверка структуры Excel файла
Скрипт `prepare_data.py` автоматически выводит доступные колонки при запуске.

### Проверка гиперссылок в Excel файле
Скрипт `prepare_data.py` включает отладочный вывод первых 3 строк при запуске.

## Проверка подключения к ККТ

### Тестирование подключения (в отдельном скрипте)
Создайте временный скрипт для проверки подключения:

```python
import win32com.client

try:
    drv = win32com.client.Dispatch('AddIn.DrvFR')
    drv.ComNumber = 3  # измените на ваш порт
    drv.BaudRate = 115200
    
    result = drv.Connect()
    
    if result == 0:
        print("Подключено к ККТ")
        print(f"ECRMode: {drv.ECRMode}")
        
        # Можно выполнить простую операцию для проверки
        drv.Disconnect()
        print("Отключено от ККТ")
    else:
        print(f"Ошибка подключения: {result}")
        
except Exception as e:
    print(f"Ошибка: {e}")
```

## Полезные команды для разработки

### Проверка синтаксиса Python файлов
```bash
python -m py_compile correction_final.py
python -m py_compile prepare_data.py
python -m py_compile parse_ofd_receipts.py
```

### Запуск с ограничением по времени (для тестирования)
```bash
timeout /t 60 /nobreak python correction_final.py
```

## Работа с документацией

### Просмотр документации
Все документы находятся в директории `docs/`:
- `docs/DETAILED_DOCS.md` - подробная документация
- `docs/API_REFERENCE.md` - справочник API
- `docs/CONFIGURATION.md` - настройка проекта
- `docs/INSTALLATION.md` - установка и запуск
- `docs/TROUBLESHOOTING.md` - решение проблем
- `docs/PROJECT_STRUCTURE.md` - структура проекта
- `docs/DEVELOPMENT.md` - руководство для разработчиков