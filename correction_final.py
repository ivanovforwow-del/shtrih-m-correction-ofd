# -*- coding: utf-8 -*-
"""
Скрипт для проведения чеков коррекции на ККТ Штрих-М
Использует Драйвер ККТ ККТЛаб v.5.20

Алгоритм:
1. Читает данные из CSV (list.csv + receipts_data.csv)
2. Исключает чеки с НДС 10%
3. Для каждого чека:
   а) Чек коррекции "Возврат прихода" (отмена ошибочного чека без НДС)
   б) Чек коррекции "Приход" (с НДС 22%, сумма НЕ МЕНЯЕТСЯ)

НДС 22% действует с 01.01.2026
В чеке коррекции прихода сумма остаётся той же, но выделяется НДС 22% из суммы.
"""
import csv
import os
import sys
import time
import json
from datetime import datetime, timedelta

try:
    import win32com.client
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    print("WARNING: win32com not available")

# Настройки
CSV_FILE = 'list.csv'
ITEMS_FILE = 'receipts_data.csv'
LOG_FILE = 'correction_process.log'
PROCESSED_FILE = 'processed.json'

# Режим работы: 'test' - тестовый (1 чек), 'prod' - полный
MODE = 'test'  # Изменить на 'prod' для реальной работы

# Настройки подключения к ККТ
COM_PORT = 3
BAUD_RATE = 115200

# НДС 22% (действует с 01.01.2026)
VAT_RATE = 22

# FP чеков с НДС 10% для исключения
VAT_10_FP = [
    '607805230', '3560220730', '572166526', '760931914', '4075689541',
    '1035906310', '896116170', '1827453878', '2163166593', '1066150790',
    '2713257049', '1825692796', '2808322941'
]


def log(message):
    """Логирование в файл и консоль"""
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    log_line = f"[{timestamp}] {message}"
    print(log_line, flush=True)
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(log_line + '\n')


def load_processed():
    """Загружает список уже обработанных чеков"""
    if os.path.exists(PROCESSED_FILE):
        try:
            with open(PROCESSED_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return []
    return []


def save_processed(fp):
    """Сохраняет fp обработанного чека"""
    processed = load_processed()
    if fp not in processed:
        processed.append(fp)
        with open(PROCESSED_FILE, 'w', encoding='utf-8') as f:
            json.dump(processed, f, ensure_ascii=False, indent=2)


def load_csv_data(csv_file):
    """Загрузка данных из CSV файла"""
    data = []
    try:
        with open(csv_file, 'r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f, delimiter=';')
            for row in reader:
                fp = row['fiscal_sign']
                # Пропускаем чеки с НДС 10%
                if fp in VAT_10_FP:
                    log(f"  Пропуск чека {fp} - содержит НДС 10%")
                    continue
                data.append({
                    'summ': float(row['summ']),
                    'type': int(row['type']),
                    'fiscal_sign': fp
                })
        log(f"Загружено чеков: {len(data)}")
        return data
    except Exception as e:
        log(f"Ошибка загрузки CSV: {e}")
        return []


def load_items_data(items_file):
    """Загрузка данных о товарах из CSV файла"""
    items = {}
    dates = {}  # fp -> date string
    try:
        with open(items_file, 'r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f, delimiter=';')
            for row in reader:
                fp = row['fp']
                # Пропускаем чеки с НДС 10%
                if fp in VAT_10_FP:
                    continue
                if fp not in items:
                    items[fp] = []
                    # Сохраняем дату первого товара для этого FP
                    if 'date' in row and row['date']:
                        dates[fp] = row['date']
                
                quantity = float(row['quantity']) if row['quantity'] else 0
                price = float(row['price']) if row['price'] else 0
                summ = float(row['summ']) if row['summ'] else 0
                
                # Добавляем только товары с суммой > 0
                if summ > 0:
                    items[fp].append({
                        'name': row['name'],
                        'quantity': quantity,
                        'unit': row['unit'],
                        'price': price,
                        'summ': summ
                    })
        log(f"Загружено товаров: {sum(len(v) for v in items.values())} для {len(items)} чеков")
        log(f"Загружено дат: {len(dates)}")
        return items, dates
    except Exception as e:
        log(f"Ошибка загрузки товаров: {e}")
        return {}, {}


def connect_kkt():
    """Подключение к ККТ"""
    log("Подключение к ККТ...")
    
    if not WIN32_AVAILABLE:
        log("ERROR: win32com не доступен - работаем в тестовом режиме")
        return None
    
    try:
        drv = win32com.client.Dispatch('AddIn.DrvFR')
        log("COM объект DrvFR создан")
        
        drv.ComNumber = COM_PORT
        drv.BaudRate = BAUD_RATE
        
        result = drv.Connect()
        
        if result == 0:
            log(f"Подключено к ККТ (порт COM{COM_PORT})")
            log(f"ECRMode: {drv.ECRMode}")
            return drv
        else:
            log(f"Ошибка подключения: код {result}")
            log(f"ResultCode: {drv.ResultCode}")
            log(f"ResultCodeDescription: {drv.ResultCodeDescription}")
            return None
            
    except Exception as e:
        log(f"Исключение при подключении: {e}")
        return None


def disconnect_kkt(kkt):
    """Отключение от ККТ"""
    if kkt is None:
        return
    try:
        kkt.Disconnect()
        log("Отключено от ККТ")
    except Exception as e:
        log(f"Ошибка отключения: {e}")


def date_to_unix(date_str):
    """Конвертация строки даты в Unix timestamp
    
    ВАЖНО: Касса имеет лимит на возраст корректируемого чека (около 30 дней).
    Если оригинальный чек старше лимита, используем вчерашнюю дату.
    """
    if not date_str:
        return None
    try:
        # Формат: 2026-01-01 00:06:00
        dt = datetime.strptime(date_str, '%Y-%m-%d %H:%M:%S')
        original_timestamp = int(dt.timestamp())
        
        # Проверяем возраст чека (лимит 30 дней)
        now = datetime.now()
        days_diff = (now - dt).days
        
        if days_diff > 30:
            # Используем вчерашнюю дату
            yesterday = now - timedelta(days=1)
            yesterday = yesterday.replace(hour=12, minute=0, second=0, microsecond=0)
            log(f"   Чек старше 30 дней ({days_diff} дней), используем вчерашнюю дату: {yesterday}")
            return int(yesterday.timestamp())
        
        return original_timestamp
    except:
        return None


def date_to_driver_format(date_str):
    """Конвертация строки даты в формат драйвера ДД.ММ.ГГГГ
    
    ВАЖНО: Касса имеет лимит на возраст корректируемого чека (около 30 дней).
    Если оригинальный чек старше лимита, используем вчерашнюю дату.
    """
    if not date_str:
        return None
    try:
        # Формат: 2026-01-01 00:06:00
        dt = datetime.strptime(date_str, '%Y-%m-%d %H:%M:%S')
        
        # Проверяем возраст чека (лимит 30 дней)
        now = datetime.now()
        days_diff = (now - dt).days
        
        if days_diff > 30:
            # Используем вчерашнюю дату
            yesterday = now - timedelta(days=1)
            yesterday = yesterday.replace(hour=12, minute=0, second=0, microsecond=0)
            log(f"   Чек старше 30 дней ({days_diff} дней), используем вчерашнюю дату: {yesterday}")
            dt = yesterday
        
        # Формат драйвера: ДД.ММ.ГГГГ
        return dt.strftime('%d.%m.%Y')
    except:
        return None


def send_tlv(kkt, tag, value):
    """
    Отправка TLV-тега через FNSendTLV
    
    Формат TLV:
    - Тег: 2 байта (little-endian)
    - Длина: 2 байта (little-endian)
    - Значение: данные
    
    tag - номер тега (int)
    value - значение (bytes или str для строковых тегов)
    """
    try:
        if isinstance(value, str):
            value_bytes = value.encode('utf-8')
        else:
            value_bytes = value
        
        # Формируем TLV в hex-формате
        tag_hex = tag.to_bytes(2, 'little').hex()
        len_hex = len(value_bytes).to_bytes(2, 'little').hex()
        value_hex = value_bytes.hex()
        tlv_hex = tag_hex + len_hex + value_hex
        
        log(f"   TLV: tag={tag}, len={len(value_bytes)}, hex={tlv_hex}")
        
        # Используем TLVDataHex для передачи данных
        kkt.TLVDataHex = tlv_hex
        result = kkt.FNSendTLV()
        
        if result != 0:
            log(f"   FNSendTLV error: {result}, ResultCode: {kkt.ResultCode}")
            return False
        
        log(f"   FNSendTLV OK")
        return True
    except Exception as e:
        log(f"   FNSendTLV exception: {e}")
        return False


def send_stlv_tag(kkt, stlv_tag, tags_dict):
    """
    Отправка STLV-тега (структурированного) через FNBeginSTLVTag + FNAddTag + FNSendSTLVTag
    
    stlv_tag - номер родительского STLV-тега (например, 1174 - Основание для коррекции)
    tags_dict - словарь вложенных тегов {tag_number: (tag_type, value)}
                tag_type: 0 - int, 1 - string, 2 - datetime, 3 - binary
    """
    try:
        log(f"   STLV: tag={stlv_tag}")
        
        # Начинаем формирование STLV-тега
        kkt.TagNumber = stlv_tag
        result = kkt.FNBeginSTLVTag()
        
        if result != 0:
            log(f"   FNBeginSTLVTag error: {result}, ResultCode: {kkt.ResultCode}")
            return False
        
        parent_id = kkt.TagID
        log(f"   FNBeginSTLVTag OK, TagID={parent_id}")
        
        # Добавляем вложенные теги
        for tag_num, (tag_type, value) in tags_dict.items():
            kkt.TagNumber = tag_num
            kkt.TagType = tag_type
            
            if tag_type == 0:  # int
                kkt.TagValueInt = value
            elif tag_type == 1:  # string
                kkt.TagValueStr = value
            elif tag_type == 2:  # datetime
                kkt.TagValueDateTime = value
            elif tag_type == 3:  # binary
                kkt.TagValueBin = value
            
            result = kkt.FNAddTag()
            log(f"   FNAddTag: tag={tag_num}, type={tag_type}, value={value}, result={result}")
            
            if result != 0:
                log(f"   FNAddTag error: {result}, ResultCode: {kkt.ResultCode}")
                return False
        
        # Отправляем сформированный STLV-тег
        result = kkt.FNSendSTLVTag()
        log(f"   FNSendSTLVTag: result={result}")
        
        if result != 0:
            log(f"   FNSendSTLVTag error: {result}, ResultCode: {kkt.ResultCode}")
            return False
        
        log(f"   STLV OK")
        return True
    except Exception as e:
        log(f"   STLV exception: {e}")
        return False


def send_tlv_date(kkt, tag, unix_timestamp):
    """Отправка TLV-тега с датой (Unix timestamp, 4 байта little-endian)"""
    value = unix_timestamp.to_bytes(4, 'little')
    return send_tlv(kkt, tag, value)


def send_tlv_string(kkt, tag, string_value):
    """Отправка TLV-тега со строкой"""
    return send_tlv(kkt, tag, string_value)


def correction_refund(kkt, summ, payment_type, fiscal_sign, items, receipt_date=None):
    """
    Чек коррекции "Возврат прихода" (отмена ошибочного чека)
    
    Использует FNOpenCheckCorrection + FNSendTag + FNOperation + FNCloseCheckEx
    
    Товары добавляются БЕЗ НДС (как в оригинальном чеке)
    
    receipt_date - дата оригинального чека (строка в формате YYYY-MM-DD HH:MM:SS)
    """
    log("=" * 50)
    log("ЧЕК КОРРЕКЦИИ: ОТМЕНА ПРИХОДА (возврат прихода)")
    log("=" * 50)
    log(f"Сумма: {summ} руб.")
    log(f"Тип оплаты: {'безнал' if payment_type == 0 else 'наличные'}")
    log(f"Фискальный признак: {fiscal_sign}")
    log(f"Дата оригинального чека: {receipt_date}")
    
    if kkt is None:
        log(f"[TEST MODE] Возврат прихода: {summ} руб.")
        return True
    
    try:
        # 1. Установка CheckType = 2 (Возврат прихода)
        log("\n1. Установка CheckType = 2 (Возврат прихода)...")
        kkt.CheckType = 2
        log(f"   CheckType = {kkt.CheckType}")
        
        # 2. Установка CorrectionType = 0 (самостоятельная)
        log("\n2. Установка CorrectionType = 0...")
        kkt.CorrectionType = 0
        log(f"   CorrectionType = {kkt.CorrectionType}")
        
        # 3. FNOpenCheckCorrection - открыть чек коррекции (ФФД 1.2)
        log("\n3. FNOpenCheckCorrection...")
        result = kkt.FNOpenCheckCorrection()
        log(f"   Result: {result}, ResultCode: {kkt.ResultCode}")
        
        if result != 0:
            log(f"   Error: {kkt.ResultCodeDescription}")
            return False
        
        log("   Чек открыт!")
        
        # 4. Отправка тега 1178 "Дата документа-основания" напрямую через FNSendTag
        if receipt_date:
            date_str = date_to_driver_format(receipt_date)
            if date_str:
                log(f"\n4. Отправка тега 1178 (Дата документа-основания)...")
                log(f"   Дата: {date_str}")
                
                kkt.TagNumber = 1178
                kkt.TagType = 1  # string
                kkt.TagValueStr = date_str
                result = kkt.FNSendTag()
                log(f"   FNSendTag(1178): result={result}")
                
                if result != 0:
                    log(f"   FNSendTag(1178) error: {result}, ResultCode: {kkt.ResultCode}")
        
        # 5. Отправка тега 1192 "ФП корректируемого чека" через FNSendTag
        log(f"\n5. Отправка тега 1192 (ФП корректируемого чека)...")
        log(f"   ФП: {fiscal_sign}")
        kkt.TagNumber = 1192
        kkt.TagType = 1  # string
        kkt.TagValueStr = str(fiscal_sign)
        result = kkt.FNSendTag()
        log(f"   FNSendTag(1192): result={result}")
        
        if result != 0:
            log(f"   FNSendTag error: {result}, ResultCode: {kkt.ResultCode}")
        
        # 5. Добавляем товары
        if items:
            for item in items:
                log(f"\n5. Добавление товара: {item['name']}")
                log(f"   Количество: {item['quantity']} {item['unit']}")
                log(f"   Цена: {item['price']} руб.")
                log(f"   Сумма: {item['summ']} руб. (БЕЗ НДС)")
                
                kkt.CheckType = 2  # Возврат прихода для FNOperation
                kkt.Price = item['price']
                kkt.Quantity = item['quantity']
                kkt.Summ1 = item['summ']
                kkt.Summ1Enabled = True
                kkt.Tax1 = 0  # Без НДС (как в оригинале)
                kkt.StringForPrinting = item['name'][:128]
                kkt.PaymentTypeSign = 4  # Полный расчёт
                kkt.PaymentItemSign = 1  # Товар
                
                result = kkt.FNOperation()
                log(f"   FNOperation: {result}, ResultCode: {kkt.ResultCode}")
                
                if result != 0:
                    log(f"   Error: {kkt.ResultCodeDescription}")
                    kkt.CancelCheck()
                    return False
        else:
            # Если товаров нет, добавляем одну позицию
            log(f"\n5. Добавление позиции (без товара)...")
            kkt.CheckType = 2
            kkt.Price = summ
            kkt.Quantity = 1
            kkt.Summ1 = summ
            kkt.Summ1Enabled = True
            kkt.Tax1 = 0  # Без НДС
            kkt.StringForPrinting = "Коррекция (отмена прихода)"
            kkt.PaymentTypeSign = 4
            kkt.PaymentItemSign = 1
            
            result = kkt.FNOperation()
            log(f"   FNOperation: {result}, ResultCode: {kkt.ResultCode}")
            
            if result != 0:
                log(f"   Error: {kkt.ResultCodeDescription}")
                kkt.CancelCheck()
                return False
        
        # 6. FNCloseCheckEx
        log("\n6. FNCloseCheckEx...")
        kkt.Summ1 = summ
        log(f"   Summ1 (итог) = {summ}")
        
        if payment_type == 1:
            kkt.Summ2 = summ
            log(f"   Summ2 (наличные) = {summ}")
        else:
            kkt.Summ3 = summ
            log(f"   Summ3 (электронные) = {summ}")
        
        result = kkt.FNCloseCheckEx()
        log(f"   FNCloseCheckEx: {result}, ResultCode: {kkt.ResultCode}")
        
        if result != 0:
            log(f"   Error: {kkt.ResultCodeDescription}")
            return False
        
        log("\n>>> Чек коррекции 'Отмена прихода' успешно пробит! <<<")
        return True
        
    except Exception as e:
        log(f"Ошибка: {e}")
        import traceback
        traceback.print_exc()
        try:
            kkt.CancelCheck()
        except:
            pass
        return False


def correction_sale(kkt, summ, payment_type, items, vat_rate, fiscal_sign=None, receipt_date=None):
    """
    Чек коррекции "Приход" с НДС 22%
    
    Использует FNOpenCheckCorrection + FNSendTag + FNOperation + FNCloseCheckEx
    
    ВАЖНО: Сумма НЕ МЕНЯЕТСЯ!
    НДС 22% выделяется из суммы (сумма включает НДС).
    
    Формула выделения НДС из суммы:
    НДС = Сумма * 22 / 122
    Сумма без НДС = Сумма - НДС
    
    fiscal_sign - ФП оригинального чека (для тега 1192)
    receipt_date - дата оригинального чека (для тега 1178)
    """
    log("=" * 50)
    log("ЧЕК КОРРЕКЦИИ: ПРИХОД (с НДС 22%)")
    log("=" * 50)
    
    # Расчет НДС 22% (выделение из суммы)
    # НДС = сумма * 22 / 122 (выделение НДС из суммы с НДС)
    vat = summ * vat_rate / (100 + vat_rate)
    vat = round(vat, 2)
    
    log(f"Сумма чека: {summ} руб. (НЕ ИЗМЕНЯЕТСЯ)")
    log(f"НДС {vat_rate}% (выделен из суммы): {vat} руб.")
    log(f"Тип оплаты: {'безнал' if payment_type == 0 else 'наличные'}")
    log(f"ФП оригинального чека: {fiscal_sign}")
    log(f"Дата оригинального чека: {receipt_date}")
    
    if kkt is None:
        log(f"[TEST MODE] Приход: {summ} руб., НДС {vat_rate}% = {vat} руб.")
        return True
    
    try:
        # 1. CheckType = 0 (Приход)
        log("\n1. Установка CheckType = 0 (Приход)...")
        kkt.CheckType = 0
        log(f"   CheckType = {kkt.CheckType}")
        
        # 2. CorrectionType = 0 (самостоятельная)
        log("\n2. Установка CorrectionType = 0...")
        kkt.CorrectionType = 0
        log(f"   CorrectionType = {kkt.CorrectionType}")
        
        # 3. FNOpenCheckCorrection - открыть чек коррекции (ФФД 1.2)
        log("\n3. FNOpenCheckCorrection...")
        result = kkt.FNOpenCheckCorrection()
        log(f"   Result: {result}, ResultCode: {kkt.ResultCode}")
        
        if result != 0:
            log(f"   Error: {kkt.ResultCodeDescription}")
            return False
        
        log("   Чек открыт!")
        
        # 4. Отправка тега 1178 "Дата документа-основания" напрямую через FNSendTag
        if receipt_date:
            date_str = date_to_driver_format(receipt_date)
            if date_str:
                log(f"\n4. Отправка тега 1178 (Дата документа-основания)...")
                log(f"   Дата: {date_str}")
                
                kkt.TagNumber = 1178
                kkt.TagType = 1  # string
                kkt.TagValueStr = date_str
                result = kkt.FNSendTag()
                log(f"   FNSendTag(1178): result={result}")
                
                if result != 0:
                    log(f"   FNSendTag(1178) error: {result}, ResultCode: {kkt.ResultCode}")
        
        # 5. Отправка тега 1192 "ФП корректируемого чека" через FNSendTag
        if fiscal_sign:
            log(f"\n   Тег 1192 (ФП корректируемого чека): {fiscal_sign}")
            kkt.TagNumber = 1192
            kkt.TagType = 1  # string
            kkt.TagValueStr = str(fiscal_sign)
            result = kkt.FNSendTag()
            log(f"   FNSendTag(1192): result={result}")
            
            if result != 0:
                log(f"   FNSendTag error: {result}, ResultCode: {kkt.ResultCode}")
        
        # 5. Добавляем товары
        if items:
            for item in items:
                # Расчет НДС для товара (выделение из суммы)
                item_vat_value = item['summ'] * vat_rate / (100 + vat_rate)
                item_vat_value = round(item_vat_value, 2)
                
                log(f"\n5. Добавление товара: {item['name']}")
                log(f"   Количество: {item['quantity']} {item['unit']}")
                log(f"   Сумма: {item['summ']} руб. (включает НДС)")
                log(f"   НДС {vat_rate}%: {item_vat_value} руб.")
                
                kkt.CheckType = 1  # Приход для FNOperation
                kkt.Price = item['summ']  # Цена = полная сумма с НДС
                kkt.Quantity = 1  # Количество = 1
                kkt.Summ1 = item['summ']  # Сумма с НДС (не меняется)
                kkt.Summ1Enabled = True
                kkt.Tax1 = 12  # НДС 22%
                kkt.TaxValue = item_vat_value
                kkt.TaxValueEnabled = True
                kkt.StringForPrinting = item['name'][:128]
                kkt.PaymentTypeSign = 4  # Полный расчёт
                kkt.PaymentItemSign = 1  # Товар
                
                result = kkt.FNOperation()
                log(f"   FNOperation: {result}, ResultCode: {kkt.ResultCode}")
                
                if result != 0:
                    log(f"   Error: {kkt.ResultCodeDescription}")
                    kkt.CancelCheck()
                    return False
        else:
            # Если товаров нет, добавляем одну позицию
            log(f"\n5. Добавление позиции (без товара)...")
            kkt.CheckType = 1
            kkt.Price = summ
            kkt.Quantity = 1
            kkt.Summ1 = summ
            kkt.Summ1Enabled = True
            kkt.Tax1 = 12  # НДС 22%
            kkt.TaxValue = vat
            kkt.TaxValueEnabled = True
            kkt.StringForPrinting = "Коррекция (приход с НДС)"
            kkt.PaymentTypeSign = 4
            kkt.PaymentItemSign = 1
            
            result = kkt.FNOperation()
            log(f"   FNOperation: {result}, ResultCode: {kkt.ResultCode}")
            
            if result != 0:
                log(f"   Error: {kkt.ResultCodeDescription}")
                kkt.CancelCheck()
                return False
        
        # 6. FNCloseCheckEx
        log("\n6. FNCloseCheckEx...")
        kkt.Summ1 = summ
        log(f"   Summ1 (итог) = {summ}")
        
        if payment_type == 1:
            kkt.Summ2 = summ
            log(f"   Summ2 (наличные) = {summ}")
        else:
            kkt.Summ3 = summ
            log(f"   Summ3 (электронные) = {summ}")
        
        result = kkt.FNCloseCheckEx()
        log(f"   FNCloseCheckEx: {result}, ResultCode: {kkt.ResultCode}")
        
        if result != 0:
            log(f"   Error: {kkt.ResultCodeDescription}")
            return False
        
        log("\n>>> Чек коррекции 'Приход' успешно пробит! <<<")
        return True
        
    except Exception as e:
        log(f"Ошибка: {e}")
        import traceback
        traceback.print_exc()
        try:
            kkt.CancelCheck()
        except:
            pass
        return False


def process_corrections(data, items_data, dates_data, mode='test'):
    """Обработка чеков коррекции"""
    
    # Загружаем список уже обработанных
    processed = load_processed()
    log(f"Уже обработано чеков: {len(processed)}")
    
    # Фильтруем уже обработанные
    data = [d for d in data if d['fiscal_sign'] not in processed]
    log(f"Осталось обработать: {len(data)}")
    
    if mode == 'test':
        data = data[:1]
        log(f"ТЕСТОВЫЙ РЕЖИМ: обрабатываем только 1 чек")
    
    if not data:
        log("Нет чеков для обработки!")
        return 0, 0
    
    # Подключаемся к ККТ
    kkt = connect_kkt()
    
    success_count = 0
    error_count = 0
    total = len(data)
    
    try:
        for i, check in enumerate(data, 1):
            log(f"\n{'='*60}")
            log(f"Обработка чека {i}/{total}:")
            log(f"  Сумма: {check['summ']} руб.")
            log(f"  Тип оплаты: {'наличные' if check['type'] == 1 else 'безналичные'}")
            log(f"  ФП: {check['fiscal_sign']}")
            
            # Получаем товары для этого чека
            fp = check['fiscal_sign']
            items = items_data.get(fp, [])
            receipt_date = dates_data.get(fp)
            
            if receipt_date:
                log(f"  Дата оригинального чека: {receipt_date}")
            
            if items:
                log(f"  Товаров в чеке: {len(items)}")
                for item in items:
                    log(f"    - {item['name']}: {item['quantity']} {item['unit']} x {item['price']} = {item['summ']}")
            else:
                log(f"  ВНИМАНИЕ: Товары не найдены для чека {check['fiscal_sign']}")
            
            # Шаг 1: Возврат прихода (отмена ошибочного чека)
            result1 = correction_refund(kkt, check['summ'], check['type'], check['fiscal_sign'], items, receipt_date)
            
            if result1:
                log("\n" + "-" * 60)
                log("Пауза 2 секунды между чеками...")
                time.sleep(2)
                
                # Шаг 2: Приход (с правильным НДС)
                result2 = correction_sale(kkt, check['summ'], check['type'], items, VAT_RATE, check['fiscal_sign'], receipt_date)
                
                if result2:
                    success_count += 1
                    save_processed(check['fiscal_sign'])
                    log(f"\n>>> Чек {i} успешно обработан! <<<")
                else:
                    error_count += 1
                    log(f"\n!!! ОШИБКА: Чек коррекции 'Приход' не пробит! !!!")
            else:
                error_count += 1
                log(f"\n!!! ОШИБКА: Чек коррекции 'Отмена прихода' не пробит! !!!")
            
            # Пауза между чеками
            if i < total:
                log("\nПауза 3 секунды...")
                time.sleep(3)
            
            # Прогресс
            progress = (i / total) * 100
            log(f"\nПРОГРЕСС: {i}/{total} ({progress:.1f}%) - Успешно: {success_count}, Ошибок: {error_count}")
            
    finally:
        disconnect_kkt(kkt)
    
    log(f"\n{'='*60}")
    log(f"ИТОГИ:")
    log(f"  Успешно: {success_count}")
    log(f"  Ошибки: {error_count}")
    log(f"  Всего: {total}")
    log(f"{'='*60}")
    
    return success_count, error_count


def main():
    """Главная функция"""
    log("=" * 60)
    log("НАЧАЛО ОБРАБОТКИ ЧЕКОВ КОРРЕКЦИИ")
    log(f"Дата: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log(f"CSV файл: {CSV_FILE}")
    log(f"Файл товаров: {ITEMS_FILE}")
    log(f"НДС: {VAT_RATE}%")
    log(f"Режим: {MODE}")
    log(f"Исключено чеков с НДС 10%: {len(VAT_10_FP)}")
    log("=" * 60)
    
    # Проверяем наличие файлов
    if not os.path.exists(CSV_FILE):
        log(f"ОШИБКА: Файл {CSV_FILE} не найден!")
        return
    
    if not os.path.exists(ITEMS_FILE):
        log(f"ВНИМАНИЕ: Файл {ITEMS_FILE} не найден! Товары не будут использованы.")
        items_data = {}
        dates_data = {}
    else:
        items_data, dates_data = load_items_data(ITEMS_FILE)
    
    # Загружаем данные
    data = load_csv_data(CSV_FILE)
    if not data:
        log("ОШИБКА: Нет данных для обработки")
        return
    
    # Запускаем обработку
    process_corrections(data, items_data, dates_data, mode=MODE)
    
    log("\nОбработка завершена.")


if __name__ == "__main__":
    main()
