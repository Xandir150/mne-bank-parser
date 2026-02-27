"""Internationalization support for Izvod web interface."""

TRANSLATIONS = {
    "ru": {
        # Navbar
        "app_title": "Извод",
        "dashboard": "Главная",
        "lang_label": "Язык",

        # Stats cards
        "total_statements": "Всего выписок",
        "new": "Новые",
        "reviewed": "Проверенные",
        "exported": "Экспортированные",

        # Upload
        "upload_statement": "Загрузить выписку",
        "bank": "Банк",
        "select_bank": "Выберите банк...",
        "file": "Файл",
        "drag_drop": "Перетащите файл сюда или нажмите для выбора",
        "upload": "Загрузить",
        "uploaded_ok": "Файл загружен успешно",
        "upload_failed": "Ошибка загрузки",

        # Statements table
        "statements": "Выписки",
        "auto_refresh": "Автообновление каждые 30 сек",
        "bank_name": "Банк",
        "date": "Дата",
        "account": "Счёт",
        "client": "Клиент",
        "debit": "Дебет",
        "credit": "Кредит",
        "tx_count": "Опер.",
        "status": "Статус",
        "no_statements": "Выписки не найдены",

        # Statement detail page
        "statement_num": "Выписка",
        "back_to_dashboard": "Назад",
        "error_label": "Ошибка",
        "statement_details": "Реквизиты выписки",
        "save_header": "Сохранить",
        "bank_code": "Код банка",
        "account_number": "Номер счёта",
        "iban": "IBAN",
        "statement_number": "Номер выписки",
        "statement_date": "Дата выписки",
        "period_start": "Начало периода",
        "period_end": "Конец периода",
        "opening_balance": "Начальный остаток",
        "closing_balance": "Конечный остаток",
        "total_debit": "Оборот дебет",
        "total_credit": "Оборот кредит",
        "currency": "Валюта",
        "client_name": "Наименование клиента",
        "client_pib": "ПИБ / ИНН",

        # Transactions table
        "transactions": "Операции",
        "save_changes": "Сохранить изменения",
        "row_num": "№",
        "value_date": "Дата вал.",
        "booking_date": "Дата книж.",
        "counterparty": "Контрагент",
        "counterparty_account": "Счёт контрагента",
        "payment_code": "Код",
        "purpose": "Назначение платежа",
        "fee": "Комиссия",
        "no_transactions": "Нет операций",

        # Actions
        "export_to_1c": "Экспорт в 1С",
        "download_1c": "Скачать файл 1С",
        "delete_statement": "Удалить выписку",
        "reparse": "Перепарсить",

        # Messages
        "header_saved": "Реквизиты сохранены",
        "save_failed": "Ошибка сохранения",
        "no_changes": "Нет изменений",
        "saved_n_tx": "Сохранено операций: {n}",
        "n_tx_failed": "Ошибок при сохранении: {n}",
        "export_ok": "Экспорт выполнен. Можно скачать файл.",
        "export_failed": "Ошибка экспорта",
        "delete_confirm": "Удалить эту выписку и все её операции?",
        "delete_failed": "Ошибка удаления",
        "source_file": "Исходный файл",
    },
    "sr": {
        # Navbar
        "app_title": "Извод",
        "dashboard": "Почетна",
        "lang_label": "Језик",

        # Stats cards
        "total_statements": "Укупно извода",
        "new": "Нови",
        "reviewed": "Прегледани",
        "exported": "Извезени",

        # Upload
        "upload_statement": "Учитај извод",
        "bank": "Банка",
        "select_bank": "Изаберите банку...",
        "file": "Фајл",
        "drag_drop": "Превуците фајл овде или кликните за избор",
        "upload": "Учитај",
        "uploaded_ok": "Фајл успјешно учитан",
        "upload_failed": "Грешка при учитавању",

        # Statements table
        "statements": "Изводи",
        "auto_refresh": "Ауто-освјежавање свака 30 сек",
        "bank_name": "Банка",
        "date": "Датум",
        "account": "Рачун",
        "client": "Клијент",
        "debit": "Задужење",
        "credit": "Одобрење",
        "tx_count": "Тр.",
        "status": "Статус",
        "no_statements": "Нема пронађених извода",

        # Statement detail page
        "statement_num": "Извод",
        "back_to_dashboard": "Назад",
        "error_label": "Грешка",
        "statement_details": "Подаци о изводу",
        "save_header": "Сачувај",
        "bank_code": "Шифра банке",
        "account_number": "Број рачуна",
        "iban": "IBAN",
        "statement_number": "Број извода",
        "statement_date": "Датум извода",
        "period_start": "Почетак периода",
        "period_end": "Крај периода",
        "opening_balance": "Почетно стање",
        "closing_balance": "Коначно стање",
        "total_debit": "Укупно задужење",
        "total_credit": "Укупно одобрење",
        "currency": "Валута",
        "client_name": "Назив клијента",
        "client_pib": "ПИБ",

        # Transactions table
        "transactions": "Трансакције",
        "save_changes": "Сачувај измјене",
        "row_num": "Р.бр.",
        "value_date": "Датум вал.",
        "booking_date": "Датум књиж.",
        "counterparty": "Партнер",
        "counterparty_account": "Рачун партнера",
        "payment_code": "Шифра",
        "purpose": "Сврха плаћања",
        "fee": "Накнада",
        "no_transactions": "Нема трансакција",

        # Actions
        "export_to_1c": "Извоз у 1С",
        "download_1c": "Преузми 1С фајл",
        "delete_statement": "Обриши извод",
        "reparse": "Поново парсирај",

        # Messages
        "header_saved": "Подаци сачувани",
        "save_failed": "Грешка чувања",
        "no_changes": "Нема измјена",
        "saved_n_tx": "Сачувано трансакција: {n}",
        "n_tx_failed": "Грешака при чувању: {n}",
        "export_ok": "Извоз успјешан. Можете преузети фајл.",
        "export_failed": "Грешка извоза",
        "delete_confirm": "Обрисати овај извод и све трансакције?",
        "delete_failed": "Грешка брисања",
        "source_file": "Изворни фајл",
    },
}

DEFAULT_LANG = "ru"
SUPPORTED_LANGS = list(TRANSLATIONS.keys())


def get_translations(lang: str) -> dict:
    """Return translations dict for the given language."""
    if lang not in TRANSLATIONS:
        lang = DEFAULT_LANG
    return TRANSLATIONS[lang]
