# Архитектура

## 1. Форма приложения

Локальный веб-интерфейс, запускаемый как Windows-приложение:

- сервер слушает только `127.0.0.1`;
- внешние CDN не используются;
- исходники и база хранятся локально;
- позже сервер и WebView упаковываются в один EXE.

## 2. Слои

### Интерфейс

- входящие документы;
- карточка пакетного PDF;
- список страниц;
- очередь проблем;
- ручной редактор страницы;
- окно АБС;
- проекты ответов;
- журнал и настройки.

### Прикладной конвейер

1. Регистрация файла и SHA-256.
2. Подсчёт страниц.
3. Создание записи каждой страницы.
4. Создание неизменяемого превью.
5. Оценка качества изображения.
6. Многопроходный QR.
7. Безопасное получение официального файла.
8. Классификация страницы.
9. Извлечение полей.
10. Группировка страниц в обращения.
11. Ручное подтверждение неопределённостей.
12. Проверка АБС.
13. Правило периода.
14. Создание проекта Word.
15. Журналирование.

### Адаптеры

- `QrDecoder`;
- `OfficialDocumentClient`;
- `OcrEngine`;
- `PageClassifier`;
- `AbsGateway`;
- `WordTemplateRenderer`.

Каждый внешний механизм сменный. Реальная АБС заменяет фейковую без изменения
правил обработки.

## 3. Данные

### uploads

- id;
- original_filename;
- stored_path;
- sha256;
- page_count;
- status;
- created_at;
- completed_at.

### pages

- id;
- upload_id;
- page_number;
- preview_path;
- enhanced_preview_path;
- page_type;
- type_confidence;
- quality_score;
- qr_status;
- qr_payload_hash;
- qr_safe_url;
- ocr_status;
- ocr_confidence;
- extracted_text;
- status;
- issue_code;
- issue_message;
- case_id;
- manual_confirmed;
- created_at;
- updated_at.

### cases

- id;
- upload_id;
- status;
- source_kind;
- official_document_path;
- district_place;
- recipient_position;
- recipient_full_name;
- recipient_display_name;
- period_start;
- period_end;
- employee_name;
- fields_confirmed;
- abs_status;
- response_status;
- created_at;
- updated_at.

### taxpayers

- id;
- case_id;
- display_order;
- name;
- inn;
- name_source;
- inn_source;
- manually_confirmed;
- abs_result.

### audit_events

- id;
- entity_type;
- entity_id;
- event_type;
- actor;
- payload_json;
- created_at.

Пароли и логины АБС отсутствуют в схеме.

## 4. Статусы страницы

- `registered`;
- `preview_ready`;
- `processing`;
- `qr_resolved`;
- `scan_fallback`;
- `needs_review`;
- `manually_confirmed`;
- `completed`;
- `technical_error`.

Терминальный статус обязателен для закрытия загрузки.

## 5. Статусы обращения

- `collecting`;
- `needs_review`;
- `ready_for_abs`;
- `abs_checking`;
- `manual_period_rule`;
- `ready_for_response`;
- `response_created`;
- `completed`;
- `technical_error`.

## 6. QR

Декодирование выполняется на:

- исходном изображении;
- полном изображении с автоконтрастом;
- увеличенной версии;
- нижней правой области;
- вариантах с фиксированным и локальным порогом;
- небольших поворотах.

Хранится SHA-256 полного payload. В интерфейсе отображается только безопасная
часть URL без значения `encodedText`.

## 7. OCR

OCR обязан возвращать:

- текст;
- язык;
- уверенность по словам и символам;
- координаты;
- альтернативы без автоматического выбора;
- сведения о применённом варианте изображения.

До подключения проверенной модели `rus+kir` fallback не имеет права
автоматически подтверждать поля плохого скана.

## 8. Безопасность

- привязка к `127.0.0.1`;
- ограничение размера файлов;
- проверка сигнатуры PDF;
- безопасные имена хранения;
- разрешённые домены QR;
- таймаут и ограничение размера официального файла;
- отсутствие паролей в логах;
- оригиналы только для чтения;
- атомарные записи;
- локальная база SQLite;
- экспорт и резервное копирование как отдельная операция.

## 9. Упаковка

После стабилизации:

- PyInstaller для backend;
- локальный WebView2 или системный браузер;
- установщик Windows;
- папка данных вне каталога программы;
- миграции SQLite;
- журнал версии приложения и схемы.

