# ТЗ: `POST /print-check` — JSON чека и доработка Delphi

**Статус:** реализовано (АТОЛ Fptr10, один `processJson` на чек).  
**Дата:** 2026-08-03 (миграция с Штрих)

---

## 1. Идея

JSON делится на две части:

1. **Фискальные данные** — сервер транслирует в JSON-задание АТОЛ (`sell` / `sellReturn`) и вызывает `processJson`.
2. **Нефискальные данные** — клиент передаёт элементами `"type": "text"` в `items[]`; в JSON АТОЛ уходят как текстовые элементы чека.

Сервер **не** генерирует скидочные строки, код продавца и прочее — это ответственность клиента (1С).

Внешняя схема JSON **для клиента не менялась** при миграции на АТОЛ; изменилась только внутренняя реализация.

---

## 2. Структура JSON

### 2.1. Шапка чека (фискальное)

| Поле | Тип | Обяз. | Описание | АТОЛ (JSON-задание) |
|------|-----|-------|----------|---------------------|
| `type` | string | да | `sell` — приход, `sellReturn` — возврат | `type`: `sell` / `sellReturn` |
| `taxationType` | string | да | СНО: `osn`, `usnIncome`, `usnIncomeOutcome`, `esn`, `patent` | `taxationType` (коды АТОЛ) |
| `operator.name` | string | да | ФИО кассира | `operator.name` (+ открытие смены) |
| `operator.vatin` / `cashierInn` | string | нет | ИНН кассира | `operator.vatin` |
| `payments` | array | да | Оплаты (см. ниже) | `payments[]` |
| `total` | number | да | Итог чека | `total` |
| `items` | array | да | Позиции и текст (см. ниже) | `items[]` |

Оплаты:

```json
"payments": [
  { "type": "cash", "sum": 100.00 },
  { "type": "electronically", "sum": 50.00 }
]
```

| `payments[].type` | АТОЛ |
|-------------------|------|
| `cash` | `type: "cash"` |
| `electronically` | `type: "electronically"` |
| `credit` | `type: "credit"` |

Смешанная оплата: несколько элементов в `payments[]`, суммы складываются по типам.

---

### 2.2. Позиция (`type: "position"`)

| Поле | Тип | Обяз. | Описание | АТОЛ |
|------|-----|-------|----------|------|
| `type` | `"position"` | да | — | элемент `items[]` позиции |
| `name` | string | да | Наименование | `name` |
| `price` | number | да | Цена за ед. с учётом скидок | `price` |
| `quantity` | number | да | Количество | `quantity` |
| `tax.type` | string | да | Ставка НДС (см. таблицу) | `tax.type` |
| `paymentObject` | string | да | Предмет расчёта | `paymentObject` |
| `paymentMethod` | string | да | Способ расчёта | `paymentMethod` |
| `department` | number | нет | Секция (default из настроек) | при поддержке драйвером |
| `measurementUnit` | string | нет | `piece` / `шт` → шт., иначе ед. | `measurementUnit` |
| `mark` | string | нет | КМ для ККТ (base64) | `imcParams.imc` (+ результат проверки) |
| `markAccepted` | bool | нет* | Марка уже принята ККТ | см. §3 |
| `permission` | object | нет** | Данные разрешительного режима | `industryInfo` (теги 1262–1265) |

\* Обязателен, если передан `mark`.  
\*\* Обязателен для маркированной позиции прихода — 1С всегда присылает готовый объект.

**`tax.type` → АТОЛ `tax.type`:**

| tax.type | АТОЛ |
|----------|------|
| `none` | `none` |
| `vat20` | `vat20` |
| `vat10` | `vat10` |
| `vat0` | `vat0` |
| `vat120` | `vat120` |
| `vat110` | `vat110` |
| `vat5` | `vat5` |
| `vat7` | `vat7` |
| `vat105` | `vat105` |
| `vat107` | `vat107` |

**`paymentObject`:** строки АТОЛ (`commodity`, `excise`, …); при наличии `mark` — маркированный товар.

**`paymentMethod`:** строки АТОЛ (`fullPayment`, `advance`, …).

**Разрешительный режим (`permission`):**

```json
"permission": {
  "uuid": "00000000-0000-0000-0000-000000000001",
  "time": "2026-07-05T10:00:00",
  "tag1262": "030",
  "tag1263": "21.11.2023",
  "tag1264": "1944"
}
```

| Поле | Тег ФФД | Обяз. | По умолчанию |
|------|---------|-------|--------------|
| `uuid` | 1265 (часть значения) | да* | — |
| `time` | 1265 (часть значения) | да* | — |
| `tag1262` | 1262 (идентификатор ФОИВ) | нет | `030` |
| `tag1263` | 1263 (дата основания) | нет | `21.11.2023` |
| `tag1264` | 1264 (номер основания) | нет | `1944` |

\* Обязательны для отправки отраслевого реквизита. Если `uuid`+`time` переданы, а `tag1262`–`tag1264` нет — сервер подставляет значения по умолчанию.

1С **всегда** присылает `permission` для маркированных позиций прихода (минимум `uuid`+`time`).  
Сервер кладёт данные в `industryInfo` JSON АТОЛ (не через низкоуровневые теги Штрих).

Если `permission` не передан (нештатно) — **заглушка:** `industryInfo` не формируется, в лог пишется предупреждение. В будущем можно добавить самостоятельное получение uuid/time.

---

### 2.3. Нефискальный текст (`type: "text"`)

Всё нефискальное: скидки, код продавца, «Оплата по терминалу», накопление, акции, разделители.

```json
{
  "type": "text",
  "text": "СКИДКА СУММА 20.00"
}
```

**АТОЛ:** элемент `items[]` с типом текста (нефискальная строка в том же `processJson`).

---

## 3. Логика маркировки в чеке

Для позиции с `mark`:

```
Сборка JSON АТОЛ sell/sellReturn
  ↓
позиция с imcParams (марка base64) + itemInfoCheckResult из кэша валидации
  ↓
при необходимости industryInfo из permission / кэша РР
  ↓
один processJson на весь чек
```

Перед сборкой чека сервер при необходимости допроверяет марки (`EnsureMarksReadyForPrint`):

- `markAccepted: true` или марка уже в кэше после `/check-marks` — повторный полный цикл не нужен;
- иначе — полный цикл АТОЛ (`beginMarkingCodeValidation` → poll → `acceptMarkingCode`) до печати.

Endpoint `/check-marks` — для предварительной проверки; при печати чека клиент ставит `markAccepted: true`, если марка уже принята.

---

## 4. Порядок обработки `items[]`

Элементы обрабатываются **по порядку** при сборке JSON АТОЛ:

| type | Действие |
|------|----------|
| `position` | позиция в `items[]` (+ маркировка / РР) |
| `text` | текстовый элемент в `items[]` |

После сборки — **один** `processJson` на весь чек; номер чека / серийный номер — из ответа АТОЛ (`fiscalParams` и др.).

---

## 5. Пример полного чека

```json
{
  "type": "sell",
  "taxationType": "usnIncome",
  "operator": {
    "name": "Емельянов Алексей Андреевич"
  },
  "items": [
    {
      "type": "position",
      "name": "Шлем неопреновый",
      "price": 200.00,
      "quantity": 1,
      "department": 1,
      "tax": { "type": "none" },
      "paymentObject": "commodity",
      "paymentMethod": "fullPayment",
      "mark": "MDEwNDYwNDA2MDAwMzAxMDIxTjRONTdSRENVRFZUTA==",
      "markAccepted": false,
      "permission": {
        "uuid": "00000000-0000-0000-0000-000000000001",
        "time": "2026-07-05T10:00:00"
      }
    },
    { "type": "text", "text": " " },
    { "type": "text", "text": "Код продавца 42" }
  ],
  "payments": [
    { "type": "electronically", "sum": 200.00 }
  ],
  "total": 200.00
}
```

---

## 6. Режимы печати

| Режим | Поведение |
|-------|-----------|
| обычный | один `processJson(sell\|sellReturn)` |
| `testReceiptMode` | проверки марок выполняются; фискализация не делается — `cancelReceipt`, HTTP успех с пояснением |
| `emulation` | мок-ответы без живой ККТ |

---

## 7. История (Штрих → АТОЛ)

| Что | Было (Штрих) | Стало (АТОЛ) |
|-----|--------------|--------------|
| Печать чека | пошагово: CheckType → FNOperation → PrintString → FNCloseCheckEx | один `processJson` |
| Марки в чеке | FNCheckItemBarcode / FNSendItemBarcode / FNAcceptMarkingCode | `imcParams` после цикла валидации АТОЛ |
| РР | `FNSendTagOperation` 1262–1265 | `industryInfo` в JSON чека |
| Текст | `PrintString` | текстовые `items[]` в JSON АТОЛ |

Внешний контракт клиента (`type`, `items`, `payments`, `mark` base64, `permission`) **сохранён**.
