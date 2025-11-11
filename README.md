# Convert Excel to JSON

Утилита на Node.js + TypeScript для конвертации Excel-файлов в JSON. Скрипт поддерживает выбор отдельных листов, настройку значения по умолчанию для пустых ячеек, форматирование вывода и может использоваться как библиотека.

## Подготовка

```bash
npm install
npm run build
```

## Использование

```bash
node dist/index.js --input ./data.xlsx --print --output ./output.json
```

Доступные опции:

- `-i, --input <path>` — путь к Excel-файлу (обязательно)
- `-o, --output <path>` — путь для сохранения JSON (обязательно)
- `-s, --sheets <names>` — список листов через запятую (по умолчанию все листы)
- `--defval <value>` — значение по умолчанию для пустых ячеек (`null`, `true`, `1`, `текст`)
- `--indent <number>` — количество пробелов при форматировании JSON (по умолчанию `2`)
- `--compact` — сохранить JSON без форматирования
- `--print` — вывести результат конвертации в консоль
- `--silent` — не выводить сообщения об успешном сохранении
- `-h, --help` — показать справку

## Сравнение двух JSON по полю "ФИО"

```bash
node dist/compare.js --left ./output.json --right ./output1.json --field "ФИО"
```

Дополнительные опции:

- `--sheets <names>` — сравнивать только указанные листы
- `--save <path>` — сохранить отчет (`leftOnly`, `rightOnly`, `common`) в JSON-файл
- `--field <name>` — поменять поле для сравнения (по умолчанию `ФИО`)
- `-h, --help` — показать справку

> Скобочные пометки в конце значения (например, `Мария Иванова (0,75)`) игнорируются при сравнении.

## Использование в виде библиотеки

```ts
import {
	convertExcelToJsonAndSave,
	excelToJson,
	compareJsonFiles,
} from 'convet-excel-json/dist/index'

await convertExcelToJsonAndSave('./data.xlsx', './output.json', {
	sheets: ['Sheet1'],
	defaultValue: '',
	indent: 4,
	print: true,
})

const data = excelToJson('./data.xlsx', { defaultValue: null })

const diff = await compareJsonFiles('./output.json', './output1.json', {
	field: 'ФИО',
})
```
