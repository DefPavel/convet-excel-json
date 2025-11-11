import * as XLSX from 'xlsx'
import { constants } from 'fs'
import { access, mkdir, writeFile } from 'fs/promises'
import { dirname, resolve } from 'path'

interface ExcelRow {
	[key: string]: unknown
}

export type WorkbookJson = Record<string, ExcelRow[]>

interface ExcelToJsonOptions {
	sheets?: string[]
	defaultValue?: unknown
}

interface SaveJsonOptions {
	indent?: number | false
}

interface ConvertOptions extends ExcelToJsonOptions, SaveJsonOptions {
	silent?: boolean
	print?: boolean
}

interface CliArguments extends ConvertOptions {
	input?: string
	output?: string
	help?: boolean
}

const getErrorMessage = (error: unknown): string => {
	if (error instanceof Error) {
		return error.message
	}
	return String(error)
}

const ensureFileReadable = async (filePath: string): Promise<void> => {
	try {
		await access(filePath, constants.R_OK)
	} catch (error) {
		throw new Error(
			`Не удалось получить доступ к файлу "${filePath}": ${getErrorMessage(
				error
			)}`
		)
	}
}

const ensureDirectoryExists = async (filePath: string): Promise<void> => {
	const targetDir = dirname(filePath)
	try {
		await mkdir(targetDir, { recursive: true })
	} catch (error) {
		throw new Error(
			`Не удалось создать директорию "${targetDir}": ${getErrorMessage(error)}`
		)
	}
}

const normalizeSheetFilter = (sheets?: string[]): Set<string> | undefined => {
	if (!sheets || sheets.length === 0) {
		return undefined
	}

	const normalized = sheets
		.map(sheet => sheet.trim())
		.filter(sheet => sheet.length > 0)

	return normalized.length > 0 ? new Set(normalized) : undefined
}

const excelToJson = (
	filePath: string,
	options: ExcelToJsonOptions = {}
): WorkbookJson => {
	const workbook = XLSX.readFile(filePath)
	const sheetFilter = normalizeSheetFilter(options.sheets)

	const sheetNames = workbook.SheetNames.filter(sheetName => {
		return !sheetFilter || sheetFilter.has(sheetName)
	})

	if (sheetNames.length === 0) {
		const message = sheetFilter
			? `В файле "${filePath}" не найдены листы: ${[...sheetFilter].join(', ')}`
			: `Файл "${filePath}" не содержит листов`
		throw new Error(message)
	}

	return sheetNames.reduce<WorkbookJson>((result, sheetName) => {
		const worksheet = workbook.Sheets[sheetName]
		if (!worksheet) {
			return result
		}

		result[sheetName] = XLSX.utils.sheet_to_json(worksheet, {
			defval: options.defaultValue ?? null,
		})

		return result
	}, {})
}

const saveJsonToFile = async (
	data: WorkbookJson,
	outputFilePath: string,
	options: SaveJsonOptions = {}
): Promise<void> => {
	await ensureDirectoryExists(outputFilePath)

	const indent = options.indent === false ? undefined : options.indent ?? 2
	const jsonData = indent
		? JSON.stringify(data, null, indent)
		: JSON.stringify(data)

	try {
		await writeFile(outputFilePath, `${jsonData}\n`, 'utf-8')
	} catch (error) {
		throw new Error(`Ошибка при записи файла: ${getErrorMessage(error)}`)
	}
}

const convertExcelToJsonAndSave = async (
	inputFilePath: string,
	outputFilePath: string,
	options: ConvertOptions = {}
): Promise<void> => {
	const absoluteInputPath = resolve(inputFilePath)
	const absoluteOutputPath = resolve(outputFilePath)

	await ensureFileReadable(absoluteInputPath)

	const data = excelToJson(absoluteInputPath, {
		sheets: options.sheets,
		defaultValue: options.defaultValue,
	})

	await saveJsonToFile(data, absoluteOutputPath, {
		indent: options.indent,
	})

	if (options.print) {
		const indent = options.indent === false ? undefined : options.indent ?? 2
		const payload = indent
			? JSON.stringify(data, null, indent)
			: JSON.stringify(data)
		console.info(payload)
	}

	if (!options.silent) {
		console.info(`JSON успешно сохранён в файл: ${absoluteOutputPath}`)
	}
}

const coerceValue = (raw: string): unknown => {
	const lower = raw.toLowerCase()
	if (lower === 'null') {
		return null
	}
	if (lower === 'true') {
		return true
	}
	if (lower === 'false') {
		return false
	}

	const numeric = Number(raw)
	if (!Number.isNaN(numeric) && raw.trim() !== '') {
		return numeric
	}

	return raw
}

const parseCliArguments = (argv: string[]): CliArguments => {
	const args: CliArguments = {}

	for (let index = 0; index < argv.length; index += 1) {
		const token = argv[index]

		switch (token) {
			case '--help':
			case '-h':
				args.help = true
				break
			case '--input':
			case '-i': {
				const value = argv[++index]
				if (!value) {
					throw new Error('Опция "--input" требует указания пути к файлу')
				}
				args.input = value
				break
			}
			case '--output':
			case '-o': {
				const value = argv[++index]
				if (!value) {
					throw new Error('Опция "--output" требует указания пути к файлу')
				}
				args.output = value
				break
			}
			case '--sheets':
			case '-s': {
				const value = argv[++index]
				if (!value) {
					throw new Error(
						'Опция "--sheets" требует указания списка имен листов'
					)
				}
				args.sheets = value.split(',')
				break
			}
			case '--defval': {
				const value = argv[++index]
				if (value === undefined) {
					throw new Error(
						'Опция "--defval" требует указания значения по умолчанию'
					)
				}
				args.defaultValue = coerceValue(value)
				break
			}
			case '--indent':
			case '--pretty': {
				const value = argv[++index]
				if (value === undefined) {
					throw new Error('Опция "--indent" требует указания числа пробелов')
				}
				const indent = Number(value)
				if (!Number.isFinite(indent) || indent < 0) {
					throw new Error(
						`Опция "--indent" должна быть неотрицательным числом, получено "${value}"`
					)
				}
				args.indent = indent
				break
			}
			case '--compact':
				args.indent = false
				break
			case '--silent':
				args.silent = true
				break
			case '--print':
				args.print = true
				break
			default:
				throw new Error(`Неизвестный аргумент командной строки: "${token}"`)
		}
	}

	return args
}

const printUsage = (): void => {
	console.info(
		[
			'Использование: node dist/index.js --input <excel.xlsx> --output <data.json> [опции]',
			'',
			'Опции:',
			'  -i, --input <path>      Путь к исходному Excel-файлу (обязательно)',
			'  -o, --output <path>     Путь для сохранения JSON-файла (обязательно)',
			'  -s, --sheets <names>    Список листов через запятую (по умолчанию все листы)',
			'      --defval <value>    Значение по умолчанию для пустых ячеек (null, true, 42, ...)',
			'      --indent <number>   Количество пробелов для форматирования JSON (по умолчанию 2)',
			'      --compact           Сохранить JSON без форматирования',
			'      --print             Вывести результат конвертации в консоль',
			'      --silent            Не выводить информацию о результате',
			'  -h, --help              Показать эту справку',
		].join('\n')
	)
}

const runFromCli = async (): Promise<void> => {
	const argv = process.argv.slice(2)

	let args: CliArguments
	try {
		args = parseCliArguments(argv)
	} catch (error) {
		console.error(getErrorMessage(error))
		printUsage()
		process.exitCode = 1
		return
	}

	if (args.help) {
		printUsage()
		return
	}

	if (!args.input || !args.output) {
		console.error('Необходимо указать параметры "--input" и "--output"')
		printUsage()
		process.exitCode = 1
		return
	}

	try {
		await convertExcelToJsonAndSave(args.input, args.output, {
			sheets: args.sheets,
			defaultValue: args.defaultValue,
			indent: args.indent,
			silent: args.silent,
			print: args.print,
		})
	} catch (error) {
		console.error(`Ошибка: ${getErrorMessage(error)}`)
		process.exitCode = 1
	}
}

if (require.main === module) {
	void runFromCli()
}

export {
	convertExcelToJsonAndSave,
	excelToJson,
	parseCliArguments,
	printUsage,
	runFromCli,
}

export { compareJsonFiles } from './compare'
