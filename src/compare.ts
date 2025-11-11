import { readFile, writeFile } from 'fs/promises'
import { resolve } from 'path'

interface JsonRow {
	[key: string]: unknown
}

type JsonSource = JsonRow[] | Record<string, JsonRow[]>

interface CompareOptions {
	field?: string
	sheets?: string[]
	save?: string
}

interface CliArguments extends CompareOptions {
	left?: string
	right?: string
	help?: boolean
}

interface CompareResult {
	leftOnly: string[]
	rightOnly: string[]
	common: string[]
}

const DEFAULT_FIELD = 'ФИО'

const getErrorMessage = (error: unknown): string => {
	if (error instanceof Error) {
		return error.message
	}
	return String(error)
}

const parseJsonFile = async (filePath: string): Promise<JsonSource> => {
	const absolutePath = resolve(filePath)
	try {
		const content = await readFile(absolutePath, 'utf-8')
		return JSON.parse(content) as JsonSource
	} catch (error) {
		throw new Error(
			`Не удалось прочитать файл "${absolutePath}": ${getErrorMessage(error)}`
		)
	}
}

const castToRecordArray = (value: JsonSource, sheets?: string[]): JsonRow[] => {
	if (Array.isArray(value)) {
		return value
	}

	const sheetFilter = sheets && sheets.length > 0 ? new Set(sheets) : undefined

	return Object.entries(value)
		.filter(([sheetName]) => !sheetFilter || sheetFilter.has(sheetName))
		.flatMap(([, rows]) => rows)
}

const normalizeFieldValue = (value: unknown): string | undefined => {
	if (typeof value !== 'string') {
		return undefined
	}

	let result = value.trim()
	if (result.length === 0) {
		return undefined
	}

	const parentheticalPattern = /\s*\([^()]*\)\s*$/
	while (parentheticalPattern.test(result)) {
		result = result.replace(parentheticalPattern, '').trim()
	}

	return result.length > 0 ? result : undefined
}

const collectFieldValues = (
	source: JsonSource,
	field: string,
	sheets?: string[]
): Set<string> => {
	const rows = castToRecordArray(source, sheets)
	const values = new Set<string>()

	rows.forEach(row => {
		const normalized = normalizeFieldValue(row[field])
		if (normalized) {
			values.add(normalized)
		}
	})

	return values
}

const compareValueSets = (
	left: Set<string>,
	right: Set<string>
): CompareResult => {
	const leftOnly: string[] = []
	const rightOnly: string[] = []
	const common: string[] = []

	left.forEach(value => {
		if (right.has(value)) {
			common.push(value)
		} else {
			leftOnly.push(value)
		}
	})

	right.forEach(value => {
		if (!left.has(value)) {
			rightOnly.push(value)
		}
	})

	leftOnly.sort()
	rightOnly.sort()
	common.sort()

	return { leftOnly, rightOnly, common }
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
			case '--left':
				args.left = argv[++index]
				if (!args.left) {
					throw new Error('Опция "--left" требует указания пути к файлу')
				}
				break
			case '--right':
				args.right = argv[++index]
				if (!args.right) {
					throw new Error('Опция "--right" требует указания пути к файлу')
				}
				break
			case '--field':
				args.field = argv[++index]
				if (!args.field) {
					throw new Error(
						'Опция "--field" требует указания названия поля для сравнения'
					)
				}
				break
			case '--sheets':
				args.sheets = argv[++index]?.split(',').map(sheet => sheet.trim())
				if (!args.sheets || args.sheets.some(sheet => sheet.length === 0)) {
					throw new Error(
						'Опция "--sheets" требует указания списка имен листов через запятую'
					)
				}
				break
			case '--save':
				args.save = argv[++index]
				if (!args.save) {
					throw new Error(
						'Опция "--save" требует указания файла для сохранения отчета'
					)
				}
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
			'Использование: node dist/compare.js --left <left.json> --right <right.json> [опции]',
			'',
			'Опции:',
			'      --field <name>      Название поля для сравнения (по умолчанию "ФИО")',
			'      --sheets <names>    Список листов через запятую (по умолчанию все листы)',
			'      --save <path>       Сохранить отчет в JSON-файл',
			'  -h, --help              Показать эту справку',
		].join('\n')
	)
}

const formatReport = (
	result: CompareResult,
	leftPath: string,
	rightPath: string,
	field: string
): string => {
	const header = [
		`Сравнение по полю "${field}"`,
		`Левый файл: ${resolve(leftPath)}`,
		`Правый файл: ${resolve(rightPath)}`,
		'',
		`Совпадающих записей: ${result.common.length}`,
		`Только в левом файле: ${result.leftOnly.length}`,
		`Только в правом файле: ${result.rightOnly.length}`,
	].join('\n')

	const sections = [
		header,
		result.leftOnly.length
			? ['\nТолько в левом файле:', ...result.leftOnly].join('\n- ')
			: '',
		result.rightOnly.length
			? ['\nТолько в правом файле:', ...result.rightOnly].join('\n- ')
			: '',
	].filter(Boolean)

	return sections.join('\n')
}

const compareJsonFiles = async (
	leftPath: string,
	rightPath: string,
	options: CompareOptions = {}
): Promise<CompareResult> => {
	const field = options.field ?? DEFAULT_FIELD

	const [leftData, rightData] = await Promise.all([
		parseJsonFile(leftPath),
		parseJsonFile(rightPath),
	])

	const leftValues = collectFieldValues(leftData, field, options.sheets)
	const rightValues = collectFieldValues(rightData, field, options.sheets)

	return compareValueSets(leftValues, rightValues)
}

const runFromCli = async (): Promise<void> => {
	let args: CliArguments
	try {
		args = parseCliArguments(process.argv.slice(2))
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

	if (!args.left || !args.right) {
		console.error('Необходимо указать параметры "--left" и "--right"')
		printUsage()
		process.exitCode = 1
		return
	}

	try {
		const result = await compareJsonFiles(args.left, args.right, {
			field: args.field,
			sheets: args.sheets,
		})

		const report = formatReport(
			result,
			args.left,
			args.right,
			args.field ?? DEFAULT_FIELD
		)
		console.info(report)

		if (args.save) {
			const absolutePath = resolve(args.save)
			await writeFile(
				absolutePath,
				`${JSON.stringify(result, null, 2)}\n`,
				'utf-8'
			)
			console.info(`Отчет сохранен в файл: ${absolutePath}`)
		}
	} catch (error) {
		console.error(`Ошибка: ${getErrorMessage(error)}`)
		process.exitCode = 1
	}
}

if (require.main === module) {
	void runFromCli()
}

export { compareJsonFiles, runFromCli, parseCliArguments, printUsage }
