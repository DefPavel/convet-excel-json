import * as XLSX from 'xlsx'
import { writeFile } from 'fs/promises'
import { resolve } from 'path'

interface ExcelSheet {
	[key: string]: any
}
const getErrorMessage = (error: unknown): string => {
	if (error instanceof Error) {
		return error.message
	}
	return String(error) // Преобразуем любые другие типы в строку
}

const excelToJson = (filePath: string): Record<string, ExcelSheet[]> => {
	try {
		// Чтение Excel файла
		const workbook = XLSX.readFile(filePath)
		// Преобразуем все листы в объект с ключами - названиями листов
		const sheetsData: Record<string, ExcelSheet[]> = {}
		workbook.SheetNames.forEach(sheetName => {
			const worksheet = workbook.Sheets[sheetName]
			sheetsData[sheetName] = XLSX.utils.sheet_to_json(worksheet, {
				defval: null,
			})
		})
		return sheetsData
	} catch (error) {
		throw new Error(`Ошибка при чтении Excel файла: ${getErrorMessage(error)}`)
	}
}

const saveJsonToFile = async (
	data: Record<string, ExcelSheet[]>,
	outputFilePath: string
): Promise<void> => {
	try {
		const jsonData = JSON.stringify(data, null, 2)
		await writeFile(outputFilePath, jsonData, 'utf-8')
		console.info(`JSON успешно сохранён в файл: ${outputFilePath}`)
	} catch (error) {
		throw new Error(`Ошибка при записи файла: ${getErrorMessage(error)}`)
	}
}

// Основная функция для преобразования и сохранения Excel в JSON
const convertExcelToJsonAndSave = async (
	inputFilePath: string,
	outputFilePath: string
): Promise<void> => {
	try {
		const absoluteInputPath = resolve(inputFilePath)
		const absoluteOutputPath = resolve(outputFilePath)

		const data = excelToJson(absoluteInputPath)
		await saveJsonToFile(data, absoluteOutputPath)
	} catch (error) {
		console.error(`Ошибка: ${getErrorMessage(error)}`)
	}
}

// Пример использования
const inputFilePath = './data.xlsx'
const outputFilePath = './output.json'

convertExcelToJsonAndSave(inputFilePath, outputFilePath)
