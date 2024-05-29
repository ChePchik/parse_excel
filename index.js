const XLSX = require("xlsx");
const fs = require("fs");

// Путь к вашему Excel-файлу
const filePath = "example.xlsx";

try {
	if (!fs.existsSync(filePath)) {
		console.error("Файл не найден. Проверьте путь к файлу.");
		throw new Error("Файл не найден. Проверьте путь к файлу.");
	}
	// Чтение файла
	const workbook = XLSX.readFile(filePath);
	console.log("🚀 ~ workbook:", workbook);
	// Выбираем первый лист из книги
	const sheetName = workbook.SheetNames[0];
	if (!sheetName) {
		console.error("Лист в файле не найден.");
		throw new Error("Лист в файле не найден.");
	}

	const sheet = workbook.Sheets[sheetName];
	// console.log("🚀 ~ sheet:", sheet);

	// Начинаем с первой строки
	let rowIndex = 2;
	// Массив для хранения преобразованных данных
	const typeQu = {
		"несколько ответов": "checkbox",
		"один ответ": "radio",
		"текстовой ответ": "text",
	};
	// Массив для хранения преобразованных данных
	const questions = [];

	// Регулярные выражения для проверок
	const allowedCharsRegex = /^[a-zA-Zа-яёА-ЯЁ0-9,.\-!?'" ]{1,60}$/;
	const allowedTypes = ["один ответ", "несколько ответов", "текстовой ответ"];
	const allowedGValuesRegex = /^[1-4](,[1-4]){0,3}$/;

	// Перебираем строки до тех пор, пока не достигнем пустой строки
	while (sheet["A" + rowIndex]) {
		const questionData = {
			type: "radio", // Значение по умолчанию, можно адаптировать под свои нужды
			question: "",
			answers: [],
		};
		let type_answers = null;

		// Получаем значения в диапазоне A1:H1
		for (let col = "A".charCodeAt(0); col <= "H".charCodeAt(0); col++) {
			const cellAddress = String.fromCharCode(col) + rowIndex;
			// const cellValue = sheet[cellAddress] ? sheet[cellAddress].v : "";
			const cellValue = sheet[cellAddress] ? String(sheet[cellAddress].v).trim() : "";

			// Проверки для столбцов A, C, D, E, F
			if (
				["A", "C"].includes(String.fromCharCode(col)) ||
				(["D", "E", "F"].includes(String.fromCharCode(col)) && cellValue !== "")
			) {
				if (!allowedCharsRegex.test(cellValue)) {
					console.error(
						`Недопустимые символы или длина в столбце ${String.fromCharCode(
							col,
						)} на строке ${rowIndex}.`,
					);
					throw new Error(
						`Недопустимые символы или длина в столбце ${String.fromCharCode(
							col,
						)} на строке ${rowIndex}.`,
					);
				}
			}
			// Проверка полноты данных: Убедимся, что важные поля не пусты
			if (
				(col === "A".charCodeAt(0) ||
					col === "B".charCodeAt(0) ||
					col === "C".charCodeAt(0) ||
					col === "G".charCodeAt(0)) &&
				!cellValue?.trim()
			) {
				console.error(
					`Пустое значение в критически важном столбце ${String.fromCharCode(
						col,
					)} на строке ${rowIndex}.`,
				);
				throw new Error(
					`Пустое значение в критически важном столбце ${String.fromCharCode(
						col,
					)} на строке ${rowIndex}.`,
				);
			}

			// Добавляем значения в соответствующие поля в объекте questionData
			switch (col) {
				case "A".charCodeAt(0):
					questionData.question = cellValue;
					break;
				case "B".charCodeAt(0):
					if (!allowedTypes.includes(cellValue.toLowerCase())) {
						console.error(`Недопустимое значение в столбце B на строке ${rowIndex}.`);
						throw new Error(`Недопустимое значение в столбце B на строке ${rowIndex}.`);
					}
					questionData.type = typeQu[cellValue.toLowerCase()]; // Преобразуем в нижний регистр для унификации
					type_answers = typeQu[cellValue.toLowerCase()];
					console.log("🚀 ~ type_answers:", type_answers);
					break;
				case "C".charCodeAt(0):
					if (type_answers == "text") {
						questionData.answers.push({
							text: cellValue,
							right: cellValue,
						});
						break;
					}

				case "D".charCodeAt(0):
					if (type_answers == "text") break;
				case "E".charCodeAt(0):
					if (type_answers == "text") break;
				case "F".charCodeAt(0):
					if (type_answers == "text") break;
					questionData.answers.push({
						text: cellValue,
						right: "false",
					});
					break;
				case "G".charCodeAt(0):
					if (type_answers == "text") break;
					if (!allowedGValuesRegex.test(cellValue)) {
						console.error(`Недопустимое значение в столбце G на строке ${rowIndex}.`);
						throw new Error(`Недопустимое значение в столбце G на строке ${rowIndex}.`);
					}
					let answerIndices = [];

					// Проверяем, является ли значение в поле G строкой
					if (typeof cellValue === "string") {
						// Разбиваем строку на массив индексов
						answerIndices = cellValue
							.split(",")
							.map((answerIndex) => parseInt(answerIndex.trim()) - 1);
					} else if (!isNaN(cellValue)) {
						// Если значение в поле G является числом (одиночный ответ)
						answerIndices = [parseInt(cellValue) - 1];
					}

					// Устанавливаем "true" для соответствующих индексов в массиве ответов
					answerIndices.forEach((answerIndex) => {
						if (questionData.answers[answerIndex]) {
							questionData.answers[answerIndex].right = "true";
						}
					});
					break;
			}
		}

		// Добавляем объект questionData в массив questions
		questions.push(questionData);

		// Переходим к следующей строке
		rowIndex++;
	}

	// Сохраняем в файл
	const outputData = { questions };
	const outputFilePath = "output.json";

	fs.writeFileSync(outputFilePath, JSON.stringify(outputData, null, 2));

	console.log(`Данные сохранены в файл: ${outputFilePath}`);
} catch (error) {
	console.error("Ошибка при чтении файла:", error);
	fs.close();
}
