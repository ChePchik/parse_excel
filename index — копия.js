const XLSX = require("xlsx");
const fs = require("fs");

// Путь к вашему Excel-файлу
const filePath = "example.xlsx";

// Чтение файла
const workbook = XLSX.readFile(filePath);
// Выбираем первый лист из книги
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];
console.log("🚀 ~ sheet:", sheet);

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

// Перебираем строки до тех пор, пока не достигнем пустой строки
while (sheet["A" + rowIndex]) {
	const questionData = {
		type: "radiobox", // Значение по умолчанию, можно адаптировать под свои нужды
		question: "",
		answers: [],
	};
	let type_answers = null;

	// Получаем значения в диапазоне A1:H1
	for (let col = "A".charCodeAt(0); col <= "H".charCodeAt(0); col++) {
		const cellAddress = String.fromCharCode(col) + rowIndex;
		const cellValue = sheet[cellAddress] ? sheet[cellAddress].v : "";

		// Добавляем значения в соответствующие поля в объекте questionData
		switch (col) {
			case "A".charCodeAt(0):
				questionData.question = cellValue;
				break;
			case "B".charCodeAt(0):
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
