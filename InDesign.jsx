// Открываем TXT файл с именами в формате UTF-16
var file = File.openDialog("Выберите TXT файл с именами");
if (!file) {
    alert("TXT файл не выбран. Процесс прекращен.");
    exit();
}

file.encoding = "UTF-16";
file.open("r");
var fileNames = [];
while (!file.eof) {
    var line = file.readln();
    if (line !== null && line !== "") {
        line = line.replace(/^\s+|\s+$/g, ''); // Удаляем начальные и конечные пробелы
        if (line.length > 0) {
            fileNames.push(line);
        }
    }
}
file.close();

// Открываем документ InDesign
var doc = app.documents.length > 0 ? app.activeDocument : null;
if (!doc) {
    alert("Нет открытого документа. Процесс прекращен.");
    exit();
}

// Запрашиваем у пользователя количество страниц для экспорта
var numPages = parseInt(prompt("Введите количество страниц для экспорта в один PDF файл:", "1"), 10);
if (isNaN(numPages) || numPages < 1) {
    alert("Некорректное количество страниц. Процесс прекращен.");
    exit();
}

// Спрашиваем у пользователя, нужно ли переводить текст в кривые
var convertToOutlines = confirm("Переводить текст в кривые перед экспортом?");

// Проверяем количество страниц и имен
var requiredFileCount = Math.ceil(doc.pages.length / numPages);
if (requiredFileCount > fileNames.length) {
    alert("Количество страниц больше количества имен в TXT файле. Процесс прекращен.");
    exit();
} else if (requiredFileCount < fileNames.length) {
    alert("Количество имен в TXT файле больше, чем требуется для экспорта страниц. Процесс прекращен.");
    exit();
}

// Выбираем папку для сохранения PDF файлов
var destFolder = Folder.selectDialog("Выберите папку для сохранения PDF");

if (destFolder) {
    // Создаем папки для каждого артикула и сохраняем PDF файлы
    for (var i = 0; i < doc.pages.length; i += numPages) {
        var pageName = fileNames[Math.floor(i / numPages)];
        var parts = pageName.split('_');
        var article = parts[0]; // Предполагаем, что артикул является первой частью имени файла
        var articleFolder = new Folder(destFolder + "/" + article);
        if (!articleFolder.exists) {
            articleFolder.create();
        }

        var exportPath = articleFolder + "/" + pageName + ".pdf";
        var exportFile = new File(exportPath);
        if (exportFile.exists) {
            exportFile.remove(); // Удаляем предыдущий файл, если существует
        }

        // Устанавливаем диапазон страниц для экспорта
        var pageRange = [];
        for (var j = 0; j < numPages; j++) {
            if (i + j < doc.pages.length) {
                pageRange.push((i + j + 1).toString());
            }
        }
        app.pdfExportPreferences.pageRange = pageRange.join(",");
        
        // Переводим текст в кривые, если выбрано
        if (convertToOutlines) {
            for (var p = 0; p < doc.pages.length; p++) {
                var pageItems = doc.pages[p].allPageItems;
                for (var q = 0; q < pageItems.length; q++) {
                    var item = pageItems[q];
                    if (item.constructor.name === "TextFrame") {
                        item.createOutlines();
                    }
                }
            }
        }

        doc.exportFile(ExportFormat.PDF_TYPE, exportFile);
    }
    alert("Экспорт завершен. Файлы сохранены в выбранных папках.");
} else {
    alert("Вы не выбрали папку для экспорта. Процесс прекращен.");
}
