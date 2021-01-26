eel.expose(dontFindFileName);
function dontFindFileName() {
	alert('Эксель файл не найден в папке');
}
eel.expose(toManyExcelFiles);
function toManyExcelFiles() {
	alert('В корневой папке не может быть более одного файла XLSX!')
}