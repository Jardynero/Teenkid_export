window.onload=function() {
	let button = document.getElementById('button');
	let progress = document.getElementById('progress');
	let firstStep = async function() {
		let res = await eel.excel_file_search_step_1()();
		document.getElementById('log').innerHTML = 'Excel файл найден и опознан!';
		progress.value = 1;
	};
	window.onload = button.addEventListener('click', firstStep);
};
