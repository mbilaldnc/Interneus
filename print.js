const cmd = require("node-cmd");
// const path = require("path");
// const fileName = process.argv[2];
// const absoluteFilePath = path.resolve(__dirname, fileName);

function getDataLogger(prefix) {
	let data_line = "";

	return function (data) {
		data_line += data;
		if (data_line[data_line.length - 1] === "\n") {
			console.log(`[${prefix}]`, data_line);
		}
	};
}

module.exports = (absoluteFileName) => {
	console.log(`Printing ${absoluteFileName}.docx`);
	const proc = cmd.run(
		`"C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE" /q /n "${absoluteFileName}.docx" /mFilePrintDefault /mFileCloseOrExit`
	);

	proc.stdout.on("data", getDataLogger("stdout"));
	proc.stderr.on("data", getDataLogger("stderr"));
};
