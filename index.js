const puppeteer = require("./puppeteer");
const docx = require("./docx");

(async () => {
	const fileName = new Date().toLocaleString().replaceAll(":", ".");
	console.log("fileName: ", fileName);
	await puppeteer(fileName);
	await docx(fileName);
	// await docx("20.08.2023 04.36.05");
})();
