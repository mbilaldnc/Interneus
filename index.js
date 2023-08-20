const puppeteer = require("./puppeteer");
const docx = require("./docx");
const { exec } = require("child_process");

(async () => {
	const fileName = new Date().toLocaleString().replaceAll(":", ".").replaceAll(" ", "_");
	console.log("fileName: ", fileName);
	await puppeteer(fileName);
	await docx(fileName);
	await exec(`node print.js "${fileName}.docx"`);
	// await docx("20.08.2023 04.36.05");
})();
