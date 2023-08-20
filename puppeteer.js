const fs = require("fs");
const puppeteer = require("puppeteer");
const prompt = require("prompt");
// const readline = require("readline").createInterface({
// 	input: process.stdin,
// 	output: process.stdout,
// });
// const util = require("util");
// const question = util.promisify(readline.question).bind(readline);

const excludedValueNames = ["MDRD", "eGFR", "FIB", "SERUM INDEX", "LIPEMIA", "ICTERUS", "HEMOLYSİS"];
let dateRegex = /^(\d{2}\.\d{2}\.\d{2})/;
let valueRegex = /^(((-|<|>)?\d+,?\d*)?\s?((NEGATİF\(-\))|(POZİTİF\(\+\)))?)/;

function sleep(ms) {
	return new Promise((resolve) => {
		setTimeout(resolve, ms);
	});
}
async function findAsync(array, predicate) {
	// find async sequential
	for (const t of array) {
		if (await predicate(t)) {
			return t;
		}
	}
	return undefined;
}

/**
 *
 * @param {puppeteer.Browser} browser
 * @param {string} username
 * @param {string} password
 * @returns {[puppeteer.Page, string[], puppeteer.ElementHandle[]]} [page, patientNames, patients]
 */
async function authenticateAndListPatients(browser, username, password) {
	const page = await browser.newPage();
	await page.setViewport({
		height: 100000,
		// height: 1080,
		width: 800,
	});

	// Navigate the page to a URL
	await page.goto("https://onlinehbys.kocaeli.edu.tr:20108/nucleus-mobile/login");

	// authenticate
	// long way
	// const usernameFieldSelector =
	// 	"#ironform > form > vaadin-vertical-layout > vaadin-text-field >>>> #vaadin-text-field-input-0 > slot:nth-child(2) > input";
	// await page.waitForSelector(usernameFieldSelector);
	// await page.type(usernameFieldSelector, username);
	// const passwordFieldSelector =
	// 	"#ironform > form > vaadin-vertical-layout > vaadin-password-field >>>> #vaadin-password-field-input-1 > slot:nth-child(2) > input[type=password]";
	// await page.waitForSelector(passwordFieldSelector);
	// await page.type(passwordFieldSelector, password);
	// const loginButtonSelector = "#submitbutton >>>> #button";
	// await page.waitForSelector(loginButtonSelector);
	// await page.click(loginButtonSelector);

	// short way
	await page
		.locator("#ironform > form > vaadin-vertical-layout > vaadin-text-field >>>> #vaadin-text-field-input-0 > slot:nth-child(2) > input")
		.fill(username);
	await page
		.locator(
			"#ironform > form > vaadin-vertical-layout > vaadin-password-field >>>> #vaadin-password-field-input-1 > slot:nth-child(2) > input[type=password]"
		)
		.fill(password);
	await page.locator("#submitbutton >>>> #button").click();

	// click Hasta İşlemleri
	await page
		.locator("body > vaadin-app-layout > div > div > div > div.view-frame__content > div > div > div > div.mobilmenu__group > div > div > span")
		.click();

	// click Hasta Sorgula
	await page
		.locator(
			"#overlay > flow-component-renderer > div > div.view-frame-dialog__footer > vaadin-horizontal-layout > vaadin-button:nth-child(1) >>>> #button"
		)
		.click();

	// wait for Hasta Sorgulama title
	await page.waitForSelector(
		"body > vaadin-app-layout > vaadin-horizontal-layout > vaadin-horizontal-layout:nth-child(2) > label ::-p-text(Hasta Sorgulama)"
	);

	let patients = await page.$$("body > vaadin-app-layout > div > div > div > div.view-frame__content > vaadin-grid > vaadin-grid-cell-content");
	patients = patients.slice(3);
	const patientNames = [];
	for (const patient of patients) {
		const patientLabel = await patient.$("label");
		const patientName = await page.evaluate((el) => el.innerText, patientLabel);
		patientNames.push(patientName);
	}
	console.log(patientNames);
	console.log("Total patient count: " + patients.length);

	return [page, patientNames, patients];
}

/**
 *
 * @param {puppeteer.Page} page
 * @param {puppeteer.ElementHandle[]} patients
 * @param {string} patienName
 * @param {puppeteer.Browser} browser
 * @returns {[{}, puppeteer.Page]} [patientData, page]
 */
async function getPatientData(browser, patienName) {
	let myBrowser = browser;
	if (!myBrowser) myBrowser = await puppeteer.launch({ headless: "new" });
	const page = await myBrowser.newPage();
	await page.setViewport({
		height: 100000,
		// height: 1080,
		width: 800,
	});
	// Navigate the page to a URL
	await page.goto("https://onlinehbys.kocaeli.edu.tr:20108/nucleus-mobile");

	// click Hasta İşlemleri
	await page
		.locator("body > vaadin-app-layout > div > div > div > div.view-frame__content > div > div > div > div.mobilmenu__group > div > div > span")
		.click();

	// click Hasta Sorgula
	await page
		.locator(
			"#overlay > flow-component-renderer > div > div.view-frame-dialog__footer > vaadin-horizontal-layout > vaadin-button:nth-child(1) >>>> #button"
		)
		.click();

	// wait for Hasta Sorgulama title
	await page.waitForSelector(
		"body > vaadin-app-layout > vaadin-horizontal-layout > vaadin-horizontal-layout:nth-child(2) > label ::-p-text(Hasta Sorgulama)"
	);

	let patients = await page.$$("body > vaadin-app-layout > div > div > div > div.view-frame__content > vaadin-grid > vaadin-grid-cell-content");
	patients = patients.slice(3);

	const patient = await findAsync(patients, async (item) => (await page.evaluate((el) => el.innerText, item)).includes(patienName));
	// console.log(await page.evaluate((el) => el.innerText, patient));
	await sleep(1000);
	await patient.click();
	// const titleElement = await page.$("body > vaadin-app-layout > vaadin-horizontal-layout > vaadin-horizontal-layout:nth-child(2) > label")
	await sleep(1000);

	await Promise.race([
		(async () => {
			await page.waitForSelector("#overlay > flow-component-renderer > div > header > h2 ::-p-text(Başvuru Seçim)");
			await page
				.locator(
					"#overlay > flow-component-renderer > div > div.view-frame-dialog__footer > vaadin-horizontal-layout > div.back-button > vaadin-button >>>> #button"
				)
				.click();
			await sleep(1000);
			// click Tetkik Sonuç
			await page
				.locator(
					"body > vaadin-app-layout > div > div > div > div.view-frame__content > div > div > div > div.mobilmenu__group > div > div > span ::-p-text(Tetkik Sonuç)"
				)
				.click();
		})(),

		// click Tetkik Sonuç
		page
			.locator(
				"body > vaadin-app-layout > div > div > div > div.view-frame__content > div > div > div > div.mobilmenu__group > div > div > span ::-p-text(Tetkik Sonuç)"
			)
			.click(),
	]);

	await page.waitForSelector(
		"body > vaadin-app-layout > vaadin-horizontal-layout > vaadin-horizontal-layout:nth-child(2) > label ::-p-text(Tetkik Sonuç)"
	);
	await sleep(10000);

	let tableItems = await page.$$(
		"body > vaadin-app-layout > div > div > div > div.view-frame__wrapper > div.view-frame__content > vaadin-vertical-layout > div > vaadin-vertical-layout > vaadin-grid > vaadin-grid-cell-content"
	);
	const patientLabResults = {};
	tableItems = tableItems.slice(3);
	for (const i in tableItems) {
		const text = await page.evaluate((el) => el.innerText, tableItems[i]);
		// console.log(text);
		if (excludedValueNames.some((excludedValue) => text.includes(excludedValue))) continue;
		let rows = text.split("\n");
		if (rows.length === 1) continue;
		if (rows[0].includes("Teknik Onaylı")) {
			rows = rows.slice(1);
		}
		if (
			rows[0].includes("Hemogram") ||
			rows[0].includes("Protrombin Zamanı") ||
			rows[0].includes("Kan gazı") ||
			rows[0].includes("Tam İdrar") ||
			rows[0].includes("Gastrointestinal Panel")
		) {
			// console.log(rows[0], "results:");
			let dateResult = dateRegex.exec(rows[1]);
			if (!dateResult) continue; // if date is not found, continue
			if (!(dateResult[0] in patientLabResults)) patientLabResults[dateResult[0]] = {};
			let dateObj = patientLabResults[dateResult[0]];
			let rowElements = await tableItems[i].$$("vaadin-vertical-layout > vaadin-vertical-layout > div");
			for (const rowIndex in rowElements) {
				const rowElement = rowElements[rowIndex];
				const rowElementText = await page.evaluate((el) => el.innerText, rowElement);
				const splitedRowText = rowElementText.split("\n");
				// console.log(rowIndex + "-" + rowElementText);
				let valueResult = valueRegex.exec(splitedRowText[1]);
				if (!valueResult) continue;
				if (!dateObj[splitedRowText[0]]) dateObj[splitedRowText[0]] = valueResult[0].trim();
			}
			// console.log(dateObj);
		} else {
			// Biyokimya veya elisa
			let dateResult = dateRegex.exec(rows.at(-1));
			if (!dateResult) continue; // if date is not found, continue
			if (!(dateResult[0] in patientLabResults)) patientLabResults[dateResult[0]] = {};
			let dateObj = patientLabResults[dateResult[0]];
			let valueResult = valueRegex.exec(rows[1]);
			if (!valueResult) continue;
			if (!dateObj[rows[0]]) dateObj[rows[0]] = valueResult[0].trim();
		}
		// console.log(i + "-" + text.split("\n").at(-1));
	}

	return [patientLabResults, page];
}

module.exports = async (fileName) => {
	const startTime = new Date();
	const fileNameWithExtension = fileName + ".json";
	// Get user account
	// const username = await question("Kullanıcı adınızı girin: ");
	// const password = await question("Şifrenizi girin: ");
	const { username, password } = await prompt.get([
		{ name: "username", description: "Kullanıcı adı", required: true },
		{ name: "password", description: "Şifre", required: true, hidden: true },
	]);
	// Launch the browser and open a new blank page
	const browser = await puppeteer.launch({ headless: "new" });
	console.log("Listing patients...");
	const [initialPage, patientNames] = await authenticateAndListPatients(browser, username, password);
	initialPage.close();

	const allData = {};
	const patientTimes = [];
	// parallel
	// ! EXTREMELY DANGEROUS WITH BROWSER FOR EACH PATIENT
	// await Promise.all(
	// 	patientNames.map(async (patientName) => {
	// 		console.log("Scraping for: ", patientName);
	// 		const patientStartTime = new Date();
	// 		// const [page, patienNames, patients] = await authenticateAndListPatients(browser, username, password);
	// 		// const patientData = await getPatientData(page, patients, patientName);
	// 		const [patientData, page] = await getPatientData(null, patientName);
	// 		allData[patientName] = patientData;
	// 		await page.close();
	// 		const patientProcessTime = (new Date() - patientStartTime) / 1000;
	// 		console.log("Patient processed in " + patientProcessTime + " seconds!");
	// 		patientTimes.push(patientProcessTime);
	// 	})
	// );
	// fs.writeFile(fileNameWithExtension, JSON.stringify(allData), "utf8", function (err) {
	// 	if (err) throw err;
	// 	console.log(patientName, "processed to", fileNameWithExtension);
	// });
	// sequential
	for (const patientName of patientNames) {
		console.log("Scraping for: ", patientName);
		const patientStartTime = new Date();
		// const [page, patienNames, patients] = await authenticateAndListPatients(browser, username, password);
		// const patientData = await getPatientData(page, patients, patientName);
		const [patientData, page] = await getPatientData(browser, patientName);
		allData[patientName] = patientData;
		await page.close();
		const patientProcessTime = (new Date() - patientStartTime) / 1000;
		console.log("Patient processed in " + patientProcessTime + " seconds!");
		patientTimes.push(patientProcessTime);
		await fs.promises.writeFile(fileNameWithExtension, JSON.stringify(allData), "utf8");
		console.log(patientName, "is successfuly scraped.");
	}
	console.log("Executed in " + (new Date() - startTime) / 1000 + " seconds!");
	console.log("Mean patient process time: " + patientTimes.reduce((a, b) => a + b, 0) / patientTimes.length + " seconds.");
	console.log("Max patient process time: " + Math.max(...patientTimes) + " seconds.");
	console.log("Min patient process time: " + Math.min(...patientTimes) + " seconds.");
	await browser.close();
};
