const docx = require("./docx");
const prompt = require("prompt");
const path = require("path");

prompt.get(["fileName"], function (err, result) {
	if (err) {
		return onErr(err);
	}
	docx(path.join(__dirname, result.fileName));
});
