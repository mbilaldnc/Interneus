const docx = require("./docx");
const prompt = require("prompt");

prompt.get(["fileName"], function (err, result) {
	if (err) {
		return onErr(err);
	}
	docx(result.fileName);
});
