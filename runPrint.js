const print = require("./print");
const prompt = require("prompt");
const path = require("path");

prompt.get(["fileName"], function (err, result) {
	if (err) {
		return onErr(err);
	}
	print(path.join(__dirname, result.fileName));
});
