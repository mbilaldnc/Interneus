const puppeteer = require("./puppeteer");
const docx = require("./docx");
const { exec } = require("child_process");
const path = require("path");

(async () => {
	const fileName = new Date().toLocaleString().replaceAll(":", ".").replaceAll(" ", "_");
	const absoluteFileName = path.resolve(__dirname, fileName);
	console.log(`
\x1b[32mIntorneus'a Hoş geldiniz!

\x1b[33mProgramın amacı; nucleus'taki hastaların laboratuvar sonuçlarını otomatik olarak toplayıp
word dosyası olarak çıktı almanızı sağlamaktır.
Belirtilen hastaların isimleri ve laboratuvar sonuçları hariç hiçbir veriyi saklamaz.
Hesap giriş bilgilerinizi depolamaz. Nucleus hariç başka bir uzak sunucuyla bağlantı kurmaz.

\x1b[36mŞu şekilde çalışır:
1-Girmiş olduğunuz hesap bilgileri ile nucleus'a giriş yapar.
2-En son seçmiş olduğunuz servis(ler)teki hastaların üstüne kullanıcı tıklamasını simüle eder.
3-Hastaların laboratuvar sonuçlarını sırayla belirtilen json dosyasında toplar.
4-Toplanan verileri kullanarak bir word dosyası oluşturur.
5-Oluşturulan word dosyasını varsayılan yazıcı ile yazdırır. (Varsayılan word konumu: C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE)

\x1b[35mİnt. Dr. Mehmet Bilal Danacı tarafından intörnler üzerindeki iş yükünü azaltmak amacıyla geliştirilmiştir.
İletişim: mbilaldnc@gmail.com
\x1b[90mHedef şablon olarak gastroenteroloji hasta sonuç kağıdı kullanılmıştır. Diğer bölümlerin sonuç kağıtlarını geliştirmeye yardımcı olmak için lütfen iletişime geçiniz.\x1b[0m
	`);
	// console.log("json file: ", absoluteFileName + ".json");
	await puppeteer(absoluteFileName);
	await docx(absoluteFileName);
	await exec(`node print.js "${fileName}.docx"`);
	// await docx("20.08.2023 04.36.05");
})();
