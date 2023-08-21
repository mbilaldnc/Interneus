const puppeteer = require("./puppeteer");
const docx = require("./docx");
const print = require("./print");
const child_process = require("child_process");
// const { exec } = require("child_process");
const prompt = require("prompt");
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

\x1b[91mÖn hazırlık ve gereksinimler:
1-Programı kullanabilmek için bir nucleus hesabına ihtiyacınız var.
2-Programı kullanmadan önce nucleus hesabınıza internet tarayıcınızdan girerek
  giriş yapıp sırasıyla Hasta İşlemleri>Hasta Sorgula bastıktan sonra "Aktif Hastalar" sekmesinden 
  "Yatan" seçeneğini seçip sonuçlarını yazdırmak istediğiniz bölüm(ler)ü seçin.
  Bu işlem, hesabınıza bir sonraki girişinizde seçtiğiniz bölümdeki hastaların direkt olarak önünüzde
  seçili olarak çıkmasını sağlayacaktır. Program da bu seçili hastaların sonuçlarını toplayacaktır.
3-(Opsiyonel)Program için özel bir klasör oluşturup. Programı onun içerisinde çalıştırmanız önerilir.

\x1b[92mİnt. Dr. Mehmet Bilal Danacı tarafından intörnler üzerindeki iş yükünü azaltmak amacıyla geliştirilmiştir.
İletişim: mbilaldnc@gmail.com
\x1b[90mHedef şablon olarak gastroenteroloji hasta sonuç kağıdı kullanılmıştır. Diğer bölümlerin sonuç kağıtlarını geliştirmeye yardımcı olmak için lütfen iletişime geçiniz.\x1b[0m
	`);
	try {
		await puppeteer(absoluteFileName);
		await docx(absoluteFileName);
		console.log("Varsayılan yazıcıyla yazdırmaya devam etmek için herhangi bir tuşa basın.");
		child_process.spawnSync("pause", { shell: true, stdio: [0, 1, 2] });
		await print(absoluteFileName);
	} catch (e) {
		console.log("Program bir hata verdi ve kapatılacak.");
		console.log(e);
		child_process.spawnSync("pause", { shell: true, stdio: [0, 1, 2] });
	}
	child_process.spawnSync("pause", { shell: true, stdio: [0, 1, 2] });
})();
