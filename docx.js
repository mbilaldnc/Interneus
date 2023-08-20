const fs = require("fs");
const docx = require("docx");
// const data = require("./data.json");
// const prompt = require("prompt");
const {
	Document,
	Paragraph,
	Packer,
	TextRun,
	Table,
	TableRow,
	TableCell,
	TextDirection,
	Style,
	VerticalAlign,
	AlignmentType,
	HeightRule,
	WidthType,
	OverlapType,
	RelativeVerticalPosition,
	RelativeHorizontalPosition,
	TableAnchorType,
	TableLayoutType,
	SectionType,
} = docx;

//Even though, google docs is probably a better way to create and share docs,
//it doesn't support rotated text fields in tables, so I had to use docx instead.

// Documents contain sections, you can have multiple sections per document, go here to learn more about sections
// This simple example will only contain one section

// function new MyTableCell(props = { text: "", width: 550, bold: false, textDirection: TextDirection.LEFT_TO_RIGHT_TOP_TO_BOTTOM }) {
// 	return new TableCell({
// 		...props,
// 		children: [
// 			new Paragraph({
// 				children: [new TextRun({ text: props.text, bold: props.bold })],
// 				alignment: AlignmentType.CENTER,
// 			}),
// 		],
// 		width: {
// 			size: props.width,
// 			type: WidthType.DXA,
// 		},
// 		textDirection: props.textDirection === "bottomToTop" ? TextDirection.BOTTOM_TO_TOP_LEFT_TO_RIGHT : TextDirection.LEFT_TO_RIGHT_TOP_TO_BOTTOM,
// 	});
// }

class MyTableCell extends TableCell {
	constructor(options = {}) {
		const { text = "", width = 550, bold = false, textDirection = TextDirection.LEFT_TO_RIGHT_TOP_TO_BOTTOM } = options;
		const newProps = {
			...options,
			children: [
				new Paragraph({
					children: [new TextRun({ text: text, bold: bold })],
					alignment: AlignmentType.CENTER,
				}),
			],
			verticalAlign: VerticalAlign.CENTER,
			textDirection: textDirection === "bottomToTop" ? TextDirection.BOTTOM_TO_TOP_LEFT_TO_RIGHT : TextDirection.LEFT_TO_RIGHT_TOP_TO_BOTTOM,
			width: {
				size: width,
				type: WidthType.DXA,
			},
		};
		// console.log(text, width, newProps.width);
		super(newProps);
	}
}

class MyTableRow extends TableRow {
	constructor(options) {
		const { height = 300, children = [], ...leftovers } = options || {};
		super({
			...leftovers,
			height: {
				value: height || 300,
				rule: HeightRule.EXACT,
			},
			children: children,
		});
	}
}

class EmptyTableRow extends MyTableRow {
	constructor(columnCount) {
		super({
			children: Array.from(Array(columnCount), () => {
				return new MyTableCell();
			}),
		});
	}
}

module.exports = async (fileName) => {
	// const { fileName } = await prompt.get([{ name: "fileName", description: "Uzantısını yazmadan dosya adını giriniz" }]);
	const data = JSON.parse(await fs.promises.readFile(`./${fileName}.json`, "utf8"));
	console.log("Generating docx...");
	const sections = [];
	for (const [name, dates] of Object.entries(data)) {
		// sort from old to new
		const revizedDates = Object.keys(dates).sort((a, b) => {
			const splitedA = a.split(".");
			const aYear = Number(splitedA[2]);
			const aMonth = Number(splitedA[1]);
			const aDay = Number(splitedA[0]);
			const splitedB = b.split(".");
			const bYear = Number(splitedB[2]);
			const bMonth = Number(splitedB[1]);
			const bDay = Number(splitedB[0]);
			return aYear - bYear || aMonth - bMonth || aDay - bDay;
		});
		// console.log(revizedDates);
		const datesHasHgOrPTOrCRPOrSedim = revizedDates.filter((date) => {
			const keys = Object.keys(dates[date]);
			return (
				keys.some((key) => key.includes("HGB")) ||
				keys.some((key) => key.includes("Protrombin Zamanı")) ||
				keys.some((key) => key.includes("CRP")) ||
				keys.some((key) => key.includes("Sedim"))
			);
		});

		const hemogramTable = new Table({
			float: {
				horizontalAnchor: TableAnchorType.TEXT,
				verticalAnchor: TableAnchorType.TEXT,
				overlap: OverlapType.NEVER,
			},
			rows: [
				new TableRow({
					height: {
						value: 1300,
						rule: HeightRule.EXACT,
					},
					children: [
						new MyTableCell({ bold: true, text: "Tarih", textDirection: "bottomToTop" }),
						new MyTableCell({ text: "HGB" }),
						new MyTableCell({ text: "HCT" }),
						new MyTableCell({ text: "MCV" }),
						new MyTableCell({ text: "BK" }),
						new MyTableCell({ text: "NÖTROFİL", textDirection: "bottomToTop" }),
						new MyTableCell({ text: "TROMBOSİT", textDirection: "bottomToTop" }),
						new MyTableCell({ text: "CRP" }),
						new MyTableCell({ text: "SEDİM", textDirection: "bottomToTop" }),
						new MyTableCell({ text: "PT" }),
						new MyTableCell({ text: "PTT" }),
						new MyTableCell({ text: "INR" }),
					],
				}),
			]
				.concat(
					datesHasHgOrPTOrCRPOrSedim.map((date) => {
						const dateValues = dates[date];
						return new MyTableRow({
							children: [
								new MyTableCell({ text: date }),
								new MyTableCell({ text: dateValues[Object.keys(dateValues).find((key) => key.includes("HGB"))] || "" }),
								new MyTableCell({ text: dateValues[Object.keys(dateValues).find((key) => key.includes("HCT"))] || "" }),
								new MyTableCell({ text: dateValues[Object.keys(dateValues).find((key) => key.includes("MCV"))] || "" }),
								new MyTableCell({ text: dateValues[Object.keys(dateValues).find((key) => key.includes("WBC"))] || "" }),
								new MyTableCell({ text: dateValues[Object.keys(dateValues).find((key) => key.includes("NEU"))] || "" }),
								new MyTableCell({ text: dateValues[Object.keys(dateValues).find((key) => key.includes("PLT"))] || "" }),
								new MyTableCell({ text: dateValues[Object.keys(dateValues).find((key) => key.includes("CRP"))] || "" }),
								new MyTableCell({ text: dateValues[Object.keys(dateValues).find((key) => key.includes("Sedim"))] || "" }),
								new MyTableCell({ text: dateValues[Object.keys(dateValues).find((key) => key.includes("Protrombin Zamanı"))] || "" }),
								new MyTableCell({ text: dateValues[Object.keys(dateValues).find((key) => key.includes("APTT"))] || "" }),
								new MyTableCell({ text: dateValues[Object.keys(dateValues).find((key) => key.includes("INR"))] || "" }),
							],
						});
					})
				)
				.concat(Array.from(Array(7 - datesHasHgOrPTOrCRPOrSedim.length), () => new EmptyTableRow(12))),
		});

		let allDateValuesCombined = {};
		for (const date of revizedDates) {
			Object.assign(allDateValuesCombined, dates[date]);
		}

		const elisaTable = new Table({
			// width: {
			// 	value: 35,
			// 	rule: WidthType.PERCENTAGE,
			// },
			float: {
				horizontalAnchor: TableAnchorType.TEXT,
				verticalAnchor: TableAnchorType.TEXT,
				overlap: OverlapType.NEVER,
				// relativeHorizontalPosition: RelativeHorizontalPosition.LEFT,
				// relativeVerticalPosition: RelativeVerticalPosition.TOP,
				leftFromText: 200,
			},
			rows: [
				new MyTableRow({
					children: [new MyTableCell({ text: "Tarih", width: 1100 }), new MyTableCell({ width: 1000 })],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "AntiHBs", width: 1100 }),
						new MyTableCell({
							width: 1000,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Anti HBs"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "HBcIgG", width: 1100 }),
						new MyTableCell({
							width: 1000,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Anti HBc IGG"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "HBcIgM", width: 1100 }),
						new MyTableCell({
							width: 1000,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Anti HBc IgM"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "HBeAg", width: 1100 }),
						new MyTableCell({
							width: 1000,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("HBeAg"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "AntiHBe", width: 1100 }),
						new MyTableCell({
							width: 1000,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Anti HBe"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "AntiHCV", width: 1100 }),
						new MyTableCell({
							width: 1000,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Anti HCV"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "AntiHIV", width: 1100 }),
						new MyTableCell({
							width: 1000,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Anti HIV"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "HBV-DNA", width: 1100 }),
						new MyTableCell({
							width: 1000,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("HBV-DNA"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "HCV-RNA", width: 1100 }),
						new MyTableCell({
							width: 1000,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("HCV-RNA"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "HBV Viral Yük", width: 1100 }),
						new MyTableCell({
							width: 1000,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("HBV Viral Yük"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "HCV Viral Yük", width: 1100 }),
						new MyTableCell({
							width: 1000,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("HCV Viral Yük"))] || "",
						}),
					],
				}),
			],
		});

		const TmAndLipidTable = new Table({
			float: {
				horizontalAnchor: TableAnchorType.TEXT,
				verticalAnchor: TableAnchorType.TEXT,
				overlap: OverlapType.NEVER,
				// relativeHorizontalPosition: RelativeHorizontalPosition.LEFT,
				// relativeVerticalPosition: RelativeVerticalPosition.TOP,
				leftFromText: 100,
			},
			rows: [
				new MyTableRow({
					children: [new MyTableCell({ text: "Tarih" }), new MyTableCell({ width: 800 })],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "AFP" }),
						new MyTableCell({
							width: 800,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Alfa-feto"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "CEA" }),
						new MyTableCell({
							width: 800,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("CEA"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "Ca19-9" }),
						new MyTableCell({
							width: 800,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Ca 19-9"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "Ca15-3" }),
						new MyTableCell({
							width: 800,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Ca 15-3"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "Ca125" }),
						new MyTableCell({
							width: 800,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Ca125"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "HDL" }),
						new MyTableCell({
							width: 800,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("HDL"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "VLDL" }),
						new MyTableCell({
							width: 800,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("VLDL"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "LDL" }),
						new MyTableCell({
							width: 800,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("(LDL"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "Tot.Kolest." }),
						new MyTableCell({
							width: 800,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Kolesterol (Total)"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "Trigliserit" }),
						new MyTableCell({
							width: 800,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Trigliserid"))] || "",
						}),
					],
				}),
			],
		});

		const TİTTable = new Table({
			// TODO
			float: {
				horizontalAnchor: TableAnchorType.TEXT,
				verticalAnchor: TableAnchorType.TEXT,
				overlap: OverlapType.NEVER,
				// relativeHorizontalPosition: RelativeHorizontalPosition.LEFT,
				// relativeVerticalPosition: RelativeVerticalPosition.TOP,
				leftFromText: 100,
			},
			rows: [
				new MyTableRow({
					children: [new MyTableCell({ text: "Tarih" }), new MyTableCell({ width: 800 })],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "AFP" }),
						new MyTableCell({
							width: 800,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Alfa-feto"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "CEA" }),
						new MyTableCell({
							width: 800,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("CEA"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "Ca19-9" }),
						new MyTableCell({
							width: 800,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Ca 19-9"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "Ca15-3" }),
						new MyTableCell({
							width: 800,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Ca 15-3"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "Ca125" }),
						new MyTableCell({
							width: 800,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Ca125"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "HDL" }),
						new MyTableCell({
							width: 800,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("HDL"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "VLDL" }),
						new MyTableCell({
							width: 800,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("VLDL"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "LDL" }),
						new MyTableCell({
							width: 800,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("(LDL"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "Tot.Kolest." }),
						new MyTableCell({
							width: 800,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Kolesterol (Total)"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "Trigliserit" }),
						new MyTableCell({
							width: 800,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Trigliserid"))] || "",
						}),
					],
				}),
			],
		});

		sections.push({
			properties: {
				page: {
					margin: {
						//DXA values which are 1/20th of a point which is 1/72 of an inch
						// which means 1/1440th of an inch
						// 567 DXA = ~1 cm
						top: 567,
						bottom: 567,
						left: 567,
						right: 567,
					},
				},
				type: SectionType.NEXT_PAGE,
			},
			children: [
				new Paragraph({
					children: [new TextRun({ text: "T.C. KOCAELİ ÜNİVERSİTESİ ARAŞTIRMA VE UYGULAMA HASTANESİ", bold: true })],
					alignment: AlignmentType.CENTER,
					style: "11pt",
				}),
				new Paragraph({
					children: [new TextRun({ text: "GASTROENTEROLOJİ YATAN HASTA LABORATUVAR BULGULARI", bold: true })],
					alignment: AlignmentType.CENTER,
					style: "11pt",
				}),
				new Paragraph({
					children: [new TextRun({ text: "Hastanın Adı Soyadı:", bold: true }), new TextRun({ text: name })],
					alignment: AlignmentType.LEFT,
					style: "11pt",
				}),
				hemogramTable,
				elisaTable,
				TmAndLipidTable,
			],
		});
	}

	const doc = new Document({
		styles: {
			default: {
				document: {
					run: {
						size: "7pt",
						font: "Times New Roman",
					},
				},
			},
			paragraphStyles: [
				{
					id: "11pt",
					name: "11pt",
					basedOn: "Normal",
					next: "Normal",
					run: {
						size: "11pt",
					},
				},
			],
		},
		sections,
	});

	// Used to export the file into a .docx file
	const buffer = await Packer.toBuffer(doc);
	await fs.promises.writeFile(`${fileName}.docx`, buffer);
	console.log("Docx generated!");

	// Done! A file called 'My Document.docx' will be in your file system.
};
