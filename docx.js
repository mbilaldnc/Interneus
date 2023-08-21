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
	/**
	 *
	 * @param {docx.ITableCellOptions} options
	 */
	constructor(options = {}) {
		const {
			text = "",
			width,
			bold = false,
			textDirection = TextDirection.LEFT_TO_RIGHT_TOP_TO_BOTTOM,
			horizontalAlign = AlignmentType.CENTER,
		} = options;
		/**
		 * @type {docx.ITableCellOptions} newProps
		 */
		const newProps = {
			...options,
			children: [
				new Paragraph({
					children: [new TextRun({ text: text, bold: bold })],
					alignment: horizontalAlign,
				}),
			],
			verticalAlign: VerticalAlign.CENTER,
			textDirection: textDirection === "bottomToTop" ? TextDirection.BOTTOM_TO_TOP_LEFT_TO_RIGHT : TextDirection.LEFT_TO_RIGHT_TOP_TO_BOTTOM,
			width: {
				size: width || 550,
				type: WidthType.DXA,
			},
		};
		// console.log(text, width, newProps.width);
		super(newProps);
	}
}

class MyTableRow extends TableRow {
	/**
	 *
	 * @param {docx.ITableRowOptions} options
	 */
	constructor(options) {
		const { height, children = [], ...leftovers } = options || {};
		super({
			...leftovers,
			height: height || {
				value: 270,
				rule: HeightRule.EXACT,
			},
			children: children,
		});
	}
}

class EmptyTableRow extends MyTableRow {
	/**
	 *
	 * @param {Number} columnCount Empty table row with given column count
	 */
	constructor(columnCount) {
		super({
			children: Array.from(Array(columnCount), () => {
				return new MyTableCell();
			}),
		});
	}
}

module.exports = async (absoluteFileName) => {
	// const { fileName } = await prompt.get([{ name: "fileName", description: "Uzantısını yazmadan dosya adını giriniz" }]);
	const data = JSON.parse(await fs.promises.readFile(`${absoluteFileName}.json`, "utf8"));
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

		const datesHasTİT = revizedDates.filter((date) => {
			const keys = Object.keys(dates[date]);
			const TİT = dates[date][keys.find((key) => key.includes("TİT"))];
			return TİT && Object.keys(TİT).length;
		});
		const TİTTable = new Table({
			// float: {
			// 	horizontalAnchor: TableAnchorType.TEXT,
			// 	verticalAnchor: TableAnchorType.TEXT,
			// 	overlap: OverlapType.NEVER,
			// 	// relativeHorizontalPosition: RelativeHorizontalPosition.LEFT,
			// 	// relativeVerticalPosition: RelativeVerticalPosition.TOP,
			// 	topFromText: 2500,
			// },
			width: {
				size: 100,
				type: WidthType.PERCENTAGE,
			},
			rows: [
				new MyTableRow({ children: [new MyTableCell({ text: "İDRAR", columnSpan: 12 })] }),
				new MyTableRow({
					height: {
						value: 1000,
						rule: HeightRule.EXACT,
					},
					children: [
						new MyTableCell({ text: "Tarih", bold: true }),
						new MyTableCell({ text: "Renk" }),
						new MyTableCell({ text: "pH" }),
						new MyTableCell({ text: "Dansite" }),
						new MyTableCell({ text: "Protein" }),
						new MyTableCell({ text: "Kan" }),
						new MyTableCell({ text: "Lökosit" }),
						new MyTableCell({ text: "Şeker" }),
						new MyTableCell({ text: "Keton" }),
						new MyTableCell({ text: "Bilirubin" }),
						new MyTableCell({ text: "Ürobilinojen", textDirection: "bottomToTop" }),
						new MyTableCell({ text: "MİKROSKOBİK MUAYENELER", width: 1600 }),
					],
				}),
			]
				.concat(
					datesHasTİT.slice(0, 3).map((date) => {
						const dateValues = dates[date];
						const TİT = dateValues[Object.keys(dateValues).find((key) => key.includes("TİT"))] || {};
						return new MyTableRow({
							children: [
								new MyTableCell({ text: date }),
								new MyTableCell({ text: TİT[Object.keys(TİT).find((key) => key.includes("Renk"))] }),
								new MyTableCell({ text: TİT[Object.keys(TİT).find((key) => key.includes("pH"))] }),
								new MyTableCell({ text: TİT[Object.keys(TİT).find((key) => key.includes("Dansite"))] }),
								new MyTableCell({ text: TİT[Object.keys(TİT).find((key) => key.includes("Protein"))] }),
								new MyTableCell({ text: TİT[Object.keys(TİT).find((key) => key.includes("Kan"))] }),
								new MyTableCell({ text: TİT[Object.keys(TİT).find((key) => key.includes("Lökosit"))] }),
								new MyTableCell({ text: TİT[Object.keys(TİT).find((key) => key.includes("Glukoz"))] }),
								new MyTableCell({ text: TİT[Object.keys(TİT).find((key) => key.includes("Keton"))] }),
								new MyTableCell({ text: TİT[Object.keys(TİT).find((key) => key.includes("Bilirubin"))] }),
								new MyTableCell({ text: TİT[Object.keys(TİT).find((key) => key.includes("Urobilinojen"))] }),
								new MyTableCell({ text: "" }),
							],
						});
					})
				)
				.concat(Array.from(Array(3 - datesHasTİT.length > 0 ? 3 - datesHasTİT.length : 0), () => new EmptyTableRow(12))),
		});

		const datesHasBK = revizedDates.filter((date) => {
			const BKKeys = [
				"Açlık Kan Şekeri (AKŞ)",
				"Ürea",
				"BUN",
				"Kreatinin",
				"Bilirubin",
				"AST", // "Hastabaşı" overlaps with AST
				"ALT",
				"GGT",
				"ALP",
				"LDH",
				"Sodyum",
				"Kalsiyum",
				"Magnezyum",
				"Fosfor",
				"Ürik asit",
				"Amilaz",
				"Lipaz",
				"Protein, Total",
				"Albumin",
			];
			const keys = Object.keys(dates[date]);
			return keys.some((key) => BKKeys.some((BKKey) => key.includes(BKKey)));
		});
		const BKTable = new Table({
			width: {
				size: 100,
				type: WidthType.PERCENTAGE,
			},
			rows: [
				new MyTableRow({ children: [new MyTableCell({ text: "BİYOKİMYA", columnSpan: 21 })] }),
				new MyTableRow({
					height: {
						value: 1000,
						rule: HeightRule.EXACT,
					},
					children: [
						new MyTableCell({ text: "Tarih", bold: true, textDirection: "bottomToTop" }),
						new MyTableCell({ text: "Glukoz", textDirection: "bottomToTop" }),
						new MyTableCell({ text: "Üre" }),
						new MyTableCell({ text: "Kreatinin", textDirection: "bottomToTop" }),
						new MyTableCell({ text: "Ca" }),
						new MyTableCell({ text: "Na" }),
						new MyTableCell({ text: "K" }),
						new MyTableCell({ text: "Mg" }),
						new MyTableCell({ text: "P" }),
						new MyTableCell({ text: "Ürik Asit", textDirection: "bottomToTop" }),
						new MyTableCell({ text: "Tot.Bil.", textDirection: "bottomToTop" }),
						new MyTableCell({ text: "Dir.Bil", textDirection: "bottomToTop" }),
						new MyTableCell({ text: "AST" }),
						new MyTableCell({ text: "ALT" }),
						new MyTableCell({ text: "ALP" }),
						new MyTableCell({ text: "GGT" }),
						new MyTableCell({ text: "LDH" }),
						new MyTableCell({ text: "Amilaz", textDirection: "bottomToTop" }),
						new MyTableCell({ text: "Lipaz", textDirection: "bottomToTop" }),
						new MyTableCell({ text: "Tot.Prot.", textDirection: "bottomToTop" }),
						new MyTableCell({ text: "Albumin", textDirection: "bottomToTop" }),
					],
				}),
			]
				.concat(
					datesHasBK.map((date) => {
						const dateValues = dates[date];
						const DüzeltilmişKalsiyum = dateValues[Object.keys(dateValues).find((key) => key.includes("Düzeltilmiş Kalsiyum"))];
						const Kalsiyum = dateValues[Object.keys(dateValues).find((key) => key === "Kalsiyum")];
						const DüzeltilmişSodyum = dateValues[Object.keys(dateValues).find((key) => key.includes("Düzeltilmiş Sodyum"))];
						const Sodyum = dateValues[Object.keys(dateValues).find((key) => key === "Sodyum (Na)")];
						return new MyTableRow({
							children: [
								new MyTableCell({ text: date }),
								new MyTableCell({ text: dateValues[Object.keys(dateValues).find((key) => key.includes("Açlık Kan Şekeri (AKŞ)"))] }),
								new MyTableCell({ text: dateValues[Object.keys(dateValues).find((key) => key.includes("Ürea"))] }),
								new MyTableCell({ text: dateValues[Object.keys(dateValues).find((key) => key.includes("Kreatinin"))] }),
								new MyTableCell({ text: DüzeltilmişKalsiyum || Kalsiyum || "" }),
								new MyTableCell({ text: DüzeltilmişSodyum || Sodyum || "" }),
								new MyTableCell({ text: dateValues[Object.keys(dateValues).find((key) => key.includes("Potasyum"))] }),
								new MyTableCell({ text: dateValues[Object.keys(dateValues).find((key) => key.includes("Magnezyum"))] }),
								new MyTableCell({ text: dateValues[Object.keys(dateValues).find((key) => key.includes("Fosfor"))] }),
								new MyTableCell({ text: dateValues[Object.keys(dateValues).find((key) => key.includes("Ürik asit"))] }),
								new MyTableCell({ text: dateValues[Object.keys(dateValues).find((key) => key.includes("Bilirubin, Total"))] }),
								new MyTableCell({ text: dateValues[Object.keys(dateValues).find((key) => key.includes("Bilirubin, Direkt"))] }),
								new MyTableCell({ text: dateValues[Object.keys(dateValues).find((key) => key.includes("AST (SGOT)"))] }),
								new MyTableCell({ text: dateValues[Object.keys(dateValues).find((key) => key.includes("ALT (SGPT)"))] }),
								new MyTableCell({ text: dateValues[Object.keys(dateValues).find((key) => key.includes("ALP(Alkalen Fosfataz)"))] }),
								new MyTableCell({ text: dateValues[Object.keys(dateValues).find((key) => key.includes("GGT"))] }),
								new MyTableCell({ text: dateValues[Object.keys(dateValues).find((key) => key.includes("LDH"))] }),
								new MyTableCell({ text: dateValues[Object.keys(dateValues).find((key) => key.includes("Amilaz"))] }),
								new MyTableCell({ text: dateValues[Object.keys(dateValues).find((key) => key.includes("Lipaz"))] }),
								new MyTableCell({ text: dateValues[Object.keys(dateValues).find((key) => key.includes("Protein, Total"))] }),
								new MyTableCell({ text: dateValues[Object.keys(dateValues).find((key) => key.includes("Albumin"))] }),
							],
						});
					})
				)
				.concat(Array.from(Array(7 - datesHasBK.length > 0 ? 7 - datesHasBK.length : 0), () => new EmptyTableRow(21))),
		});

		const datesHasHemogram = revizedDates.filter((date) => {
			const keys = Object.keys(dates[date]);
			return (
				keys.some((key) => key.includes("Hemogram")) ||
				keys.some((key) => key.includes("Protrombin Zamanı")) ||
				keys.some((key) => key.includes("APTT")) ||
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
						value: 1100,
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
					datesHasHemogram.map((date) => {
						const dateValues = dates[date];
						const hemogram = dateValues[Object.keys(dateValues).find((key) => key.includes("Hemogram"))] || {};
						const koagulometre = dateValues[Object.keys(dateValues).find((key) => key.includes("Protrombin Zamanı"))] || {};
						return new MyTableRow({
							children: [
								new MyTableCell({ text: date }),
								new MyTableCell({ text: hemogram[Object.keys(hemogram).find((key) => key.includes("HGB"))] || "" }),
								new MyTableCell({ text: hemogram[Object.keys(hemogram).find((key) => key.includes("HCT"))] || "" }),
								new MyTableCell({ text: hemogram[Object.keys(hemogram).find((key) => key.includes("MCV"))] || "" }),
								new MyTableCell({ text: hemogram[Object.keys(hemogram).find((key) => key.includes("WBC"))] || "" }),
								new MyTableCell({ text: hemogram[Object.keys(hemogram).find((key) => key.includes("NEU"))] || "" }),
								new MyTableCell({ text: hemogram[Object.keys(hemogram).find((key) => key.includes("PLT"))] || "" }),
								new MyTableCell({ text: dateValues[Object.keys(dateValues).find((key) => key.includes("CRP"))] || "" }),
								new MyTableCell({ text: dateValues[Object.keys(dateValues).find((key) => key.includes("Sedim"))] || "" }),
								new MyTableCell({
									text: koagulometre[Object.keys(koagulometre).find((key) => key.includes("Protrombin Zamanı"))] || "",
								}),
								new MyTableCell({ text: dateValues[Object.keys(dateValues).find((key) => key.includes("APTT"))] || "" }),
								new MyTableCell({ text: koagulometre[Object.keys(koagulometre).find((key) => key.includes("INR"))] || "" }),
							],
						});
					})
				)
				.concat(Array.from(Array(7 - datesHasHemogram.length > 0 ? 7 - datesHasHemogram.length : 0), () => new EmptyTableRow(12))),
		});

		let allDateValuesCombined = {};
		for (const date of revizedDates) {
			Object.assign(allDateValuesCombined, dates[date]);
		}

		const eliseRowHeight = 240;
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
					children: [new MyTableCell({ text: "Tarih", width: 1100, bold: true }), new MyTableCell({ width: 1150 })],
					height: {
						value: eliseRowHeight,
						rule: HeightRule.EXACT,
					},
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "HbsAg", width: 1100, horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							width: 1150,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("HBs Ag"))] || "",
						}),
					],
					height: {
						value: eliseRowHeight,
						rule: HeightRule.EXACT,
					},
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "AntiHBs", width: 1100, horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							width: 1150,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Anti HBs"))] || "",
						}),
					],
					height: {
						value: eliseRowHeight,
						rule: HeightRule.EXACT,
					},
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "HBcIgG", width: 1100, horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							width: 1150,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Anti HBc IGG"))] || "",
						}),
					],
					height: {
						value: eliseRowHeight,
						rule: HeightRule.EXACT,
					},
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "HBcIgM", width: 1100, horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							width: 1150,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Anti HBc IgM"))] || "",
						}),
					],
					height: {
						value: eliseRowHeight,
						rule: HeightRule.EXACT,
					},
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "HBeAg", width: 1100, horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							width: 1150,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("HBe Ag"))] || "",
						}),
					],
					height: {
						value: eliseRowHeight,
						rule: HeightRule.EXACT,
					},
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "AntiHBe", width: 1100, horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							width: 1150,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Anti HBe"))] || "",
						}),
					],
					height: {
						value: eliseRowHeight,
						rule: HeightRule.EXACT,
					},
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "AntiHCV", width: 1100, horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							width: 1150,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Anti HCV"))] || "",
						}),
					],
					height: {
						value: eliseRowHeight,
						rule: HeightRule.EXACT,
					},
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "AntiHIV", width: 1100, horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							width: 1150,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Anti HIV"))] || "",
						}),
					],
					height: {
						value: eliseRowHeight,
						rule: HeightRule.EXACT,
					},
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "HBV-DNA", width: 1100, horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							width: 1150,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("HBV-DNA"))] || "",
						}),
					],
					height: {
						value: eliseRowHeight,
						rule: HeightRule.EXACT,
					},
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "HCV-RNA", width: 1100, horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							width: 1150,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("HCV-RNA"))] || "",
						}),
					],
					height: {
						value: eliseRowHeight,
						rule: HeightRule.EXACT,
					},
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "HBV Viral Yük", width: 1100, horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							width: 1150,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("HBV Viral Yük"))] || "",
						}),
					],
					height: {
						value: eliseRowHeight,
						rule: HeightRule.EXACT,
					},
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "HCV Viral Yük", width: 1100, horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							width: 1150,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("HCV Viral Yük"))] || "",
						}),
					],
					height: {
						value: eliseRowHeight,
						rule: HeightRule.EXACT,
					},
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "AntiHAV IgG", width: 1100, horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							width: 1150,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Anti HAV IgG"))] || "",
						}),
					],
					height: {
						value: eliseRowHeight,
						rule: HeightRule.EXACT,
					},
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "AntiHAV IgM", width: 1100, horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							width: 1150,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Anti HAV IgM"))] || "",
						}),
					],
					height: {
						value: eliseRowHeight,
						rule: HeightRule.EXACT,
					},
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "Rubella lgG", width: 1100, horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							width: 1150,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Rubella lgG"))] || "",
						}),
					],
					height: {
						value: eliseRowHeight,
						rule: HeightRule.EXACT,
					},
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "Rubella lgM", width: 1100, horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							width: 1150,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Rubella lgM"))] || "",
						}),
					],
					height: {
						value: eliseRowHeight,
						rule: HeightRule.EXACT,
					},
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "CMV IgG", width: 1100, horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							width: 1150,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("CMV IgG"))] || "",
						}),
					],
					height: {
						value: eliseRowHeight,
						rule: HeightRule.EXACT,
					},
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "CMV IgM", width: 1100, horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							width: 1150,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("CMV IgM"))] || "",
						}),
					],
					height: {
						value: eliseRowHeight,
						rule: HeightRule.EXACT,
					},
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
					children: [new MyTableCell({ text: "Tarih", bold: true }), new MyTableCell({ width: 800 })],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "AFP", horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							width: 800,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Alfa-feto"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "CEA", horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							width: 800,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("CEA"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "Ca19-9", horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							width: 800,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Ca 19-9"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "Ca15-3", horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							width: 800,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Ca 15-3"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "Ca125", horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							width: 800,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Ca 125"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "HDL", horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							width: 800,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("HDL"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "VLDL", horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							width: 800,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("VLDL"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "LDL", horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							width: 800,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("(LDL"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "Tot.Kolest.", horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							width: 800,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Kolesterol (Total)"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "Trigliserit", horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							width: 800,
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Trigliserid"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [new MyTableCell(), new MyTableCell()],
				}),
				new MyTableRow({
					children: [new MyTableCell(), new MyTableCell()],
				}),
			],
		});

		const HormonTable = new Table({
			float: {
				horizontalAnchor: TableAnchorType.TEXT,
				verticalAnchor: TableAnchorType.TEXT,
				overlap: OverlapType.NEVER,
				// relativeHorizontalPosition: RelativeHorizontalPosition.LEFT,
				// relativeVerticalPosition: RelativeVerticalPosition.TOP,
				leftFromText: 0,
			},
			rows: [
				new MyTableRow({
					children: [
						new MyTableCell({ text: "Tarih", bold: true, width: 1000 }),
						new MyTableCell({ text: revizedDates.find((date) => dates[date].TSH), width: 1000 }),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "TSH", width: 1000, horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({ text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("TSH"))] || "" }),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "sT3", width: 1000, horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Serbest T3"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "sT4", width: 1000, horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Serbest T4"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "PTH", width: 1000, horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Parathormon"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "Prokalsitonin", width: 1000, horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Prokalsitonin"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "Demir", width: 1000, horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Demir (Fe)"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "TDBK", width: 1000, horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							text:
								allDateValuesCombined[
									Object.keys(allDateValuesCombined).find((key) => key.includes("Demir Bağlama Kapasitesi (Total)"))
								] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "Transferrin", width: 1000, horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Transferrin"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "Ferritin", width: 1000, horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Ferritin"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "Folik Asit", width: 1000, horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Folik Asit"))] || "",
						}),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "B-12 Vit", width: 1000, horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("Vitamin B12"))] || "",
						}),
					],
				}),
			],
		});

		const AssitTable = new Table({
			float: {
				horizontalAnchor: TableAnchorType.MARGIN,
				verticalAnchor: TableAnchorType.TEXT,
				overlap: OverlapType.NEVER,
				// relativeHorizontalPosition: RelativeHorizontalPosition.LEFT,
				// relativeVerticalPosition: RelativeVerticalPosition.TOP,
				leftFromText: 200,
				topFromText: 0,
			},
			rows: [
				new MyTableRow({
					children: [new MyTableCell({ text: "Assit Analizi", bold: true, width: 1000 }), new MyTableCell({ width: 600 })],
				}),
				new MyTableRow({
					children: [new MyTableCell({ text: "Tot.Protein", width: 1000, horizontalAlign: AlignmentType.LEFT }), new MyTableCell()],
				}),
				new MyTableRow({
					children: [new MyTableCell({ text: "Albumin", width: 1000, horizontalAlign: AlignmentType.LEFT }), new MyTableCell()],
				}),
				new MyTableRow({
					children: [new MyTableCell({ text: "SAAG", width: 1000, horizontalAlign: AlignmentType.LEFT }), new MyTableCell()],
				}),
				new MyTableRow({
					children: [new MyTableCell({ text: "Lökosit", width: 1000, horizontalAlign: AlignmentType.LEFT }), new MyTableCell()],
				}),
				new MyTableRow({
					children: [new MyTableCell({ text: "Eritrosit", width: 1000, horizontalAlign: AlignmentType.LEFT }), new MyTableCell()],
				}),
			],
		});

		const SeruloplazminTable = new Table({
			float: {
				horizontalAnchor: TableAnchorType.TEXT,
				verticalAnchor: TableAnchorType.TEXT,
				overlap: OverlapType.NEVER,
				// relativeHorizontalPosition: RelativeHorizontalPosition.LEFT,
				// relativeVerticalPosition: RelativeVerticalPosition.TOP,
				absoluteVerticalPosition: "3cm",
				leftFromText: 200,
			},
			rows: [
				new MyTableRow({
					children: [new MyTableCell({ text: "Tarih", bold: true, width: 900 }), new MyTableCell({ width: 500 })],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "Seruloplazmin", width: 900, horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({ width: 500 }),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "Amonyak", width: 900, horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({ width: 500 }),
					],
				}),
			],
		});

		const OtoimmünMarkerTable = new Table({
			float: {
				horizontalAnchor: TableAnchorType.TEXT,
				verticalAnchor: TableAnchorType.TEXT,
				overlap: OverlapType.NEVER,
				// relativeHorizontalPosition: RelativeHorizontalPosition.LEFT,
				// relativeVerticalPosition: RelativeVerticalPosition.TOP,
				absoluteVerticalPosition: "0cm",
				absoluteHorizontalPosition: "7.5cm",
				// leftFromText: "1.5cm",
				topFromText: "0.5cm",
			},
			rows: [
				new MyTableRow({
					height: {
						value: 1000,
						rule: HeightRule.EXACT,
					},
					children: [
						new MyTableCell({ text: "Otoimmün Marker", bold: true, width: 800 }),
						new MyTableCell({ text: "ANA" }),
						new MyTableCell({ text: "PR3 ANCA", textDirection: "bottomToTop" }),
						new MyTableCell({ text: "MPO ANCA", textDirection: "bottomToTop" }),
						new MyTableCell({ text: "AMA" }),
						new MyTableCell({ text: "ASMA" }),
						new MyTableCell({ text: "AntiLKM", textDirection: "bottomToTop" }),
						new MyTableCell({ text: "IgA" }),
						new MyTableCell({ text: "IgM" }),
						new MyTableCell({ text: "IgG" }),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell(),
						new MyTableCell({
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key === "ANA")] || "",
						}),
						new MyTableCell({
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("PR3 ANCA"))] || "",
						}),
						new MyTableCell({
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("MPO ANCA"))] || "",
						}),
						new MyTableCell({
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key === "AMA")] || "",
						}),
						new MyTableCell({
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key === "ASMA")] || "",
						}),
						new MyTableCell({
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("LKM"))] || "",
						}),
						new MyTableCell({
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("IgA (İmmün kompleks)"))] || "",
						}),
						new MyTableCell({
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("IgM (İmmün kompleks)"))] || "",
						}),
						new MyTableCell({
							text: allDateValuesCombined[Object.keys(allDateValuesCombined).find((key) => key.includes("IgG (İmmün kompleks)"))] || "",
						}),
					],
				}),
			],
		});

		const ChildPughSkorlama = new Table({
			float: {
				horizontalAnchor: TableAnchorType.TEXT,
				verticalAnchor: TableAnchorType.TEXT,
				overlap: OverlapType.NEVER,
				// relativeHorizontalPosition: RelativeHorizontalPosition.LEFT,
				// relativeVerticalPosition: RelativeVerticalPosition.TOP,
				leftFromText: 0,
				// topFromText: "0.5cm",
				absoluteVerticalPosition: "6cm",
			},
			rows: [
				new MyTableRow({
					children: [new MyTableCell({ text: "Child-Pugh Skorlama", bold: true, columnSpan: 5 })],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "Measure", bold: true, width: 1800 }),
						new MyTableCell({ text: "1 point", bold: true }),
						new MyTableCell({ text: "2 point", bold: true, width: 800 }),
						new MyTableCell({ text: "3 point", bold: true, width: 800 }),
						new MyTableCell({ text: "Skor", bold: true }),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "Total bilirubin, µmol/l (mg/dl)", horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({ text: "<2", horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({ text: "2-3", horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({ text: ">3", horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({ text: "" }),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "Serum albumin, g/l", horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({ text: ">3,5", horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({ text: "2,8-3,5", horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({ text: "<2,8", horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({ text: "" }),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "INR", horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({ text: "<1,7", horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({ text: "1,7-2,3", horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({ text: ">2,3", horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({ text: "" }),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "Ascites", horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({ text: "Yok", horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({ text: "Hafif", horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({ text: "Yaygın", horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({ text: "" }),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "Hepatic encephalopathy", horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({ text: "Yok", horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({ text: "Grade I-II", horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({ text: "Grade III-IV", horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({ text: "" }),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell({ text: "Toplam", bold: true, columnSpan: 4, horizontalAlign: AlignmentType.LEFT }),
						new MyTableCell({ text: "" }),
					],
				}),
			],
		});

		const RansonKriterleriTable = new Table({
			float: {
				horizontalAnchor: TableAnchorType.TEXT,
				verticalAnchor: TableAnchorType.TEXT,
				overlap: OverlapType.NEVER,
				// relativeHorizontalPosition: RelativeHorizontalPosition.LEFT,
				// relativeVerticalPosition: RelativeVerticalPosition.TOP,
				leftFromText: "0.5cm",
				topFromText: "0.5cm",
				absoluteVerticalPosition: "6cm",
			},
			width: {
				size: "10cm",
				type: WidthType.AUTO,
			},
			columnWidths: [200, 200, 200, 200, 200, 200, 200, 200, 200, 200, 200, 200],
			layout: TableLayoutType.FIXED,
			rows: [
				new MyTableRow({
					children: [new MyTableCell({ text: "Ranson Kriterleri", bold: true, columnSpan: 12 })],
				}),
				new MyTableRow({
					height: {
						value: 1200,
						rule: HeightRule.EXACT,
					},
					children: [
						new MyTableCell({ rowSpan: 2 }),
						new MyTableCell({ text: "Yaş>55", textDirection: "bottomToTop" }),
						new MyTableCell({ text: "WBC >16.000", textDirection: "bottomToTop" }),
						new MyTableCell({ text: "Glukoz >200", textDirection: "bottomToTop" }),
						new MyTableCell({ text: "LDH >350", textDirection: "bottomToTop" }),
						new MyTableCell({ text: "AST >250", textDirection: "bottomToTop" }),
						new MyTableCell({ text: "Hct düşüşü >%10", textDirection: "bottomToTop" }),
						new MyTableCell({ text: "BUN artışı >5 mg", textDirection: "bottomToTop" }),
						new MyTableCell({ text: "Baz Açığı >4 mEq", textDirection: "bottomToTop" }),
						new MyTableCell({ text: "PO2 <60", textDirection: "bottomToTop" }),
						new MyTableCell({ text: "Kalsiyum <8 mEq", textDirection: "bottomToTop" }),
						new MyTableCell({ text: "Sıvı Sekest. >6000ml", textDirection: "bottomToTop" }),
					],
				}),
				new MyTableRow({
					children: [
						new MyTableCell(),
						new MyTableCell(),
						new MyTableCell(),
						new MyTableCell(),
						new MyTableCell(),
						new MyTableCell(),
						new MyTableCell(),
						new MyTableCell(),
						new MyTableCell(),
						new MyTableCell(),
						new MyTableCell(),
					],
				}),
				new MyTableRow({
					children: [new MyTableCell({ text: "SKOR", bold: true }), new MyTableCell({ columnSpan: 5 }), new MyTableCell({ columnSpan: 6 })],
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
					style: "header",
				}),
				new Paragraph({
					children: [new TextRun({ text: "GASTROENTEROLOJİ YATAN HASTA LABORATUVAR BULGULARI", bold: true })],
					alignment: AlignmentType.CENTER,
					style: "header",
				}),
				new Paragraph({
					children: [],
				}),
				new Paragraph({
					children: [new TextRun({ text: "Hastanın Adı Soyadı: ", bold: true }), new TextRun({ text: name })],
					alignment: AlignmentType.LEFT,
					style: "11pt",
				}),
				new Paragraph({
					children: [],
				}),
				TİTTable,
				new Paragraph({
					children: [],
				}),
				BKTable,
				new Paragraph({
					children: [],
				}),
				hemogramTable,
				elisaTable,
				TmAndLipidTable,
				new Paragraph({
					children: [],
				}),
				HormonTable,
				AssitTable,
				SeruloplazminTable,
				OtoimmünMarkerTable,
				ChildPughSkorlama,
				RansonKriterleriTable,
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
					paragraph: {
						spacing: {
							line: 260,
						},
					},
				},
				{
					id: "header",
					name: "header",
					basedOn: "Normal",
					next: "Normal",
					run: {
						size: "11pt",
					},
					paragraph: {
						spacing: {
							line: 300,
						},
					},
				},
			],
		},
		sections,
	});

	// Used to export the file into a .docx file
	const buffer = await Packer.toBuffer(doc);
	await fs.promises.writeFile(`${absoluteFileName}.docx`, buffer);
	console.log(`${absoluteFileName}.docx successfully generated!`);
};
