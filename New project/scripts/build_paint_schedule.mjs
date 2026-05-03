import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "/Users/aurora/Documents/New project/outputs/paint_schedule";
const outputPath = `${outputDir}/multi_restaurant_paint_schedule.xlsx`;

const palette = {
  saRose: "#F48B7F",
  saRoseSoft: "#FAEDE1",
  saGilt: "#AC936D",
  saBlue: "#5AA6E5",
  feliceRed: "#C43632",
  feliceGreen: "#124718",
  feliceGold: "#BC985E",
  feliceStone: "#E8E1D1",
  feliceFig: "#4F2F49",
  ink: "#1F2933",
  muted: "#5C6470",
  line: "#D9DEE7",
  paper: "#FFFDF8",
  white: "#FFFFFF",
  missing: "#F3F4F6",
  missingText: "#6B7280",
};

const sourceFiles = {
  saEastHampton: "SAEH PAINT SCHEDULE.pdf; SA-East Hampton - Copy of PAINT.pdf",
  saBrookfield: "SA BROOKFIELD PAINT SCHEDULE MARCH 2025.pdf",
  saGelateria: "SA - Gelateria - PAINT SCHEDULE SEPTEMBER 2024.pdf",
  saSouthampton: "SA Southampton - PAINT.pdf",
  feliceHudson: "FELICE HUDSON PAINT SCHEDULE - SEPTEMBER 2024.pdf",
  felice64: "Felice - 64TH - PAINT Schedule.pdf",
  felice71: "Felice - 71st - PAINT Schedule.pdf",
  felice83: "Felice - 83RD - 4. Paint (1).pdf",
  feliceRoslyn: "Felice-Roslyn - PAINT SCHEDULE.pdf",
};

const santLocations = [
  ["Brookfield", "Available", "March 2025", sourceFiles.saBrookfield, ""],
  ["East Hampton", "Available", "Modified 05.25.23", sourceFiles.saEastHampton, "Duplicate East Hampton files supplied; normalized from the keyed schedule."],
  ["Gelateria", "Available", "September 2024", sourceFiles.saGelateria, ""],
  ["Hanley", "Missing", "", "", "Listed in reference image; no schedule PDF supplied."],
  ["Lafayette-SOHO", "Missing", "", "", "Listed in reference image; no schedule PDF supplied."],
  ["Madison Ave", "Missing", "", "", "Listed in reference image; no schedule PDF supplied."],
  ["Palm Beach", "Missing", "", "", "Listed in reference image; no schedule PDF supplied."],
  ["Regency", "Missing", "", "", "Listed in reference image; no schedule PDF supplied."],
  ["Sotheby's", "Missing", "", "", "Listed in reference image; no schedule PDF supplied."],
  ["Southampton", "Available", "", sourceFiles.saSouthampton, ""],
  ["West 4th", "Missing", "", "", "Listed in reference image; no schedule PDF supplied."],
];

const feliceLocations = [
  ["Hudson", "Available", "9/16/24", sourceFiles.feliceHudson, ""],
  ["64th", "Available", "", sourceFiles.felice64, ""],
  ["71st", "Available", "", sourceFiles.felice71, ""],
  ["83rd", "Available", "", sourceFiles.felice83, ""],
  ["Roslyn", "Available", "", sourceFiles.feliceRoslyn, ""],
];

const detailHeaders = [
  "Brand",
  "Restaurant",
  "Schedule Status",
  "Key",
  "Room / Zone",
  "Surface / Item",
  "Material",
  "Vendor / Paint Brand",
  "Product / Paint Type",
  "Color",
  "Finish / Sheen",
  "Notes / Directions",
  "Source File",
  "Source Date",
];

const details = [];
const add = (brand, restaurant, status, key, zone, item, material, vendor, product, color, finish, notes, source, date) => {
  details.push([brand, restaurant, status, key, zone, item, material, vendor, product, color, finish, notes, source, date]);
};

add("Sant Ambroeus", "Brookfield", "Available", "P 01", "Hallway", "Base boards", "", "Benjamin Moore", "AURA Advanced", "Cromwell Gray HC-103", "Satin", "Released 7/18/19; brush painted.", sourceFiles.saBrookfield, "March 2025");
add("Sant Ambroeus", "Brookfield", "Available", "P 02", "Hallway bathrooms (ground floor)", "Upper walls and ceiling", "", "Benjamin Moore", "AURA Bath & Spa", "Bed and Breakfast CC-184", "Flat", "Released 7/18/19.", sourceFiles.saBrookfield, "March 2025");
add("Sant Ambroeus", "Brookfield", "Available", "P 03", "Ceilings", "Metal ceiling soffit / soffit coves", "", "Benjamin Moore", "AURA Advanced", "Bermuda Sands 2100-60", "Flat", "Released 7/18/19.", sourceFiles.saBrookfield, "March 2025");
add("Sant Ambroeus", "Brookfield", "Available", "P 04", "Mezzanine bathrooms", "Upper walls and ceiling", "", "Benjamin Moore", "AURA Bath & Spa", "Webster Green HC-130", "Flat", "Released 7/18/19.", sourceFiles.saBrookfield, "March 2025");
add("Sant Ambroeus", "Brookfield", "Available", "P 05", "Back bar", "MDF box", "", "Benjamin Moore", "AURA Advanced", "Bermuda Sands 2100-60", "Matte", "Released 7/25/19.", sourceFiles.saBrookfield, "March 2025");
add("Sant Ambroeus", "Brookfield", "Available", "W 01A", "Walls", "Cafe, dining room and mezzanine plaster wall", "", "Domingue Finishes", "Venetian Marmorino plaster", "Magnolia", "", "Must be protected with additional coating for cleaning purposes; coating provided by Chateau Domingue.", sourceFiles.saBrookfield, "March 2025");
add("Sant Ambroeus", "Brookfield", "Available", "W 01B", "Walls", "Hallway plaster wall", "", "Domingue Finishes", "Venetian Marmorino plaster", "Boxwood", "", "Must be protected with additional coating for cleaning purposes; coating provided by Chateau Domingue.", sourceFiles.saBrookfield, "March 2025");
add("Sant Ambroeus", "Brookfield", "Available", "W 02", "Walls", "Hallway lacquered panels", "", "Benjamin Moore", "Gloss", "Avon Green HC-126", "Gloss", "To match hallway vitrines.", sourceFiles.saBrookfield, "March 2025");

add("Sant Ambroeus", "East Hampton", "Available", "P 01", "Dining room", "MDF wainscotting", "", "Chateau Domingue", "Limewash paint", "Dauphine", "", "", sourceFiles.saEastHampton, "Modified 05.25.23");
add("Sant Ambroeus", "East Hampton", "Available", "P 02", "Dining room", "Upper walls and ceiling", "", "Benjamin Moore by GC", "", "Glen Davis White", "Flat", "", sourceFiles.saEastHampton, "Modified 05.25.23");
add("Sant Ambroeus", "East Hampton", "Available", "P 03", "Dining room", "Floors", "", "Fine Paints of Europe per GC", "E25-12 - The One Thousand Collection; Hollandlac Brilliant", "", "Brilliant", "", sourceFiles.saEastHampton, "Modified 05.25.23");
add("Sant Ambroeus", "East Hampton", "Available", "P 04", "Restrooms", "Wainscotting", "", "Fine Paints of Europe per GC", "E8-34 Hollandlac Brilliant", "", "Brilliant", "", sourceFiles.saEastHampton, "Modified 05.25.23");
add("Sant Ambroeus", "East Hampton", "Available", "P 05", "Restrooms", "Upper walls and ceiling", "", "Chateau Domingue", "Limewash paint", "Espalier", "", "", sourceFiles.saEastHampton, "Modified 05.25.23");
add("Sant Ambroeus", "East Hampton", "Available", "P 06", "Back of house (BOH)", "Kitchens throughout", "", "Benjamin Moore by GC", "", "Decorator's White OC-149", "Semi-gloss", "", sourceFiles.saEastHampton, "Modified 05.25.23");
add("Sant Ambroeus", "East Hampton", "Available", "P 07", "Bus station area", "Kitchen wall and stair wall", "", "Benjamin Moore by GC", "", "Color match Dauphine limewash", "Semi-gloss", "", sourceFiles.saEastHampton, "Modified 05.25.23");
add("Sant Ambroeus", "East Hampton", "Available", "P 08", "Bus station area", "Millwork walls", "", "Benjamin Moore by GC", "", "BM 2173-10", "Semi-gloss", "", sourceFiles.saEastHampton, "Modified 05.25.23");
add("Sant Ambroeus", "East Hampton", "Available", "P 09", "Hallway", "Bus station", "", "Chateau Domingue", "Belgian satin finish trim paint", "Panier", "Belgian satin", "", sourceFiles.saEastHampton, "Modified 05.25.23");
add("Sant Ambroeus", "East Hampton", "Available", "P 10", "Hallway", "Bathroom doors", "", "Chateau Domingue", "Belgian satin finish trim paint", "Panier", "Belgian satin", "", sourceFiles.saEastHampton, "Modified 05.25.23");
add("Sant Ambroeus", "East Hampton", "Available", "P 11", "Dining room", "Baseboard", "", "Chateau Domingue", "Belgian satin finish trim paint", "Dauphine", "Belgian satin", "", sourceFiles.saEastHampton, "Modified 05.25.23");

add("Sant Ambroeus", "Gelateria", "Available", "", "Ceiling", "Painted drywall", "Painted drywall", "Chateau Domingue Finishes", "Limewash", "Brique", "", "", sourceFiles.saGelateria, "September 2024");
add("Sant Ambroeus", "Gelateria", "Available", "", "Ceiling: interior FOH", "Painted drywall", "Painted drywall", "Benjamin Moore", "AURA", "TBD", "Semi-gloss; brush painted", "", sourceFiles.saGelateria, "September 2024");
add("Sant Ambroeus", "Gelateria", "Available", "", "Ceiling: BOH", "Painted drywall", "Painted drywall", "Benjamin Moore", "AURA", "Cream TBD", "Semi-gloss; brush painted", "", sourceFiles.saGelateria, "September 2024");
add("Sant Ambroeus", "Gelateria", "Available", "", "Ceiling: restroom", "Painted drywall", "Painted drywall", "Chateau Domingue Finishes", "Limewash", "TBD", "", "", sourceFiles.saGelateria, "September 2024");
add("Sant Ambroeus", "Gelateria", "Available", "", "Walls: top wall section through FOH", "Painted drywall", "Painted drywall", "Chateau Domingue", "Limewash", "Espalier", "", "", sourceFiles.saGelateria, "September 2024");
add("Sant Ambroeus", "Gelateria", "Available", "", "Walls: restroom", "Painted drywall", "Painted drywall", "Chateau Domingue", "Limewash", "TBD", "", "", sourceFiles.saGelateria, "September 2024");
add("Sant Ambroeus", "Gelateria", "Available", "", "Walls: FOH counter beam behind POS", "Top of intumescent coating", "", "Benjamin Moore", "AURA", "TBD", "Semi-gloss; brush painted", "", sourceFiles.saGelateria, "September 2024");

add("Sant Ambroeus", "Southampton", "Available", "", "Dining room", "Walls", "", "Farrow and Ball", "", "Dimity No. 2008", "Modern Emulsion", "", sourceFiles.saSouthampton, "");

for (const [location, status, date, source, notes] of santLocations.filter((r) => r[1] === "Missing")) {
  add("Sant Ambroeus", location, status, "", "", "", "", "", "", "", "", notes, source, date);
}

add("Felice", "Hudson", "Available", "", "Restaurant", "Walls", "", "Domingue Finishes", "Limewash", "Arcade", "Limewash", "Was matched with Benjamin Moore.", sourceFiles.feliceHudson, "9/16/24");
add("Felice", "Hudson", "Available", "", "Restaurant", "Wainscoting", "", "Domingue Finishes", "Limewash", "Terre", "Limewash", "Was matched with Benjamin Moore.", sourceFiles.feliceHudson, "9/16/24");
add("Felice", "Hudson", "Available", "", "Restaurant", "Ceilings", "", "Domingue Finishes", "Limewash", "Oyster", "Limewash", "Was matched with Benjamin Moore.", sourceFiles.feliceHudson, "9/16/24");

add("Felice", "64th", "Available", "", "", "Trim", "", "Benjamin Moore", "", "Silver Satin OC-26", "Semi-gloss", "All trim.", sourceFiles.felice64, "");
add("Felice", "64th", "Available", "", "", "Walls / doors", "", "Benjamin Moore", "", "Stormy Sky 1616", "Semi-gloss", "All walls and doors.", sourceFiles.felice64, "");

add("Felice", "71st", "Available", "", "Dining room", "Tin ceiling", "", "Domingue Finishes", "", "Oyster", "Matte / flat", "", sourceFiles.felice71, "");
add("Felice", "71st", "Available", "", "Dining room", "Wood ceiling", "", "Domingue Finishes", "", "Arcade", "Matte / flat", "", sourceFiles.felice71, "");
add("Felice", "71st", "Available", "", "Dining room", "Electrical doors", "", "Behr", "", "Rustic Tobacco", "Matte / flat", "", sourceFiles.felice71, "");
add("Felice", "71st", "Available", "", "Hallway", "Walls / ceiling", "", "Domingue Finishes", "", "Boxwood", "Matte / flat", "", sourceFiles.felice71, "");
add("Felice", "71st", "Available", "", "Lower level", "Tin ceiling", "", "Domingue Finishes", "", "Arcade", "Matte / flat", "", sourceFiles.felice71, "");

add("Felice", "83rd", "Available", "", "", "Upper walls", "", "Benjamin Moore", "", "Turtle Green 2142-20", "Matte", "", sourceFiles.felice83, "");
add("Felice", "83rd", "Available", "", "", "Lower wall near door to Felice bar", "", "Benjamin Moore", "", "Aegean Olive 1491", "Matte", "Lower wall by bar doors.", sourceFiles.felice83, "");
add("Felice", "83rd", "Available", "", "", "High ceiling", "", "Benjamin Moore", "Regal Select Interior Paint", "N547", "Flat", "Color code Y3-0x8.000; S1-0x5.000; R3-0x1.4375.", sourceFiles.felice83, "");
add("Felice", "83rd", "Available", "", "", "Soffit", "", "Benjamin Moore", "", "Stardust 2108-40", "Matte", "", sourceFiles.felice83, "");
add("Felice", "83rd", "Available", "", "", "Bread station", "", "Benjamin Moore", "", "Stardust 2108-40", "Matte", "", sourceFiles.felice83, "");
add("Felice", "83rd", "Available", "", "", "Hallway ceiling", "", "Benjamin Moore", "", "Stardust 2108-40", "Matte", "", sourceFiles.felice83, "");
add("Felice", "83rd", "Available", "", "", "Bathroom ceiling", "", "Benjamin Moore", "", "Stardust 2108-40", "Matte", "", sourceFiles.felice83, "");
add("Felice", "83rd", "Available", "", "", "Bathroom frames and doors", "", "Benjamin Moore", "", "Stardust 2108-40", "Semi", "", sourceFiles.felice83, "");
add("Felice", "83rd", "Available", "", "", "Banquettes", "", "Fine Paints of Europe", "", "G18750", "Matte", "", sourceFiles.felice83, "");
add("Felice", "83rd", "Available", "", "", "Door box / door to Felice bar", "Metal paint", "TBD painter", "", "Dark Black", "", "", sourceFiles.felice83, "");

add("Felice", "Roslyn", "Available", "P 01", "Walls", "Private dining room", "", "Domingue Finishes", "Limewash", "Chapel", "", "Solid painted.", sourceFiles.feliceRoslyn, "");
add("Felice", "Roslyn", "Available", "P 02", "Ceiling", "Throughout", "", "Domingue Finishes", "Limewash", "Terre", "", "Solid painted.", sourceFiles.feliceRoslyn, "");
add("Felice", "Roslyn", "Available", "P 03", "Walls", "Throughout", "", "Domingue Finishes", "Limewash", "Giza; Beeswax; Boxwood; Bastide; Cassis", "", "Paint colors in alternating stripes.", sourceFiles.feliceRoslyn, "");
add("Felice", "Roslyn", "Available", "P 04", "Walls", "Throughout", "", "Domingue Finishes", "Limewash", "Terre", "", "Solid painted.", sourceFiles.feliceRoslyn, "");
add("Felice", "Roslyn", "Available", "P 05", "Walls", "Bathroom", "", "Domingue Finishes", "Final color TBD", "", "", "Solid painted.", sourceFiles.feliceRoslyn, "");

const entryCount = (brand, restaurant) =>
  details.filter((row) => row[0] === brand && row[1] === restaurant && row[2] === "Available").length;

const colLetter = (n) => {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - m) / 26);
  }
  return s;
};

const writeBlock = (sheet, startRow, startCol, values) => {
  const rows = values.length;
  const cols = values[0].length;
  const start = `${colLetter(startCol)}${startRow}`;
  const end = `${colLetter(startCol + cols - 1)}${startRow + rows - 1}`;
  const range = sheet.getRange(`${start}:${end}`);
  range.values = values;
  return range;
};

const styleRange = (range, format) => {
  range.format = { ...range.format, ...format };
};

const workbook = Workbook.create();
const summary = workbook.worksheets.getOrAdd("Summary", { renameFirstIfOnlyNewSpreadsheet: true });
const all = workbook.worksheets.add("All Details");
const saSheet = workbook.worksheets.add("Sant Ambroeus");
const feSheet = workbook.worksheets.add("Felice");
const sourceSheet = workbook.worksheets.add("Sources");

for (const sh of [summary, all, saSheet, feSheet, sourceSheet]) {
  sh.showGridLines = false;
}

summary.getRange("A1:M1").merge();
summary.getRange("A1").values = [["Multi-Restaurant Paint Schedule"]];
styleRange(summary.getRange("A1:M1"), {
  fill: palette.ink,
  font: { bold: true, size: 22, color: palette.white, name: "Aptos Display" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
});
summary.getRange("A2:M2").merge();
summary.getRange("A2").values = [["Normalized from supplied Sant Ambroeus and Felice paint schedule PDFs. Blank cells mean the field was not present in the source document."]];
styleRange(summary.getRange("A2:M2"), {
  fill: palette.paper,
  font: { italic: true, color: palette.muted, size: 10 },
  horizontalAlignment: "center",
});

summary.getRange("A4:F4").merge();
summary.getRange("H4:M4").merge();
summary.getRange("A4").values = [["SANT AMBROEUS"]];
summary.getRange("H4").values = [["FELICE"]];
styleRange(summary.getRange("A4:F4"), {
  fill: palette.saRose,
  font: { bold: true, size: 15, color: palette.ink, name: "Georgia" },
  horizontalAlignment: "center",
  borders: { preset: "outside", style: "thin", color: palette.saGilt },
});
styleRange(summary.getRange("H4:M4"), {
  fill: palette.feliceRed,
  font: { bold: true, size: 15, color: palette.white, name: "Georgia" },
  horizontalAlignment: "center",
  borders: { preset: "outside", style: "thin", color: palette.feliceGold },
});

const summaryHeaders = ["Location", "Status", "Entries", "Date", "Source", "Notes"];
writeBlock(summary, 5, 1, [summaryHeaders, ...santLocations.map((r) => [r[0], r[1], entryCount("Sant Ambroeus", r[0]), r[2], r[3], r[4]])]);
writeBlock(summary, 5, 8, [summaryHeaders, ...feliceLocations.map((r) => [r[0], r[1], entryCount("Felice", r[0]), r[2], r[3], r[4]])]);

styleRange(summary.getRange("A5:F5"), {
  fill: palette.saGilt,
  font: { bold: true, color: palette.white },
  horizontalAlignment: "center",
});
styleRange(summary.getRange("H5:M5"), {
  fill: palette.feliceGreen,
  font: { bold: true, color: palette.white },
  horizontalAlignment: "center",
});
styleRange(summary.getRange(`A6:F${5 + santLocations.length}`), {
  fill: palette.white,
  borders: { preset: "inside", style: "thin", color: palette.line },
  verticalAlignment: "center",
  wrapText: true,
});
styleRange(summary.getRange(`H6:M${5 + feliceLocations.length}`), {
  fill: palette.white,
  borders: { preset: "inside", style: "thin", color: palette.line },
  verticalAlignment: "center",
  wrapText: true,
});
styleRange(summary.getRange("B6:B16"), { horizontalAlignment: "center", font: { bold: true } });
styleRange(summary.getRange("I6:I10"), { horizontalAlignment: "center", font: { bold: true } });
summary.getRange("B6:B8").format.fill = "#FFF176";
summary.getRange("B15:B15").format.fill = "#FFF176";
summary.getRange("B9:B14").format.fill = palette.missing;
summary.getRange("B16:B16").format.fill = palette.missing;
summary.getRange("I6:I10").format.fill = "#DDEFD6";
summary.getRange("A6:A16").format.font = { bold: true, color: palette.ink };
summary.getRange("H6:H10").format.font = { bold: true, color: palette.ink };
summary.getRange("E6:F16").format.font = { size: 9, color: palette.muted };
summary.getRange("L6:M10").format.font = { size: 9, color: palette.muted };

summary.getRange("A19:M19").merge();
summary.getRange("A19").values = [["Source + Missing-Information Rule"]];
styleRange(summary.getRange("A19:M19"), {
  fill: palette.ink,
  font: { bold: true, color: palette.white, size: 12 },
  horizontalAlignment: "center",
});
summary.getRange("A20:M22").merge();
summary.getRange("A20").values = [[
  "Each paint line uses the same columns: brand, restaurant, schedule status, key, room/zone, surface/item, material, vendor, product/type, color, finish, notes, source file, and source date. If a PDF did not include a field, that cell is intentionally blank. Sant Ambroeus missing rows come from the supplied reference image; no comparable missing-location list was supplied for Felice, so the Felice summary includes the Felice schedules present in the provided documents.",
]];
styleRange(summary.getRange("A20:M22"), {
  fill: palette.paper,
  font: { color: palette.muted, size: 10 },
  wrapText: true,
  verticalAlignment: "center",
  borders: { preset: "outside", style: "thin", color: palette.line },
});

summary.getRange("A:A").format.columnWidthPx = 140;
summary.getRange("B:B").format.columnWidthPx = 92;
summary.getRange("C:C").format.columnWidthPx = 68;
summary.getRange("D:D").format.columnWidthPx = 110;
summary.getRange("E:F").format.columnWidthPx = 185;
summary.getRange("G:G").format.columnWidthPx = 28;
summary.getRange("H:H").format.columnWidthPx = 130;
summary.getRange("I:I").format.columnWidthPx = 92;
summary.getRange("J:J").format.columnWidthPx = 68;
summary.getRange("K:K").format.columnWidthPx = 110;
summary.getRange("L:M").format.columnWidthPx = 185;
summary.getRange("A1:M1").format.rowHeight = 34;
summary.getRange("A20:M22").format.rowHeight = 54;
summary.freezePanes.freezeRows(5);

const allValues = [detailHeaders, ...details];
writeBlock(all, 1, 1, allValues);
styleRange(all.getRange("A1:N1"), {
  fill: palette.ink,
  font: { bold: true, color: palette.white },
  horizontalAlignment: "center",
  wrapText: true,
});
styleRange(all.getRange(`A2:N${allValues.length}`), {
  fill: palette.white,
  borders: { preset: "inside", style: "thin", color: palette.line },
  wrapText: true,
  verticalAlignment: "top",
  font: { color: palette.ink, size: 10 },
});
all.getRange(`C2:C${allValues.length}`).format.font = { bold: true };
all.getRange(`A2:A${allValues.length}`).conditionalFormats.addCustom('=$A2="Sant Ambroeus"', {
  fill: palette.saRoseSoft,
  font: { color: palette.ink, bold: true },
});
all.getRange(`A2:A${allValues.length}`).conditionalFormats.addCustom('=$A2="Felice"', {
  fill: palette.feliceStone,
  font: { color: palette.feliceGreen, bold: true },
});
all.getRange(`C2:C${allValues.length}`).conditionalFormats.addCustom('=$C2="Missing"', {
  fill: palette.missing,
  font: { color: palette.missingText, bold: true },
});
all.getRange("A:N").format.columnWidthPx = 130;
all.getRange("E:F").format.columnWidthPx = 190;
all.getRange("H:L").format.columnWidthPx = 170;
all.getRange("M:M").format.columnWidthPx = 220;
all.getRange("N:N").format.columnWidthPx = 120;
all.freezePanes.freezeRows(1);

const writeBrandSheet = (sheet, brand, brandColor, accentColor, rows) => {
  sheet.getRange("A1:N1").merge();
  sheet.getRange("A1").values = [[`${brand.toUpperCase()} PAINT SCHEDULE`]];
  styleRange(sheet.getRange("A1:N1"), {
    fill: brandColor,
    font: { bold: true, size: 18, color: brand === "Felice" ? palette.white : palette.ink, name: "Georgia" },
    horizontalAlignment: "center",
  });
  sheet.getRange("A2:N2").merge();
  sheet.getRange("A2").values = [["Same format across every restaurant; blanks reflect information not shown in the supplied PDF."]];
  styleRange(sheet.getRange("A2:N2"), {
    fill: palette.paper,
    font: { italic: true, color: palette.muted, size: 10 },
    horizontalAlignment: "center",
  });
  writeBlock(sheet, 4, 1, [detailHeaders, ...rows]);
  styleRange(sheet.getRange("A4:N4"), {
    fill: accentColor,
    font: { bold: true, color: palette.white },
    horizontalAlignment: "center",
    wrapText: true,
  });
  styleRange(sheet.getRange(`A5:N${4 + rows.length}`), {
    fill: palette.white,
    borders: { preset: "inside", style: "thin", color: palette.line },
    wrapText: true,
    verticalAlignment: "top",
    font: { color: palette.ink, size: 10 },
  });
  sheet.getRange(`C5:C${4 + rows.length}`).conditionalFormats.addCustom('=$C5="Missing"', {
    fill: palette.missing,
    font: { color: palette.missingText, bold: true },
  });
  sheet.getRange("A:N").format.columnWidthPx = 130;
  sheet.getRange("E:F").format.columnWidthPx = 190;
  sheet.getRange("H:L").format.columnWidthPx = 170;
  sheet.getRange("M:M").format.columnWidthPx = 220;
  sheet.getRange("N:N").format.columnWidthPx = 120;
  sheet.freezePanes.freezeRows(4);
};

writeBrandSheet(saSheet, "Sant Ambroeus", palette.saRose, palette.saGilt, details.filter((r) => r[0] === "Sant Ambroeus"));
writeBrandSheet(feSheet, "Felice", palette.feliceRed, palette.feliceGreen, details.filter((r) => r[0] === "Felice"));

const sourceRows = [
  ["Brand", "Reference", "Use in Workbook"],
  ["Sant Ambroeus", "SA STYLE GUIDE - REV 2020.pdf", "Brand colors: Rosa, Gilt, Blush, Cerulean, Dove White."],
  ["Sant Ambroeus", "CleanShot 2026-04-29 at 21.22.44.png", "Location list and missing-status reference for Sant Ambroeus summary."],
  ["Felice", "FE STYLE GUIDE-July 2020.pdf", "Brand colors: Felice Red, Green, Gold, Fig, Stone, Olive, Slate."],
  ["Both", "Supplied paint schedule PDFs", "Paint schedule data normalized into the Details sheets."],
];
writeBlock(sourceSheet, 1, 1, sourceRows);
styleRange(sourceSheet.getRange("A1:C1"), {
  fill: palette.ink,
  font: { bold: true, color: palette.white },
  horizontalAlignment: "center",
});
styleRange(sourceSheet.getRange("A2:C5"), {
  fill: palette.white,
  borders: { preset: "inside", style: "thin", color: palette.line },
  wrapText: true,
  verticalAlignment: "top",
});
sourceSheet.getRange("A:C").format.columnWidthPx = 240;
sourceSheet.getRange("C:C").format.columnWidthPx = 520;

const summaryCheck = await workbook.inspect({
  kind: "table",
  range: "Summary!A1:M22",
  include: "values,formulas",
  tableMaxRows: 24,
  tableMaxCols: 13,
});
console.log(summaryCheck.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);

await workbook.render({ sheetName: "Summary", range: "A1:M22", scale: 1 });
await workbook.render({ sheetName: "All Details", range: "A1:N25", scale: 1 });
await workbook.render({ sheetName: "Sant Ambroeus", range: "A1:N25", scale: 1 });
await workbook.render({ sheetName: "Felice", range: "A1:N25", scale: 1 });
await workbook.render({ sheetName: "Sources", range: "A1:C5", scale: 1 });

await fs.mkdir(outputDir, { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(`Saved ${outputPath}`);
