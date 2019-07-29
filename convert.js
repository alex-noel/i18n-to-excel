const flatten = require("flat");
const XLSX = require("xlsx");
XLSXstyle = require("xlsx-style");
const fs = require("fs");
const shell = require("shelljs");
const ospath = require("ospath");

const createSpreadsheet = (data, filepath) => {
  const { startCoords, sheet } = mapDataToSheet(data);
  const book = createBook(sheet, startCoords);
  createFile(book, filepath);
};

const createBook = (sheet, startCoords) => {
  const book = XLSX.utils.book_new();

  const xlsSheet = XLSX.utils.aoa_to_sheet(sheet);

  const sheetName = "Translations";
  xlsSheet["!cols"] = [
    { wch: 31 },
    { wch: 55 },
    { wch: 55 },
    { wch: 55 },
    { wch: 55 }
  ];
  xlsSheet["!merges"] = startCoords.map(coordinate => {
    return { s: { c: 0, r: coordinate }, e: { c: 4, r: coordinate } };
  });

  const styledSheet = setCellStyles(xlsSheet, startCoords);
  XLSX.utils.book_append_sheet(book, styledSheet, sheetName);
  return book;
};

const setCellStyles = (sheet, startCoords) => {
  let prevStartFileCoord = 0;
  Object.keys(sheet).forEach(sheetKey => {
    const [key, cellKey, rowKey] = sheetKey.match(/([A-Z])([0-9].*)/) || [
      null,
      null,
      null
    ];
    if (!sheet[key]) {
      return;
    }

    const startOfFileCoord = startCoords.indexOf(rowKey - 1);
    const isStartOfFile = startOfFileCoord >= 0;
    prevStartFileCoord = isStartOfFile
      ? startCoords[startOfFileCoord] + 1
      : prevStartFileCoord;
    const isHeader = parseInt(rowKey) === prevStartFileCoord + 1;

    if (sheet[key].s === undefined) {
      sheet[key].s = {};
    }
    if (cellKey === "A" || isStartOfFile || isHeader) {
      sheet[key].s.font = {
        bold: true
      };
    }
    if ((cellKey === "A" || isHeader) && !isStartOfFile) {
      sheet[key].s.fill = {
        patternType: "solid",
        fgColor: {
          rgb: "00EEEEEE"
        }
      };
      const borderStyle = {
        style: "thin",
        color: {
          rgb: "00000000"
        }
      };
      sheet[key].s.border = {
        top: borderStyle,
        bottom: borderStyle,
        left: borderStyle,
        right: borderStyle
      };
    }
    sheet[key].s.alignment = {
      wrapText: true
    };
  });

  return sheet;
};

const mapDataToSheet = data => {
  const sheet = [];
  let sheetLength = 0;
  const startCoords = [];
  data.forEach(value => {
    const { sheetName: name, translations } = value;
    const languages = Object.keys(translations);
    const flattenTranslations = {};

    languages.forEach(lang => {
      flattenTranslations[lang] = flatten(translations[lang]);
    });
    const translationKeys = Object.keys(flattenTranslations.enus);
    // heading row
    startCoords.push(sheetLength);
    sheet.push([`FILENAME: ${name}`]);
    sheetLength++;
    sheet.push(["Keys \\ Languages", ...languages]);
    sheetLength++;

    translationKeys.forEach(translationKey => {
      const translationRow = [translationKey];
      languages.map(lang => {
        const value = flattenTranslations[lang][translationKey];
        translationRow.push(value || "");
      });
      sheet.push(translationRow);
      sheetLength++;
    });
  });

  return { startCoords, sheet };
};

const createFile = (book, filepath) => {
  const fileContents = XLSXstyle.write(book, {
    type: "buffer",
    bookType: "xlsx",
    bookSST: false
  });

  let sheetPath;
  if (filepath) {
    sheetPath = filepath.replace(/^~/, ospath.home());
  }

  fs.writeFileSync(`${sheetPath}.xlsx`, fileContents);
};

module.exports = {
  createSpreadsheet
};
