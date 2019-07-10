const flatten = require("flat");
const XLSX = require("xlsx");
XLSXstyle = require("xlsx-style");
const fs = require("fs");
const shell = require("shelljs");
const ospath = require("ospath");

const createSpreadsheet = (data, filepath) => {
  const sheets = mapDataToSheets(data);
  const book = createBook(sheets);
  createFile(book, filepath);
};

const createBook = sheets => {
  const book = XLSX.utils.book_new();

  sheets.forEach(sheet => {
    const xlsSheet = XLSX.utils.aoa_to_sheet(sheet.rows);
    // sheet name cannot exceed 31 char.
    // probably later on we might want to have a single sheet
    // instead of multiple, and store filename as a plain cell
    const sheetName = sheet.name
      .replace(/\/translations.js/gi, "")
      .slice(-31)
      .replace(/[\/\\]/gi, "_");
    xlsSheet["!cols"] = [{ wch: 31 }, { wch: 55 }, { wch: 55 }, { wch: 55 }];
    xlsSheet["!merges"] = [{ s: { c: 0, r: 0 }, e: { c: 4, r: 0 } }];
    const styledSheet = setCellStyles(xlsSheet);
    XLSX.utils.book_append_sheet(book, styledSheet, sheetName);
  });

  return book;
};

const setCellStyles = sheet => {
  Object.keys(sheet).forEach(sheetKey => {
    const [key, cellKey, rowKey] = sheetKey.match(/([A-Z])([0-9].*)/) || [
      null,
      null,
      null
    ];
    if (!sheet[key]) {
      return;
    }
    if (sheet[key].s === undefined) {
      sheet[key].s = {};
    }
    if (cellKey === "A" && rowKey === "1") {
      sheet[key].s.font = {
        bold: true
      };
    }
    if ((cellKey === "A" || parseInt(rowKey) < 3) && rowKey !== "1") {
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

const mapDataToSheets = data => {
  const sheets = data.map(value => {
    const { sheetName: name, translations } = value;
    const rows = [];
    const languages = Object.keys(translations);
    const flattenTranslations = {};

    languages.forEach(lang => {
      flattenTranslations[lang] = flatten(translations[lang]);
    });
    const translationKeys = Object.keys(flattenTranslations.enus);
    // heading row
    rows.push([`FILENAME: ${name}`]);
    rows.push(["Keys", ...languages]);

    translationKeys.forEach(translationKey => {
      const translationRow = [translationKey];
      languages.map(lang => {
        const value = flattenTranslations[lang][translationKey];
        translationRow.push(value || "");
      });
      rows.push(translationRow);
    });
    return {
      name,
      rows
    };
  });

  return sheets;
};

const defaultCellStyle = {
  alignment: {
    wrapText: true
  }
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

  fs.writeFileSync(`${sheetPath}.xlsx`, fileContents, defaultCellStyle);
};

module.exports = {
  createSpreadsheet
};
