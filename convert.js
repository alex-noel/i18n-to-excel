const flatten = require("flat");
const XLSX = require("xlsx");
XLSX.style_utils = require("xlsx-style");
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
    xlsSheet["!merges"] = [{ s: { c: 0, r: 0 }, e: { c: 3, r: 0 } }];
    XLSX.utils.book_append_sheet(book, xlsSheet, sheetName);
  });

  return book;
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
    rows.push([`FILENAME:, ${name}`]);
    rows.push(["keys", ...languages]);

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

const createFile = (book, filepath) => {
  const fileContents = XLSX.write(book, {
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
