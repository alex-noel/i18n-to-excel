# Translation convertor from JS/JSON format into excel spreadsheet

File format in codebase:
 
 ```
const translations = {};

translations.en = {};

translations.fr = {};

export translations;
// or export { translations as somethingElse };
```

Files with translations can be anywhere in folder. Filename should be `translations.js`

Styling for xlsx table made via https://www.npmjs.com/package/xlsx-styles

Sheet creation and data manipulations made via https://www.npmjs.com/package/js-xlsx
