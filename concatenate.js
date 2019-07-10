const shell = require("shelljs");
const tmp = require("tmp");
const Promise = require("bluebird");
const fs = require("fs");
const vm = require("vm");
const getTranslationsList = PATH => {
  const array = shell.find(PATH || shell.pwd()).filter(file => {
    return file.match(/translations\.js$/);
  });
  return array;
};

const createTempFiles = filePathList => {
  const tmpFolder = tmp.dirSync();
  filePathList.forEach((filePath, key) => {
    const newPath = `${tmpFolder.name}/${key}-translation.js`;
    shell.cp(filePath, newPath);
    shell.ShellString(`export const filePath = '${filePath}';`).toEnd(newPath);
  });
  const newList = shell.find(`${tmpFolder.name}/*.js`);
  return tmpFolder;
};

const preprocessJs = async tmpFolder => {
  const newList = shell.find(`${tmpFolder.name}/*.js`);
  await Promise.map(newList, filePath => {
    return new Promise((resolve, reject) => {
      shell.exec(
        `npx ${__dirname}/node_modules/.bin/babel ${filePath} --out-file ${filePath} --presets=@babel/preset-env`,
        { async: true },
        (code, stdout, stderr) => {
          if (code != 0) return reject(new Error(stderr));
          return resolve(stdout);
        }
      );
    });
  });
  return newList;
};

const readData = filePathList => {
  let data = filePathList.map(filePath => {
    const fileSourceCode = fs.readFileSync(filePath);
    const sandbox = {
      exports: {}
    };
    const scope = vm.runInNewContext(fileSourceCode, sandbox);
    const sheet = sandbox.exports;
    const translationKey = Object.keys(sheet).filter(key => key !== "filePath");
    return {
      sheetName: sheet.filePath,
      translations: sheet[translationKey[0]]
    };
  });
  return data;
};

const concatenate = async PATH => {
  const fileList = getTranslationsList(PATH);
  const tmpFolder = createTempFiles(fileList);
  const esFileList = await preprocessJs(tmpFolder);
  const data = readData(esFileList);

  tmpFolder.removeCallback();
  return data;
};

module.exports = {
  concatenate
};
