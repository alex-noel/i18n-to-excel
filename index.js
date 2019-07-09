#!/usr/bin/env node
const inquirer = require("inquirer");
const chalk = require("chalk");
const figlet = require("figlet");
const ora = require("ora");
const shell = require("shelljs");

const { createSpreadsheet } = require("./convert");
const { concatenate } = require("./concatenate");

const spinner = ora({ color: "cyan", hideCursor: false });

const init = () => {
  console.log(chalk`{green.bold Translations file convertion initialized}. 
    Hit CTRL+C to disrupt.`);
};

const askQuestions = () => {
  const questions = [
    {
      name: "SOURCEPATH",
      type: "input",
      message: chalk`Enter absolute path to folder with translations. {gray (Skip to use current directory) }`
    },
    {
      name: "FILEPATH",
      type: "input",
      message: chalk`Enter xls table absolute file path and name (Without extension) {gray (Skip to use current directory) }`
    }
  ];
  return inquirer.prompt(questions);
};

const success = filepath => {
  spinner.stop();
  console.log(
    chalk.white.bgGreen.bold(`Done! File can be found at ${filepath}.xlsx`)
  );
};

const run = async () => {
  init();

  const args = await askQuestions();
  const { FILEPATH, SOURCEPATH } = args;
  const filePath = FILEPATH || `${shell.pwd()}/sheet-${Date.now()}`;
  spinner.start("Processing translations");
  const data = await concatenate(SOURCEPATH);
  createSpreadsheet(data, filePath);

  success(filePath);
};

run();
