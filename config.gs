const main = SpreadsheetApp.getActiveSpreadsheet();
const specOne = main.getSheetByName("Главспец АИ (Макулова)");
const svodSheet = main.getSheetByName("Сводная загрузка");
const objectColumn = 2; // B
const startDateCol = 8; // H
const endDateCol = 9; // I
const startDateLine = 16; // P
const namesCol = 13; // M
const datesLine = specOne.getRange(3, 16, 1, specOne.getLastColumn()).getValues()[0].map(date => new Date(date)).map(date => date.getDate() + "." + date.getMonth() + "." + date.getFullYear());
