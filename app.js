var mysql = require('mysql2');
const Excel = require('exceljs');
const ObjectsToCsv = require('objects-to-csv');
var fs = require('fs');
var moment = require('moment');

const connection = mysql.createConnection({
  host: '127.0.0.1',
  user: 'root',
  password: 'root',
  database: 'tc-db-main',
});

connection.connect(function (err) {
  if (err) {
    return console.error('Ошибка: ' + err.message);
  } else {
    console.log('Подключение к серверу MySQL успешно установлено');
  }
});

connection.query(
  "select DISTINCT `tc-db-main`.`personal`.`ID`, `NAME`, `POS`, `TABID`, `tc-db-log`.`logs`.`LOGTIME`, `LOGDATA` from `tc-db-main`.`personal` inner join  `tc-db-log`.`logs` On (`tc-db-main`.`personal`.`ID` = `tc-db-log`.`logs`.`EMPHINT`) where `tc-db-log`.`logs`.`LOGTIME` >= '2021-03-02T00:00:00.000' AND `tc-db-log`.`logs`.`LOGTIME` <= '2021-03-02T23:59:59.000' AND LENGTH(`tc-db-log`.`logs`.`LOGDATA`) >= 21",
  function (err, results) {
    if (err) {
      return console.log('Ошибка: ' + err.message);
    }

    const fullArray = [];
    const users = results;

    for (let i = 0; i < users.length; i++) {
      // console.log('LOGDATA', users[i]);
      //   console.log(`object`, users[i]?.LOGDATA.toString('hex'));
      let cell = users[i]?.LOGDATA.toString('hex').substr(18, 8);

      fullArray.push({
        id: i + 1,
        name: users[i].NAME,
        pos: users[i].POS,
        tabId: users[i].TABID,
        date: moment(users[i].LOGTIME).format('D-MM-YYYY hh:mm'),
        temp: parseInt(cell, 16) / 10,
      });
    }

    // recordToFile(fullArray);
    toExcel(fullArray);

    // console.log(`fullArray`, fullArray);
  }
);

const recordToFile = async (fullArray) => {
  const csv = new ObjectsToCsv(fullArray, { bom: false });
  await csv.toDisk('./list.csv');
};

const toExcel = (data) => {
  let workbook = new Excel.Workbook();
  let worksheet = workbook.addWorksheet('Журнал');

  worksheet.mergeCells('A1:h1');
  worksheet.getRow(1).values = [
    'Журнал регистрации измерения температуры работников для профилактики коронавируса',
  ];

  // header

  const title = worksheet.getRow(2);

  title.values = [
    '№ п/п',
    'Дата измерения',
    'ФИО Работника',
    'Должность',
    'Табельный номер',
    'Температура',
    'Подпись',
    'ФИО должность работника, проводившего измерения температуры',
  ];

  worksheet.columns = [
    { key: 'id' },
    { key: 'date', width: 15 },
    { key: 'name', width: 25 },
    { key: 'pos', width: 15 },
    { key: 'tabId', width: 10 },
    { key: 'temp', width: 10 },
    { key: 'podpis', width: 15 },
    { key: 'fioPodpis', width: 15 },
  ];

  // worksheet.columns.forEach((column) => {
  //   column.width = column.header.length < 12 ? 12 : column.header.length;
  // });

  worksheet.getRow(2).font = { bold: true };

  data.forEach((e, index) => {
    // row 1 is the header.
    const rowIndex = index + 1;

    worksheet.addRow({
      ...e,
      amountRemaining: {
        formula: `=C${rowIndex}-D${rowIndex}`,
      },
      percentRemaining: {
        formula: `=E${rowIndex}/C${rowIndex}`,
      },
    });
  });

  workbook.xlsx.writeFile('excel.xlsx');
};

connection.end(function (err) {
  if (err) {
    return console.log('Ошибка: ' + err.message);
  }
  console.log('Подключение закрыто');
});
