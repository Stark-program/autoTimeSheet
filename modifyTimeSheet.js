const { range } = require("lodash");
const XlsxPopulate = require("xlsx-populate");

XlsxPopulate.fromFileAsync("./Blank Time Sheet.xlsx").then((workbook) => {
  // sets 40 hour work week values @ 8 hrs a day
  const r1 = workbook.sheet("Front").range("B6:C6");
  const r2 = workbook.sheet("Front").range("F6:J6");
  const r3 = workbook.sheet("Front").range("M6:O6");
  r1.value("8");
  r2.value("8");
  r3.value("8");

  // variables to get correct date format
  var m = new Date().getMonth() + 1;
  var d = new Date().getDate();
  const y = new Date().getFullYear();

  //variables to set dates for timesheet
  const rangeDates = workbook.sheet("Front").range("B4:N4");

  //beginning of pay period
  const payPeriodStartDay = d - 13;
  const payPeriodStartDate = m + "/" + payPeriodStartDay + "/" + y;
  workbook.sheet("Front").cell("I2").value(payPeriodStartDate);
  var lastDayDate = m + "/" + d;
  workbook.sheet("Front").cell("O4").value(lastDayDate);
  workbook
    .sheet("Front")
    .cell("B4")
    .value(m + "/" + payPeriodStartDay);

  //middle pay periods
  workbook
    .sheet("Front")
    .cell("C4")
    .value(m + "/" + payPeriodStartDay);

  //end of pay period
  const payPeriodTurnInDate = new Date().toLocaleDateString();
  workbook.sheet("Front").cell("O2").value(payPeriodTurnInDate);

  return workbook.toFileAsync("./Blank Time Sheet.xlsx");
});
