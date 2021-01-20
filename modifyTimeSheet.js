const { range } = require("lodash");
const XlsxPopulate = require("xlsx-populate");
const cron = require("node-cron");
var nodemailer = require("nodemailer");
require("dotenv").config({ path: "./secret.env" });

const autoTimeSheet = () => {
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
      .value(m + "/" + (d - 12));
    workbook
      .sheet("Front")
      .cell("D4")
      .value(m + "/" + (d - 11));
    workbook
      .sheet("Front")
      .cell("E4")
      .value(m + "/" + (d - 10));
    workbook
      .sheet("Front")
      .cell("F4")
      .value(m + "/" + (d - 9));
    workbook
      .sheet("Front")
      .cell("G4")
      .value(m + "/" + (d - 8));
    workbook
      .sheet("Front")
      .cell("H4")
      .value(m + "/" + (d - 7));
    workbook
      .sheet("Front")
      .cell("I4")
      .value(m + "/" + (d - 6));
    workbook
      .sheet("Front")
      .cell("J4")
      .value(m + "/" + (d - 5));
    workbook
      .sheet("Front")
      .cell("K4")
      .value(m + "/" + (d - 4));
    workbook
      .sheet("Front")
      .cell("L4")
      .value(m + "/" + (d - 3));
    workbook
      .sheet("Front")
      .cell("M4")
      .value(m + "/" + (d - 2));
    workbook
      .sheet("Front")
      .cell("N4")
      .value(m + "/" + (d - 1));

    //end of pay period
    const payPeriodTurnInDate = new Date().toLocaleDateString();
    workbook.sheet("Front").cell("O2").value(payPeriodTurnInDate);

    return workbook.toFileAsync("./Blank Time Sheet.xlsx");
  });
};

const sendEmail = () => {
  console.log(user);
  console.log(from);
  var transporter = nodemailer.createTransport({
    service: "gmail",
    auth: { user: process.env.user, pass: process.env.pass },
  });

  const mailOptions = {
    from: process.env.from,
    to: process.env.to,

    subject: "Timesheet",
    attachments: [
      {
        filename: "Blank Time Sheet.xlsx",
        path: "./Blank Time Sheet.xlsx",
      },
    ],
  };

  transporter.sendMail(mailOptions, function (err, info) {
    if (err) {
      console.log(err);
    } else {
      console.log(info);
    }
  });
};
cron.schedule("* 8 * * 3", () => {
  autoTimeSheet();
  sendEmail();
  var x = new Date();
  console.log(x);
});
