// Your code goes here
require("dotenv").config();
const fs = require("fs");
const nodemailer = require("nodemailer");
const excelToJson = require("convert-excel-to-json");

const MAILTRAP_USER = process.env.MAILTRAP_USER;
const MAILTRAP_PASSWORD = process.env.MAILTRAP_PASSWORD;
const GMAIL_PASSWORD = process.env.GMAIL_PASSWORD;
const EMAIL = "user.alamiri@gmail.com";
const TEMPLATE_PATH = "./email.html";
const NAMES = excelToJson({
  sourceFile: "names.xlsx",
  header: {
    rows: 1,
  },
}).Sheet1;

const fileContent = (path) => fs.readFileSync(path, "utf8");
const getEmailAndHtml = (person) => {
  const courseNames = {
    js: "Javascript",
    html: "HTML",
    css: "CSS",
  };
  [title, email, course, grade] = Object.values(person);
  let html = fileContent(TEMPLATE_PATH);

  html = html.replace("__", title);
  html = html.replace("__", courseNames[course.toLowerCase()]);
  html = html.replace("__", grade);
  html = html.replace("__", new Date().toLocaleDateString("en-GB"));
  return [email, html];
};

const transport = {
  test: nodemailer.createTransport({
    host: "smtp.mailtrap.io",
    port: 2525,
    auth: {
      user: MAILTRAP_USER,
      pass: MAILTRAP_PASSWORD,
    },
  }),
  real: nodemailer.createTransport({
    service: "gmail",
    auth: {
      user: EMAIL,
      pass: GMAIL_PASSWORD,
    },
  }),
};

// main function
const sendEmail = (email, content, type) => {
  const valid = type === "real" || type === "test";
  if (!valid) return;

  const mailOptions = {
    from: EMAIL,
    to: email,
    subject: "CODED Task",
    html: content,
  };

  transport[type].sendMail(mailOptions, function (error, info) {
    if (error) {
      console.log(error);
    } else {
      console.log("Email sent: " + info.response);
    }
  });
};

const main = () => {
  NAMES.forEach((person) => {
    [email, html] = getEmailAndHtml(person);
    sendEmail(email, html, "test");
  });
};

main();
