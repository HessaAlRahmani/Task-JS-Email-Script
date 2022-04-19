// Your code goes here
const excelToJson = require("convert-excel-to-json");

const result = excelToJson({
  sourceFile: "names.xlsx",
});

//console.log(result);

const sgMail = require("@sendgrid/mail");
const APIkey = "*";

sgMail.setApiKey(APIkey);

let date = new Date();
let year = date.getFullYear();
let month = date.getMonth() + 1;
let day = date.getDate();

for (let i = 0; i < result.Sheet1.length; i++) {
  let name = result.Sheet1[i].A;
  let recpt = result.Sheet1[i].B;
  let course = result.Sheet1[i].C;
  let grade = result.Sheet1[i].D;

  const email = {
    to: "h.m.alrahmani@outlook.com",
    from: "h.m.alrahmani@gmail.com",
    subject: "1, 2, test, test",
    html: `<div class="container" style= "background-img: url("border-design.jpg"); width:100%; display: flex; justify-content: center; align-items: center;">
                  <h1>Certificate of Completion</h1>
                  <p>This certificate is awarded to</p>
                  <h2 style="margin:40px;">${name}</h2>
                  <p>for the secssful completion of</p>
                  <h4 style="margin:20px;">${course}</h4>
                  <h6>with a grade of ${grade} on ${day}/${month}/${year}</h6>
                  <h5>Someone Important</h5>
                  <h6>CEO of the world</h6>
               </div>`,
  };

  sgMail
    .send(email)
    .then((response) => console.log("Sent!"))
    .catch((error) => console.log("error.."));
}
