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
    html: `<div class="container" style= "width: 100%; height: 100; padding: 600px; display: inline; justify-content: center; text-align: center;">
                <h1 style="border-top-style: double; padding-top: 180px; font-size: xx-large;"><strong> Certificate of Completion</strong></h1>
                <p><strong> certificate is awarded to </strong></p>
                <h2 style="margin:40px;">${name}</h2>
                <h4>for the secssful completion of</h4>
                <h3 style="margin:20px;">${course}</h3>
                <h4 style="margin-bottom: 80px;">with a grade of ${grade} on ${day}/${month}/${year}</h4>
                <h5>Someone Important</h5>
                <h6 style="border-bottom-style: double; padding-bottom: 100px;">CEO of the world</h6>
            </div>`,
  };

  sgMail
    .send(email)
    .then((response) => console.log("Sent!"))
    .catch((error) => console.log("error.."));
}
