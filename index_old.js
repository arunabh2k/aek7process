var axios = require("axios");
var Excel = require('exceljs');

var schemaURL = "https://customer.appesteem.com/api/schema";
axios.get(schemaURL).then(function(response) {
  console.log("Reading schema ....");
  var questions = response.data;
  var scDict = {};
  var rowQues = {};
  var colpos = {};
  for(var ix in questions) {
    var ques = questions[ix];
    //all the scorecard question will have answertype as attest
    if(ques.answertype == "attest") {
      var col = ques.panel;
      var rix = ques.questionid.indexOf("ACR-");
      var row = ques.questionid.substring(rix);

      if(!scDict[row]) {
        scDict[row] = {};
      }
      //storing column name, row information and matrix information
      colpos[col] = col;
      rowQues[row] = ques;
      scDict[row][col] = ques;
    }
  }

  console.log("Creating excel ....");
  var workbook = new Excel.Workbook();
  var worksheet = workbook.addWorksheet('Scorecard');

  //static header
  console.log("Creating excel Static Header");
  worksheet.getCell('A1').value = "ACR/Details";
  worksheet.getCell('B1').value = "Category";
  //questions and conditions

  console.log("Creating excel questionNames");
  var ax = 2;
  for(var ix in rowQues) {
    worksheet.getCell('A'+ax).value = rowQues[ix].question;
    worksheet.getCell('B'+ax).value = rowQues[ix].conditionalquestion;
    ax++;
  }

  console.log("Creating excel panelName and corresponding value");
  var rx = 3;
  for(var jx in colpos) {
    var colNum = 64+rx;
    var colIx = String.fromCharCode(colNum);
    var colName = colpos[jx];
    worksheet.getCell(colIx+'1').value = colName;

    var bx = 2;
    for(var ix in rowQues) {
      if(scDict[ix][jx]) {
        //here we will just fill with green color to identify that question needs to be answered
        worksheet.getCell(colIx+bx).fill = {
          type: 'pattern',
          pattern:'darkVertical',
          fgColor:{argb:'00009900'}
        };
      }
      bx++;
    }
    rx++;
  }
  console.log("Writing file ....");
  var fileName = "./scorecard.xlsx";
  workbook.xlsx.writeFile(fileName)
  .then(function() {
      console.log("Written Scorecard in: " + fileName);
  })
  .error(function(error){
    console.log("something wrong: " + error);
  });
});
