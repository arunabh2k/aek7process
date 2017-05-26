var axios = require("axios");
var Excel = require('exceljs');
var setHeader = function(colRef, name) {
  colRef.value = name;
  colRef.font = {size: 10,bold: true,color: { argb: '00f2f2f2' }};
  colRef.fill = {type: 'pattern',pattern:'solid',fgColor:{argb:'001c75bc'}};
  colRef.border = {
    top: {style:'thick', color: {argb:'00e6e6e6'}},
    left: {style:'thick', color: {argb:'00e6e6e6'}},
    bottom: {style:'thick', color: {argb:'00e6e6e6'}},
    right: {style:'thick', color: {argb:'00e6e6e6'}}
  };
}

var setQuestion = function(colRef, name) {
  colRef.value = name;
  colRef.border = {
    top: {style:'thick', color: {argb:'00e6e6e6'}},
    left: {style:'thick', color: {argb:'00e6e6e6'}},
    bottom: {style:'thick', color: {argb:'00e6e6e6'}},
    right: {style:'thick', color: {argb:'00e6e6e6'}}
  };
  colRef.fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb:'00d9d9d9'}
  };
}

exports.generate = function(appId, baseUrl) {
  var schemaURL = baseUrl + "api/schema?schema=DeceptorList";
  console.log("Fetching Schema from:" + schemaURL);
  axios.get(schemaURL).then(function(response) {
    console.log("Reading schema ....");
    var questions = response.data;
    var scDict = {};
    var rowQues = {};
    var colpos = {};
    var otherQues = {};
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
      else {
        otherQues[ques.id] = ques;
      }
    }

    console.log("Creating excel ....");
    var workbook = new Excel.Workbook();
    var worksheet = workbook.addWorksheet('ACR_List');
    worksheet.views = [
        {state: 'frozen', xSplit: 2, ySplit: 1}
    ];
    //static header
    console.log("Creating excel Static Header");
    worksheet.getCell('A1').value = "ACR/Details";
    worksheet.getCell('B1').value = "Category";
    worksheet.getCell('A1').font = {size: 10,bold: true,color: { argb: '00f2f2f2' }};
    worksheet.getCell('A1').fill = {type: 'pattern',pattern:'solid',fgColor:{argb:'001c75bc'}};
    worksheet.getCell('B1').font = {size: 10,bold: true,color: { argb: '00f2f2f2' }};
    worksheet.getCell('B1').fill = {type: 'pattern',pattern:'solid',fgColor:{argb:'001c75bc'}};
    worksheet.getColumn(1).width = 40;
    worksheet.getColumn(1).alignment = {wrapText: true };
    worksheet.getColumn(2).width = 12;
    worksheet.getColumn(2).alignment = {wrapText: true };

    //borders
    worksheet.getCell('A1').border = {
      top: {style:'thick', color: {argb:'00e6e6e6'}},
      left: {style:'thick', color: {argb:'00e6e6e6'}},
      bottom: {style:'thick', color: {argb:'00e6e6e6'}},
      right: {style:'thick', color: {argb:'00e6e6e6'}}
    };
    worksheet.getCell('B1').border = {
      top: {style:'thick', color: {argb:'00e6e6e6'}},
      left: {style:'thick', color: {argb:'00e6e6e6'}},
      bottom: {style:'thick', color: {argb:'00e6e6e6'}},
      right: {style:'thick', color: {argb:'00e6e6e6'}}
    };
    //questions and conditions

    console.log("Creating excel questionNames");
    var ax = 2;
    for(var ix in rowQues) {
      worksheet.getCell('A'+ax).value = rowQues[ix].question;
      worksheet.getCell('B'+ax).value = rowQues[ix].conditionalquestion;

      //borders
      worksheet.getCell('A'+ax).border = {
        top: {style:'thick', color: {argb:'00e6e6e6'}},
        left: {style:'thick', color: {argb:'00e6e6e6'}},
        bottom: {style:'thick', color: {argb:'00e6e6e6'}},
        right: {style:'thick', color: {argb:'00e6e6e6'}}
      };
      worksheet.getCell('B'+ax).border = {
        top: {style:'thick', color: {argb:'00e6e6e6'}},
        left: {style:'thick', color: {argb:'00e6e6e6'}},
        bottom: {style:'thick', color: {argb:'00e6e6e6'}},
        right: {style:'thick', color: {argb:'00e6e6e6'}}
      };
      ax++;
    }

    console.log("Creating excel panelName and corresponding value");
    var rx = 3;
    for(var jx in colpos) {
      var colNum = 64+rx;
      var colIx = String.fromCharCode(colNum);
      var colName = colpos[jx];
      worksheet.getCell(colIx+'1').value = colName;
      worksheet.getCell(colIx+'1').font = {size: 10,bold: true,color: { argb: '00f2f2f2' }};
      worksheet.getCell(colIx+'1').fill = {type: 'pattern',pattern:'solid',fgColor:{argb:'001c75bc'}};
      worksheet.getColumn(rx).width = 12;
      worksheet.getColumn(rx).alignment = {wrapText: true };

      //borders
      worksheet.getCell(colIx+'1').border = {
        top: {style:'thick', color: {argb:'00e6e6e6'}},
        left: {style:'thick', color: {argb:'00e6e6e6'}},
        bottom: {style:'thick', color: {argb:'00e6e6e6'}},
        right: {style:'thick', color: {argb:'00e6e6e6'}}
      };

      var bx = 2;
      for(var ix in rowQues) {
        if(scDict[ix][jx]) {
          //here we will just fill with green color to identify that question needs to be answered
          worksheet.getCell(colIx+bx).fill = {
            type: 'pattern',
            pattern:'solid',
            fgColor:{argb:'0092d032'}
          };
        }
        else {
          worksheet.getCell(colIx+bx).fill = {
            type: 'pattern',
            pattern:'solid',
            fgColor:{argb:'00d9d9d9'}
          };
        }
        worksheet.getCell(colIx+bx).border = {
          top: {style:'thick', color: {argb:'00e6e6e6'}},
          left: {style:'thick', color: {argb:'00e6e6e6'}},
          bottom: {style:'thick', color: {argb:'00e6e6e6'}},
          right: {style:'thick', color: {argb:'00e6e6e6'}}
        };
        bx++;
      }
      rx++;
    }
    console.log("Creating Interview Question Sheet");
    var iQues = workbook.addWorksheet('Interview_Question');
    workbook.addWorksheet('Missed_ACR');
    workbook.addWorksheet('Executables');
    workbook.addWorksheet('Queries');
    //now the question sheet for other questions apart from scorecard
    console.log("Creating Interview Question Sheet 1");
    setHeader(iQues.getCell('A1'), "Id");
    iQues.getColumn(1).width = 40;
    iQues.getColumn(1).alignment = {wrapText: true };

    console.log("Creating Interview Question Sheet 2");
    setHeader(iQues.getCell('B1'), "Question");
    iQues.getColumn(2).width = 40;
    iQues.getColumn(2).alignment = {wrapText: true };

    setHeader(iQues.getCell('C1'), "Answer");
    iQues.getColumn(3).width = 40;
    iQues.getColumn(3).alignment = {wrapText: true };

    var qx=2;
    console.log("Creating Interview Question Sheet 3");
    for(var id in otherQues) {
      setQuestion(iQues.getCell('A' +qx), id);
      setQuestion(iQues.getCell('B' +qx), otherQues[id].question);
      qx++;
    }

    console.log("Writing file ....");
    var fileName = "./" + appId + ".xlsx";
    workbook.xlsx.writeFile(fileName)
    .then(function() {
        console.log("Written Scorecard in: " + fileName);
    })
    .error(function(error){
      console.log("something wrong: " + error);
    });
  });
}
