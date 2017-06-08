#!/usr/bin/env node
var axios = require("axios");
var Excel = require('exceljs');
var fs = require('fs');
var extract = require('extract-zip');
var violationDict = null;
var scDict = null;
var samples = null;
var baseURL = "http://localhost:46562/";
var acrQidDict = {};
var unknownQid = null;
//copy all the files and download to local
//process xlsx and other files and images/videos
//this will generate a deceptor and call process deceptor api, which will take appId and if it exist in the download storage
//it will process it and create a deceptor which is inspected,
//approver can then view/approve/publish it
//console.log(baseURL);
if(!process.env.AE_BLOBSERVICE_SAS_URL)
{
  console.log(" ");
  console.log("Please set environment variable 'AE_BLOBSERVICE_SAS_URL' to Blob service SAS URL");
  console.log(" ");
  return;
}

var command = "help";
var appId = null;
var appPath = null;
var processDeceptor = "parse";
process.argv.forEach(function (val, index, array) {
  if(index == 2)
  {
    appId = val;
    appPath = "Completed/DeceptorReview/" + appId;
  }
  if(index == 3)
  {
    if(val == "process")  {
      processDeceptor = val;
    }
    else if(val == "generate")  {
      processDeceptor = val;
    }
  }
  /*
  else if(index == 3)
  {
    appPath = val;
  }
  */
});

if(!appId || !appPath)
{
  console.log("To Run the application:");
  console.log("aek7process <appId> (this will basically read the uploaded information and create json in deceptorinterview folder)");
  console.log("aek7process <appId> process (this will read from deceptorinterview folder and create new AppId)");
  console.log("aek7process <appId> generate (this will generate xlsx for application scorecard)");
  return;
}
//read the xlsx and create json for answers
//create a folder for appId similar to required in deceptor page
//mput all the files
//also find how to read sample and source files

function ProcessDeceptorInterview() {
  console.log("processing for deceptor: " + appId +" ....");
  var processURL = baseURL + "deceptor/process?displayId=" + appId;
  axios.get(processURL).then(function(response) {
    console.log("Successfully processed deceptor:" + appId);
  });
}

function uploadMedia(fileName, basePath, citem) {
  var put = require("./put.js");
  put.put("deceptorinterview", fileName, basePath + citem, function(resp, error) {
    if(error) {
      console.log("Error Uploading Media: " + citem);
    }
    else {
      console.log("Uploaded Sucessfully ..." + citem);
    }
  });
}

function getQuesId(fileName) {
  var acrIx = fileName.indexOf("ACR-");
  if(acrIx > -1) {
    var acrNo = fileName.substring(acrIx, acrIx+7);
    if(acrQidDict[acrNo]) {
      return acrQidDict[acrNo];
    }
    else {
      console.log("not found acr: " +  acrNo + " in dict, returning: " + unknownQid)
      return unknownQid;
    }
  }
  else {
    console.log("fileName doesnt contain acr number: " + fileName + " returning: " + unknownQid);
    return unknownQid;
  }
}

function processDir(item) {
  if(item == "ACR-INFO") {
    fs.readdir(__dirname + "/" + appId + "/interview/Images/" + item + "/", function(cerr, citems) {
      if(cerr) {
        console.log("Error copying files in : " + item);
      }
      else {
        for(var j in citems) {
          var citem = citems[j];
          var fileName =  __dirname + "/" + appId + "/interview/Images/" + item + "/" + citem;
          var qid = getQuesId(citem);
          uploadMedia(fileName, appId + "/review/inspect/" + qid + "/", citem);
        }
      }
    });

  }
  else if(item.indexOf("ACR-") > -1) {
    fs.readdir(__dirname + "/" + appId + "/interview/Images/" + item + "/", function(cerr, citems) {
      if(cerr) {
        console.log("Error copying files in : " + item);
      }
      else {
        var qid = getQuesId(item);
        for(var j in citems) {
          var citem = citems[j];
          var fileName =  __dirname + "/" + appId + "/interview/Images/" + item + "/" + citem;
          uploadMedia(fileName, appId + "/review/inspect/" + qid + "/", citem);
        }
      }
    });
  }
}
function CopyMedia() {
  console.log("Copying Images");
  fs.readdir(__dirname + "/" + appId + "/interview/Images/", function(err, items) {
    if(err) {
      console.log(err);
    }
    else {
      for(var i in items) {
        var item = items[i];
        processDir(item);
      }
    }
  });
}

function CopyExecutable() {
  console.log("Copying executables");
  var dir = require("./dir.js");
  dir.dir("downloads", "Completed/DeceptorReview/" + appId + "/*", function(aList){
    if(aList && aList.length > 0) {
      for(var ix=0; ix < aList.length; ix++) {
        var fileName = aList[ix];
        if(fileName.indexOf("Interview.zip") > 0) {
          continue;
        }
        else {
          var copy = require("./copy.js");
          var fName = fileName.split('/').pop();
          console.log("Copying file " + fileName + " as " + "files/" + fName);
          copy.copy("downloads", fileName, "deceptorinterview", appId + "/files/" + fName, function(succ, error){
            if(error) {
              console.log("Error copying file: "  + fileName);
            }
            else {
              console.log("Successfully copied file: "  + fileName);
            }
          })
        }
      }
    }
    else {
        console.log("Error Finding Files");
    }
  });
}

function PutNotes() {
  CopyMedia();

  /*
  var put = require("./put.js");
  var fileName =  "./" + appId + "/Notes.txt";
  put.put("deceptorinterview", fileName, appId + "/Notes.txt", function(resp, error) {
    if(error) {
      console.log("Error Uploading Notes exiting ...");
    }
    else {
      console.log("Uploaded Notes Sucessfully ...");
      CopyMedia();
    }
  });
  */
}

function UploadMetadata() {
  console.log("Uploading Metadata ...")
  var fileName =  __dirname + "/" + appId + "/metadata.json";
  fs.writeFile(fileName, JSON.stringify(samples), "utf8", function(resp){
    var put = require("./put.js");
    put.put("deceptorinterview", fileName, appId + "/metadata.json", function(cresp, error) {
      if(error) {
        console.log("Error Uploading Metadata exiting ...");
      }
      else {
        console.log("Uploaded Metadata Sucessfully ...");
        PutNotes();
      }
    });
  });
}

function UploadViolation() {
//  var get = require("./get.js");
  console.log("Uploading Violations ...")
  var fileName =  __dirname + "/" + appId + "/violations.json";
  fs.writeFile(fileName, JSON.stringify(violationDict), "utf8", function(resp){
    var put = require("./put.js");
    put.put("deceptorinterview", fileName, appId + "/violations.json", function(cresp, error) {
      if(error) {
        console.log("Error Uploading Violations exiting ...");
      }
      else {
        console.log("Uploaded Violations Sucessfully ...");
        UploadMetadata();
      }
    });
  });
}

function ProcessExecs(ws) {
  console.log("Processing MetaData ....")
  var colNameDict = {};
  var maxCol = 27;
  for(var i=1;i<400;i++) {
    //console.log("For : " + i);
    var sample = {};
    for(var j=1;j<maxCol;j++) {
      var colNum = 64+j;
      var colIx = String.fromCharCode(colNum);
      var val = ws.getCell(colIx+i).value;
      if(val && isNaN(val)) {
        val = val.trim();
      }
      //console.log(colIx + i + ":---:" + val);
      if(i==1) {
        if(!val)
        {
          maxCol = j;
          break;
        }

        if(val == "File Name and Path" || val == "fileName") {
          colNameDict[colIx] = "FileName";
        }
        else if(val == "Thumbprint" || val == "digitalCertThumbprint") {
          colNameDict[colIx] = "DigitalCertThumbprint";
        }
        else if(val == "Company Name" || val == "companyName") {
          colNameDict[colIx] = "CompanyName";
        }
        else if(val == "Product Name" || val == "productName") {
          colNameDict[colIx] = "ProductName";
        }
        else if(val == "Product Version" || val == "productVersion") {
          colNameDict[colIx] = "ProductVersion";
        }
        else if(val == "File Version" || val == "fileVersion") {
          colNameDict[colIx] = "FileVersion";
        }
        else if(val == "MD5" || val == "hashMD5") {
          colNameDict[colIx] = "HashMD5";
        }
        else if(val == "SHA1" || val == "hashSHA1") {
          colNameDict[colIx] = "HashSHA1";
        }
        else if(val == "SHA256" || val == "hashSHA256") {
          colNameDict[colIx] = "HashSHA256";
        }
        else if(val == "Issuer Name" || val == "issuerName") {
          colNameDict[colIx] = "IssuerName";
        }
        else if(val == "Issued To" || val == "issuedTo") {
          colNameDict[colIx] = "IssuedTo";
        }
        else {
          console.log("Unknown Header Found :" + val + ", continuing the process");
        }
      }
      else {
        if(!sample["Hash"]) {
          sample["Hash"] = {};
        }
        if(colNameDict[colIx]) {
          if(colNameDict[colIx].indexOf("Hash") == 0) {
            sample["Hash"][colNameDict[colIx].substring(4)] = val;
          }
          else {
            sample[colNameDict[colIx]] = val;
          }
        }
      }
    }
    if(i > 1) {
      if(!samples) {
        samples = [];
      }
      if(sample.FileName){
        samples.push(sample);
      }
    }
  }
  if(!samples) {
    console.log("No MetaData found");
  }
  UploadViolation();
}

function ProcessACRAndExecs(ws, wsExec) {
  console.log("Proccessing Violations ....");
  var colNameDict = {};
  var maxCol = 27;
  for(var i=1;i<400;i++) {
    for(var j=1;j<maxCol;j++) {
      var colNum = 64+j;
      var colIx = String.fromCharCode(colNum);
      var ques = ws.getCell("A"+i).value;
      var val = ws.getCell(colIx+i).value;
      if(!ques || ques == ""){
        break;
      }
      //console.log(colIx + i + ":---:" + val);
      if(i==1) {
        if(!val)
        {
          maxCol = j;
          break;
        }
        colNameDict[colIx] = val;
      }
      else if (j > 2){
        var acrIx = ques.indexOf("ACR-");
        if(val && val != "MET" && val != "NA"  && val != "N/A" && acrIx > -1) {
          var acrNo = ques.substring(acrIx, acrIx+7);
          var quesId = null;
          //console.log("--" + scDict[acrNo] + "--");
          if(scDict[acrNo] && scDict[acrNo][colNameDict[colIx]]) {
            quesId = scDict[acrNo][colNameDict[colIx]];
            if(!acrQidDict[acrNo])
            {
              acrQidDict[acrNo] = quesId;
            }
          }
          if(!quesId) {
            console.log("Error finding questionId for " + acrNo + " and Panel:" + colNameDict[colIx] + ":");
            continue;
          }
          //console.log(acrNo + " ---- " + colNameDict[colIx] + " ----- : " + val + "\n");
          if(!violationDict) {
            violationDict = {};
          }
          if(!violationDict[quesId])
            violationDict[quesId] = val;
          else
            violationDict[quesId] = "\n" + val;
        }
      }
    }
  }
  if(!violationDict) {
    console.log("No Violation found in interview");
  }
  else {
    console.log("Successfully extracted violations");
    ProcessExecs(wsExec);
  }

}

function  ReadInterview() {
  console.log("Reading Interview ....");
  var workbook = new Excel.Workbook();
  workbook.xlsx.readFile(__dirname + "/" + appId + "/interview/" + appId + ".xlsx").then(function() {
    var wsACR = workbook.getWorksheet('ACR_List');
    if(!wsACR) {
      wsACR = workbook.getWorksheet('ACR_ScoreCard');
    }
    if(!wsACR) {
      wsACR = workbook.getWorksheet('ACR_Scorecard');
    }
    if(!wsACR) {
      wsACR = workbook.getWorksheet('Deceptor_List');
    }
    if(!wsACR) {
      wsACR = workbook.getWorksheet('ACR_Details');
    }
    var wsExec = workbook.getWorksheet('Executables');
    if(!wsExec) {
      wsExec = workbook.getWorksheet('Executable');
    }
    if(!wsACR) {
      console.log("Not able to find tab for DeceptorList with Name ACR_List or ACR_ScoreCard or Deceptor_List or ACR_Details ... exiting");
      return;
    }
    if(!wsExec) {
      console.log("Not able to find tab for Executables with Name Executables or Executables ... exiting");
      return;
    }
    ProcessACRAndExecs(wsACR, wsExec);
  });
}

function ReadQuesSchema() {
  var schemaURL = baseURL + "api/schema?schema=DeceptorList";
  axios.get(schemaURL).then(function(response) {
    console.log("Reading schema ....");
    var questions = response.data;
    if(!scDict)
      scDict = {};
    for(var ix in questions) {
      var ques = questions[ix];
      if(ques.answerType != "attest" && !unknownQid) {
        unknownQid = ques.id;
      }
      //all the scorecard question will have answertype as attest
      if(ques.answertype == "attest") {
        var col = ques.panel;
        var rix = ques.questionid.indexOf("ACR-");
        var row = ques.questionid.substring(rix);

        if(!scDict[row]) {
          scDict[row] = {};
        }
        scDict[row][col] = ques.id;
      }
    }
    ReadInterview();
  });
}

function UnZipInterview() {
  console.log("Extracting interview.zip ...");
  extract("./" + appId + "/interview.zip", {dir: __dirname + "/" + appId}, function (err) {
    if(err) {
      console.log("Error extracting interview.zip ...");
      console.log(err);
    }
    else {
      console.log("Successfully extracted interview.zip");
      ReadQuesSchema();
    }
 });
}

function downloadFiles() {
  console.log("downloading interview data ....")
  var get = require("./get.js");
  get.get("downloads", appPath + "/Interview.zip", "./" + appId + "/interview.zip",function(succ, error) {
    if(succ) {
      console.log("Successfully downloaded interview.zip")
      get.get("downloads", appPath + "/Notes.txt", "./" + appId + "/Notes.txt", function(csucc, cerror){
        if(csucc) {
          console.log("Successfully downloaded Notes.txt");
          //now downloaded
        }
        else {
          console.log("Error downloading Notes.txt ....");
          console.log(error);
        }
        UnZipInterview();
      });
    }
    else {
      console.log("Error downloading interview.zip exiting ....");
      console.log(error);
    }
  });
}

if(processDeceptor == "process") {
  console.log("Processing K7 data of deceptor and creating deceptor for portal");
  ProcessDeceptorInterview();
}
else if(processDeceptor == "generate") {
  console.log("Generating xlsx for appId");
  var generate = require("./generate.js");
  generate.generate(appId, baseURL);
}
else {
  console.log("Parsing K7 data of xlsx");
  if (!fs.existsSync("./" + appId)){
      fs.mkdirSync("./" + appId);
  }

  CopyExecutable();
  downloadFiles();
}
