#!/usr/bin/env node
var azure = require("azure");
var fs = require('fs');

var sasurl = process.env.AE_BLOBSERVICE_SAS_URL
var blobURL = sasurl.substring(0, sasurl.indexOf('?'));
var blobCred = sasurl.substring(sasurl.indexOf('?'));
var blobService = azure.createBlobService(null, null, blobURL, blobCred);

exports.put = function(containerName, fileName, blobName, cb) {
  var stats = fs.statSync(fileName);
  var fileSizeInBytes = stats["size"];
  blobService.createBlockBlobFromStream(containerName, blobName, fs.createReadStream(fileName), fileSizeInBytes, function(error, result, response){
    if(error){
        if(cb) {
          cb(null, error);
        }
        else {
          console.log("Couldn't upload file %s", fileName);
          console.error(error);
        }
    } else {
      if(cb) {
        cb("success", null);
      }
      else {
        console.log('File %s uploaded successfully', fileName);
      }
    }
  });
}
