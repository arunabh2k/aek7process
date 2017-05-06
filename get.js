#!/usr/bin/env node
var azure = require("azure");
var fs = require('fs');
var sasurl = process.env.AE_BLOBSERVICE_SAS_URL
var blobURL = sasurl.substring(0, sasurl.indexOf('?'));
var blobCred = sasurl.substring(sasurl.indexOf('?'));
var blobService = azure.createBlobService(null, null, blobURL, blobCred);

exports.get = function(containerName, blobName, fileName, cb) {
  blobService.getBlobToStream(containerName, blobName, fs.createWriteStream(fileName), function(error, result, response){
    if (error) {
        if(cb) {
          cb(null, error);
        }
        else {
          console.error("Couldn't download blob %s", blobName);
          console.error(error);
        }
    } else {
        if(cb) {
          cb("succes", null);
        }
        else {
          console.log("Sucessfully downloaded blob %s to %s", blobName, fileName);
        }
    }
  });
}
