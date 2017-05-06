#!/usr/bin/env node
var azure = require("azure");
var fs = require('fs');
var sasurl = process.env.AE_BLOBSERVICE_SAS_URL
var blobURL = sasurl.substring(0, sasurl.indexOf('?'));
var blobCred = sasurl.substring(sasurl.indexOf('?'));
var blobService = azure.createBlobService(null, null, blobURL, blobCred);

exports.copy = function(containerName, fromBlobName, toContainerName, toBlobName, cb) {
  var fromBlobUrl = blobService.getUrl(containerName, fromBlobName, null);
  fromBlobUrl += blobCred;
  blobService.startCopyBlob(fromBlobUrl, toContainerName, toBlobName, null, function(error, result, response){
    if (error) {
      if(cb) {
        cb(null, error);
      }
      else {
        console.error("Couldn't copy blob %s to (%s)", fromBlobName, toBlobName);
        console.error(error);
      }
    } else {
      if(cb) {
        cb("success", null);
      }
      else {
        console.error("Success copy blob %s to (%s)", fromBlobName, toBlobName);
      }
    }
  });
}
