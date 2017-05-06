#!/usr/bin/env node
var azure = require("azure");
var blobs = [];
var sasurl = process.env.AE_BLOBSERVICE_SAS_URL
var blobURL = sasurl.substring(0, sasurl.indexOf('?'));
var blobCred = sasurl.substring(sasurl.indexOf('?'));
var blobService = azure.createBlobService(null, null, blobURL, blobCred);

function matchRule(str, rule) {
  return new RegExp("^" + rule.split("*").join(".*") + "$").test(str);
}

function aggregateBlobs(containerName, err, result, cb) {
    if (err) {
        cb(err);
    } else {
        blobs = blobs.concat(result.entries);
        if (result.continuationToken !== null) {
            blobService.listBlobsSegmented(
                containerName,
                result.continuationToken,
                aggregateBlobs);
        } else {
            cb(null, blobs);
        }
    }
}
exports.dir = function(containerName, blobPattern, cb) {
  var list = [];
  blobService.listBlobsSegmented(containerName, null, function(err, result) {
      aggregateBlobs(containerName, err, result, function(err, blobs) {

          if (err) {
            if(cb) {
              return cb([]);
            }
            else {
              console.log("Couldn't list blobs");
              console.error(err);
            }
          } else {
              for(var ix in blobs)
              {
                if(!blobPattern || blobPattern == "")
                {
                  list.push(blobs[ix].name);
                }
                else if(blobPattern && matchRule(blobs[ix].name,blobPattern))
                {
                  list.push(blobs[ix].name);
                }
              }
              if(cb) {
                return cb(list);
              }
              else {
              console.log(list);
              }
          }
      });
  });
}
