const fs = require('fs'),
    xml2js = require('xml2js'),
    Q = require('q'),
    builder = new xml2js.Builder(),
    parser = new xml2js.Parser();

var MANIFEST_PATH = __dirname;
var buildType = "";
var manifestData = {};

exports.setPath = function(path){
    MANIFEST_PATH = path;
}

exports.getPath = function(){
    return MANIFEST_PATH;
}

exports.updatePermissions = function(build, cb) {
    var deferred = Q.defer();
    buildType = build;
    
    Q.fcall(readManifestFile.bind(null, build))   // Build manifest
    .then(extractAppPermissionsFromManifest)     // gets AppPermissionRequests node 
    .then(saveManifestWithPermissions)
    .catch(function(error) {
        if(cb) return cb(error)
        deferred.reject(error);
    })
    .done(function(){
        deferred.resolve(null);
    });

    return deferred.promise;
}

function readManifestFile(build) {
    var deferred = Q.defer();
    fs.readFile(getManifestPath(build), "utf-8", (err, data) => {
        if(err) return deferred.reject(new Error(err));

        manifestData[build || "main"] = data; 
        deferred.resolve(data);
    });
    return deferred.promise;
}

/**
 * Extracts AppPermissionRequests from XML file
 * @param xmlData XML file as string
 * return{xmlObject, AppPermissionRequests};
 */
function extractAppPermissionsFromManifest() {
    var deferred = Q.defer();
    if(!buildType) return deferred.reject(new Error("No build type defined."));

    parser.parseString(manifestData[buildType], function (err, xmlObject) {
        if(err) return deferred.reject(new Error(err));

        if(xmlObject.App && xmlObject.App.AppPermissionRequests) {
            deferred.resolve(xmlObject.App.AppPermissionRequests)
        } else {
            deferred.reject(new Error("App.AppPermissionRequests not found"));
        }
    });    

    return deferred.promise;
}
function xmlString2JSON(_XMLstring) {
    var deferred = Q.defer();

    parser.parseString(_XMLstring, function (err, xmlObject) {
        if(err) {
            return deferred.reject(new Error(err));
        }

        deferred.resolve(xmlObject);
    });    
    return deferred.promise;
}

function saveManifestWithPermissions(permissionsObject) {
    var deferred = Q.defer();

    readManifestFile()
    .then(xmlString2JSON)
    .then(function(xmlObject){
        xmlObject.App.AppPermissionRequests = permissionsObject;
        var manifestXMLData = builder.buildObject(xmlObject);

        fs.writeFile(getManifestPath(), manifestXMLData, "utf8", function(err){
            if(err) return deferred.reject(new Error(err));
            deferred.resolve();
        })
    });
    return deferred.promise;        
}

function getManifestPath(build) {
    if(build) {
        return MANIFEST_PATH + '/AppManifest.' + build + '.xml'; 
    } else {
        return MANIFEST_PATH + '/AppManifest.xml';
    }
}