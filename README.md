# SP-Manifest
Small library to construct SharePoint manifest files from different partial files.

It can only update App.AppPermissionRequests node.

## Install
Run npm command: 
`npm i https://github.com/Inmeta/sp-manifest --save-dev`

## Methods

#### .setPath(path:string)

Sets the path of AppManifest.xml files. All files has to be in the same directory.
The partial manifest files must have a name like: AppManifest.{build}.xml.

Example:

```
AppManifest.xml             // Main AppManifest
AppManifest.version1.xml    // build type 1 (holds partial manifest elements )
AppManifest.version2.xml    // build type 1 (holds partial manifest elements )
```

#### .getPath()

#### .updatePermissions(build:string, cb:Function:(error))
Updates the main AppManifest file with nodes from the provided `build` manifest.

## Example

```javascript
var manifest = require("sp-manifest");

// update AppPermissionRequests node
// give AppManifest partial name
//      f.ex. AppManifest.AppCatalog.xml

manifest.updatePermissions("AppCatalog", function(error){
    console.log(error || "Done");
})
```