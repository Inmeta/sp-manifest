## Manipulate SharePoint manifest


```
var manifest = require("./sp-manifest");

// update AppPermissionRequests node
// give AppManifest partial name
//      f.ex. AppManifest.AppCatalog.xml

manifest.updatePermissions("AppCatalog", function(error){
    console.log(error || "Done");
})
```