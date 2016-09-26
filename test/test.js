var expect = require("chai").expect;
describe('Manifest', function() {
  describe('#updatePermissions AppCatalog Version', function() {
    it('shold update with FullControl permissions', function(done) {
        var manifest = require('./../index');
        manifest.updatePermissions("AppCatalog").done(function(d){
            expect(d).to.include("FullControl");
            done();
        })
    });
  });

  describe('#updatePermissions OfficeVersion', function() {
    it('shold update with Office permissions', function(done) {
        var manifest = require('./../index');
        manifest.updatePermissions("OfficeStore").done(function(d){
            expect(d).to.include("Manage");
            done();
        })
    });
  });
});
