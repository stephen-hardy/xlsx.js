var assert = require('chai').assert;
var xlsx = require('../xlsx');
var fs = require('fs-extra');
var path = require('path');

describe('XlsX.js unit tests', function() {

  var file = path.join('test', 'output', 'inflate-1.xlsx');
  
  before(function(done) {
    fs.remove(file, function(err) {
      if (err) {
        return done(err);
      }
      fs.createFile(file, done);
    });
  })

  it('should simple xlsx be written (you can manually check '+file+')', function(done) {
    var sheet = xlsx({
      creator: 'John Doe',
      lastModifiedBy: 'Meg White',
      worksheets: [{
        data: [
          ['green', 'white', {value:'orange', autoWidth:true}, 'blue', 'red'],
          [{formatCode: '0.00', value:'1'}, {italic:1, bold:1, hAlign:'center', value:'2', borders:{bottom:'DEE31D'}}, {value:'3', borders:{bottom:64}}, '4', '5'],
          [{fontName: 'Arial', value:'6'}, {fontSize: 8, value:'7'}, {italic:1, bold:1, value:'8'}, '9', '10']
        ],
        table: true,
        name: 'Sheet 1'
      }]
    });
    fs.writeFile(file, sheet.base64, 'base64', done);
  })

  it('should generated xlsx be readable', function(done) {
    fs.readFile(file, 'base64', function(err, content) {
      if (err) {
        return done(err);
      }
      var sheet = xlsx(content);
      assert.isNotNull(sheet);
      assert.deepEqual(sheet.worksheets, [{
        name: 'Sheet 1',
        data: [
          [{
            value: 'green',
            formatCode: 'General'
          },{
            value: 'white',
            formatCode: 'General'
          },{
            value: 'orange',
            formatCode: 'General'
          },{
            value: 'blue',
            formatCode: 'General'
          },{
            value: 'red',
            formatCode: 'General'
          }],
          [{
            value: 1,
            formatCode: '0.00'
          },{
            value: 2,
            formatCode: 'General'
          },{
            value: 3,
            formatCode: 'General'
          },{
            value: 4,
            formatCode: 'General'
          },{
            value: 5,
            formatCode: 'General'
          }],
          [{
            value: 6,
            formatCode: 'General'
          },{
            value: 7,
            formatCode: 'General'
          },{
            value: 8,
            formatCode: 'General'
          },{
            value: 9,
            formatCode: 'General'
          },{
            value: 10,
            formatCode: 'General'
          }]
        ],
        table: false,
        maxCol: 5,
        maxRow: 3
      }]);
      assert.deepEqual(sheet.creator, 'John Doe');
      assert.deepEqual(sheet.lastModifiedBy, 'Meg White');
      done();
    });
  })

});