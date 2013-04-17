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
          ['1', '2', '3', '4', '5'],
          [6, 7, 8, 9, 10]
        ],
        table: true,
        name: 'Sheet 1'
      },{
        data: [
          ['formatting test'],
          [{formatCode: '0.00', value:'1'}, 
           {italic:1, bold:1, hAlign:'center', borders:{bottom:'DEE31D'}, value:'B1'}, 
           {borders:{bottom:64}, value:'C1'},
           {fontName: 'Arial', value:'D1'},
           {fontSize: 8, value:'E1'}, 
           {italic:1, bold:1, value: 'F1'}
          ]
        ],
        name: 'Sheet 2'
      },{
        data: [
          ['merge test'],
          ['A1', {colSpan:3, value:'B1'}, 'E1'],
          [{rowSpan: 3, value:'A2', vAlign: 'center', hAlign: 'center'}, 'B2', 'C2', 'D2', 'E2'],
          ['B3', 'C3', 'D3', 'E3']
        ],
        name: 'Sheet 3'
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
      assert.equal(sheet.worksheets.length, 3);
      assert.deepEqual(sheet.worksheets[0], {
        name: 'Sheet 1',
        data: [
          [
            {value: 'green', formatCode: 'General'},
            {value: 'white', formatCode: 'General'},
            {value: 'orange', formatCode: 'General'},
            {value: 'blue', formatCode: 'General'},
            {value: 'red', formatCode: 'General'}
          ],[ 
            {value: 1, formatCode: 'General'},
            {value: 2, formatCode: 'General'},
            {value: 3, formatCode: 'General'},
            {value: 4, formatCode: 'General'},
            {value: 5, formatCode: 'General'}
          ],[
            {value: 6, formatCode: 'General'},
            {value: 7, formatCode: 'General'},
            {value: 8, formatCode: 'General'},
            {value: 9, formatCode: 'General'},
            {value: 10, formatCode: 'General'}
          ]
        ],
        table: false,
        maxCol: 5,
        maxRow: 3
      });
      assert.deepEqual(sheet.worksheets[1], {
        name: 'Sheet 2',
        data: [
          [{value: 'formatting test', formatCode: 'General'}],
          [
            {value: 1, formatCode: '0.00'},
            {value: 'B1', formatCode: 'General'},
            {value: 'C1', formatCode: 'General'},
            {value: 'D1', formatCode: 'General'},
            {value: 'E1', formatCode: 'General'},
            {value: 'F1', formatCode: 'General'}
          ]
        ],
        table: false,
        maxCol: 1,
        maxRow: 2
      });
      assert.equal(JSON.stringify(sheet.worksheets[2]), JSON.stringify({
        name: 'Sheet 3',
        data: [
          [{value: 'merge test', formatCode: 'General'}],
          [
            {value: 'A1', formatCode: 'General'},
            {value: 'B1', formatCode: 'General'},
            {value: null, formatCode: 'General'},
            {value: null, formatCode: 'General'},
            {value: 'E1', formatCode: 'General'}
          ],[
            {value: 'A2', formatCode: 'General'},
            {value: 'B2', formatCode: 'General'},
            {value: 'C2', formatCode: 'General'},
            {value: 'D2', formatCode: 'General'},
            {value: 'E2', formatCode: 'General'}
          ],[
            {value: null, formatCode: 'General'},
            {value: 'B3', formatCode: 'General'},
            {value: 'C3', formatCode: 'General'},
            {value: 'D3', formatCode: 'General'},
            {value: 'E3', formatCode: 'General'}
          ]
        ],
        table: false,
        maxCol: 1,
        maxRow: 4
      }));
      assert.deepEqual(sheet.creator, 'John Doe');
      assert.deepEqual(sheet.lastModifiedBy, 'Meg White');
      done();
    });
  })

});