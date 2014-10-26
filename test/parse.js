'use strict';
// nodemon -w . --exec npm test

var util = require('util'),
    fs = require('fs'),
    async = require('async');

var log = function() {
  var args = Array.prototype.slice.call(arguments, 0);
  return util.log(util.inspect.call(null, args.length === 1 ? args[0] : args, false, null, true));
};

var plist = require('../index');

module.exports.parse = function(assert) {

  var fixture = JSON.parse(fs.readFileSync(__dirname + '/fixtures/test.json'));
  var filename = __dirname + '/fixtures/test.xlsx';
  var xlsObject;

  // parse file
  xlsObject = plist.parse(filename);
  assert.deepEqual(JSON.parse(JSON.stringify(xlsObject)), fixture);

  // parse file using async
  async.series([
      plist.parseSync(filename),
  ], function(err, result){
      assert.deepEqual(JSON.parse(JSON.stringify(result)), fixture);
  });

  // parse buffer
  xlsObject = plist.parse(fs.readFileSync(filename));
  assert.deepEqual(JSON.parse(JSON.stringify(xlsObject)), fixture);

  assert.done();

};
