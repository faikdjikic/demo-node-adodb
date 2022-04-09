const path = require('path');

opendb('biblio.xlsx','select * from [titles$]').then(result=>console.log(result))

function opendb(databasename, queryString) {
  'use strict';
  let rootpath = path.resolve("./");
  if (rootpath.substr(-1) != path.sep) rootpath += path.sep;
  const ADODB = require('node-adodb');
  const connection = ADODB.open(`Provider=Microsoft.ACE.OLEDB.12.0;Data Source=${rootpath + databasename};Extended Properties="Excel 12.0 Xml;HDR=YES;IMEX=1;"`);
  return new Promise((resolve, reject) => {
    connection.query(queryString)
      .then(data => {
        resolve({
          Error: 0,
          Result: data
        });
      })
      .catch(error => {
        let ErrorResponse = {};
        ErrorResponse.Error = error.process.code;
        ErrorResponse.ErrorMessage = error.process.message;
        resolve(ErrorResponse);
      });
  });
}