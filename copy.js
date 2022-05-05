
var res = new  Date().toString();
var d = res.substr(0,21);
var dd= d.replace(':','_');
console.log(dd);
const fs = require('fs');
fs.copyFile('./Statusstamping for CA.xlsx', './copy/Statusstamping for '+dd+'.xlsx', (err) => {
    
    if (err) throw err;
    console.log('File was copied to destination');
  });
