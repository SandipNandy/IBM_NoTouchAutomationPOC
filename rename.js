var res = new  Date().toString();
var d = res.substr(0,25);
var fs = require('fs');
 
fs.rename('C:/Users/06318O744/Documents/AUTOMATION/CA 14 countries Prod Sanity/CA 14 countries Prod Sanity/Statusstamping for CA.xlsx', 'C:/Users/06318O744/Documents/AUTOMATION/CA 14 countries Prod Sanity/CA 14 countries Prod Sanity/copy/Statusstamping for CA'+d+'.xlsx', function (err) {
  if (err) throw err
  console.log('File Renamed');
});