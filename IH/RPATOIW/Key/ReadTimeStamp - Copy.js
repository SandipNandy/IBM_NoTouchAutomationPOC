const xlsx = require('xlsx');
const workbook = xlsx.readFile('./ExcelA.xlsx');
const workbook1 = xlsx.readFile('./ExcelB.xlsx')
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
const worksheet1 = workbook1.Sheets[workbook1.SheetNames[1]];
var myDate=[], myTime = [];
var today = new Date();
var todaydate = today.toISOString().split('T')[0]
var time = today.getHours() + ":" + today.getMinutes() + ":" + today.getSeconds();

const columnA = Object.keys(worksheet).filter(x => /^A\d+/.test(x)).map(x => worksheet[x].v);
const columnB = Object.keys(worksheet).filter(x => /^B\d+/.test(x)).map(x => worksheet[x].v);
const columnH = Object.keys(worksheet1).filter(x => /^H\d+/.test(x)).map(x => worksheet1[x].v);
var XLSX = require('xlsx-populate')
XLSX.fromFileAsync('ExcelB.xlsx').then(workbook => {
  let ws=workbook.sheet("Sheet2");
console.log(columnA);
console.log(columnB);
console.log(columnH);
array1 = columnA.filter(val => columnH.includes(val));
console.log('array1 :- ' + array1);
for (let outerloop = 0; outerloop <= array1.length - 1; outerloop++) {

  for (let innerloop = 0; innerloop <= columnA.length - 1; innerloop++) {
    if (array1[outerloop] == columnA[innerloop]) {
      //console.log(columnA.indexOf(columnA[innerloop]));
      //columnB[columnA.indexOf(columnA[innerloop])]
      //console.log(columnB[columnA.indexOf(columnA[innerloop])]);
      myTime.push(columnB[columnA.indexOf(columnA[innerloop])].split(' ')[1])
      myDate.push(columnB[columnA.indexOf(columnA[innerloop])].split(' ')[0])
    }
  }
}
console.log(time);
console.log(myDate);
console.log(myTime);
for(let getindex=0;getindex<=myDate.length-1;getindex++){
      console.log('myDate:- '+myDate.length)
      console.log(columnH.indexOf(columnH[getindex+1]));
      console.log('Country Name - '+columnH[getindex+1]);
      var  haha=columnH.indexOf(columnH[getindex+1])+1;
//if(myTime[getindex]<time){
//console.log('HI I AM HERE');
//}
//if (time == myTime[g] & todaydate == myDate[g]) {
if (myDate[getindex] == todaydate) {
  
   
    ws.cell("F"+haha+"").value("Y");
  console.log("My name is Sandip Nandi")
}
else {
  ws.cell("F"+haha+"").value("N");

 console.log("My name is Sandip Nandi and I am working in IBM under R2R");
}
}

return workbook.toFileAsync('ExcelB.xlsx');
    
});
//var date = new Date();
//var newdate= date.toISOString().split('T')[0]
//console.log(newdate);
/*
************************************
for (let z in worksheet) {
  if(z.toString()[0] === 'A'){
    columnA.push(worksheet[z].v);
  }
}
***********************************
*/

