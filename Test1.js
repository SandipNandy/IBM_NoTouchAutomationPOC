

//function myCal(Parms){
/*var XLSX = require('xlsx');

//const filePath = 'C:/Users/0038IG744/Downloads/4762_IFRS 15_Mrch 20_Icj_31.3.2020_in000iuw/Audit_WorkBook_4762_02-03-2022/02-03-2022_Audit_WorkBook_EMEA.xlsx';
const filePath = Parms

//const workbookHeaders = XLSX.readFile(filePath, { sheetRows: 1 });
const workbookHeaders = XLSX.readFile(filePath);
let HeaderRow,Column;
let storeV=[];
const sheetList = workbookHeaders.SheetNames;
console.log(sheetList.toString());
let sheetposition = 0;
let cvt = function(n) {return(String.fromCharCode(n+'A'.charCodeAt(0)-1))}

for (let j = 0; j < sheetList.length; j++) {
    // console.log('sheetlist',sheetList[j].includes('Input File-'));
    if (sheetList[j].includes('Input File-')) {
        //console.log('j',j);
        sheetposition = j;
    }
}
var sheet = workbookHeaders.Sheets['' + sheetList[sheetposition] + '']
for (let i = 1; i <= 10000; i++) {
    if (sheet['A' + i] == null) {
        console.log(i);
        HeaderRow = i+1;
        break;
    }
}
for(let j=1;j<=10000;j++){
    let Cell=cvt(j)+HeaderRow;
    if(sheet[Cell] == null){
        Column=j;
        console.log(cvt(j));
        console.log(Column);

        break;

    }
}
 for(let y=1;y<Column;y++){
    let Cell=cvt(y)+(HeaderRow+1);
    if(sheet[Cell].t!=='s'){
      storeV.push(cvt(y))
    console.log(sheet[Cell].v+'=>'+sheet[Cell].t);
    }
 }
 console.log('StoreV',storeV);
//}



// It('dg',function(){
//     myCal(sheet1['AE2'].v)
// })
*/
let TotalCondition=((2*2)+2)+1
console.log('TotalCondition : ',TotalCondition);
let cvt = function (n) { return (String.fromCharCode(n + 'A'.charCodeAt(0) - 1)) }
console.log('CVT : ',cvt(7));
