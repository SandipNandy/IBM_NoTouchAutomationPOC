var XLSX = require('xlsx');

// var workbook = XLSX.readFile('./21-09-2021_Audit_WorkBook_Avnet-Americas.xlsx');
// const sheetList = workbook.SheetNames;
// console.log(sheetList.toString());
// let sheetposition = 0;
// for ( let j=0; j<sheetList.length; j++)
// {
//    // console.log('sheetlist',sheetList[j].includes('Input File-'));
//     if(sheetList[j].includes('Input File-'))
//     {
//         //console.log('j',j);
//     sheetposition = j+1;
// }
// }
// console.log('sheetposition',sheetposition);

//function myCal(Parms) {

    const filePath = './02-03-2022_Audit_WorkBook_EMEA1.xlsx';
   //const filePath = Parms

   //const workbookHeaders = XLSX.readFile(filePath, { sheetRows: 1 });
   const workbookHeaders = XLSX.readFile(filePath);
   let HeaderRow, Column;
   let storeV = [];
   const sheetList = workbookHeaders.SheetNames;
   console.log(sheetList.toString());
   let sheetposition = 0;
   let RegExdate=/^\d{1,2}([./-])\d{1,2}([./-])\d{2,4}$/;
   let cvt = function (n) { return (String.fromCharCode(n + 'A'.charCodeAt(0) - 1)) }

   for (let j = 0; j < sheetList.length; j++) {
       // console.log('sheetlist',sheetList[j].includes('Input File-'));
       if (sheetList[j].includes('Input File-')) {
           //console.log('j',j);
           sheetposition = j;
       }
   }
   var sheet = workbookHeaders.Sheets['' + sheetList[sheetposition] + '']
   for (let i = 1; i <= 10000; i++) {
       if (sheet['A' + i] !== null && (sheet['B' + i] == null || sheet['C' + i] == null)) {

           continue;
       }
       if (sheet['A' + i] == null || sheet['B' + i] == null || sheet['C' + i] == null) {

           continue;
       }
       if (sheet['A' + i] !== null || sheet['B' + i] !== null || sheet['C' + i] !== null) {
           console.log('i', i);
           //HeaderRow = i - 1;
           HeaderRow = i ;

           break;
       }
   }
   for (let j = HeaderRow; j <= 10000; j++) {
       let Cell = cvt(j) + HeaderRow;
       console.log(Cell);
       if (sheet[Cell] == null) {
           Column = j;
           console.log(cvt(j));
           console.log(Column);

           break;

       }
   }
   for (let y = 1; y < Column; y++) {
       let Cell = cvt(y) + (HeaderRow + 1);
       //let Cell = cvt(y) + HeaderRow;
 
       console.log('69 : ',Cell);
       console.log('76 : ',sheet[Cell].v);
       if (sheet[Cell].t !== 's' ) {
           storeV.push(cvt(y))
           console.log(sheet[Cell].v + '=>' + sheet[Cell].t);
       }
   }
   console.log('StoreV', storeV);
   //return storeV;

//}


// let bb='A,G,I,L';
// let a=['B','C','D','E','F','G','H','I','J','K','L']
// let b=bb.split(',');
// console.log(a.filter(e => b.includes(e)));
// for(let i=0;i<=a.length;i++){
//     console.log(b.includes(a[i]));
// }
// let cleaning=0;
// let cleaningstepsArray=[];
// let cleaningFilesArray=[];

// while(cleaning<ar.length){
//     cleaningFilesArray.push(Number(ar[cleaning].split('-')[0].split(/[Ff]/)[1]))
//     cleaningstepsArray.push(Number(ar[cleaning].split('-')[1]));
//     cleaning++
// }
// console.log('cleaningFilesArray: ', cleaningFilesArray)
// console.log(cleaningstepsArray)
// let ar=r.forEach(element => {
//  element.split('-')[1]
    
// });

//console.log(ar);

/*let y='A,L';
console.log(y.split(','));
//let yy=
y.split(',').map(el=>{console.log(el)});
//let yyy=
// yy.map(el=>{
//     console.log(el.split('->'))
// });
//console.log(yy);
*/








/*8888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888*/
if (ProcessingStepsAddColumnMainSheet['C' + AddColumnVariable].v === 'Date') {
    element(by.xpath("//div[contains(text(),'" + ProcessingStepsAddColumnDate['C' + AddColumnDateRow].v + "')]/preceding-sibling::button/span/mat-icon")).click();
    browser.sleep(1000);
    element(by.xpath("//*[@name='selectedStepType']")).click();
    browser.sleep(1000);
    //element(by.xpath("//mat-option/span[contains(text(),'"+ProcessingStepsAddColumnDate['D3'].v+"')]")).click();
    let DFTY = Jfun.DateFormulaType(ProcessingStepsAddColumnDate['D' + AddColumnDateRow].v);
    browser.executeScript("arguments[0].click()", DFTY);
    browser.sleep(1000);
    element(by.xpath("//input[@aria-label='Click and Select Column from Grid' and @type='search']")).click();
    browser.sleep(1000);
    element(by.xpath("//input[@aria-label='Value']")).sendKeys(ProcessingStepsAddColumnDate['F' + AddColumnDateRow].v);
    browser.sleep(4000);
    AddColumnDateRow++;

}
if (ProcessingStepsAddColumnMainSheet['C' + AddColumnVariable].v === 'Math') {
    element(by.xpath("//div[contains(text(),'" + ProcessingStepsAddColumnMath['C' + AddColumnMath].v + "')]/preceding-sibling::button/span/mat-icon")).click();
    browser.sleep(1000);
    element(by.xpath("//div[@class='mat-select-value']//span/span[contains(text(),'Arithmetic')]")).click();
    browser.sleep(1000);
    element(by.xpath("//mat-option/span[contains(text(),'" + ProcessingStepsAddColumnMath['D' + AddColumnMath].v + "')]")).click();
    browser.sleep(1000);
    Jfun.SelectMathOperations(ProcessingStepsAddColumnMath['J' + AddColumnMath].v);
    browser.sleep(1000);
    // if(ProcessingStepsAddColumnMath[cvt(7)+].v)
    for (let itrvalue = 1; itrvalue <= 2; itrvalue++) {
        if (ProcessingStepsAddColumnMath['G' + AddColumnMath].v === 'YES' || ProcessingStepsAddColumnMath['G' + AddColumnMath].v === 'yes') {
            element(by.xpath("(//input[@aria-label='Click and Select Column from Grid' or @aria-label='Value'])[" + itrvalue + "]")).click();
            browser.sleep(1000);
            if (itrvalue === 1) {
                element(by.xpath("//div[@class='ag-grid-custom-header-bar']//div[contains(@class,'ag-grid-custom-header')][contains(text(),'" + ProcessingStepsAddColumnMath['H' + AddColumnMath].v + "')]")).click();
            }
            if (itrvalue === 2) {
                element(by.xpath("//div[@class='ag-grid-custom-header-bar']//div[contains(@class,'ag-grid-custom-header')][contains(text(),'" + ProcessingStepsAddColumnMath['K' + AddColumnMath].v + "')]")).click();

            }
        }
        if (ProcessingStepsAddColumnMath['F' + AddColumnMath].v !== 'YES' || ProcessingStepsAddColumnMath['F' + AddColumnMath].v !== 'yes') {
            let value = element(by.xpath("(//div[contains(text(),'Value')]/preceding::mat-icon[contains(@aria-label,'checkbox')])[" + itrvalue + "]"))//.click();
            browser.sleep(1000);
            browser.executeScript("arguments[0].scrollIntoView()", value);
            browser.executeScript("arguments[0].click()", value);

            if (itrvalue === 1 && ProcessingStepsAddColumnMath['I' + AddColumnMath].v === 'NA') {
                element(by.xpath("(//input[@aria-label='Click and Select Column from Grid' or @aria-label='Value'])[" + itrvalue + "]")).sendKeys(ProcessingStepsAddColumnMath['I3'].v);
            }
            if (itrvalue === 2 && ProcessingStepsAddColumnMath['L' + AddColumnMath].v === 'NA') {
                element(by.xpath("(//input[@aria-label='Click and Select Column from Grid' or @aria-label='Value'])[" + itrvalue + "]")).sendKeys(ProcessingStepsAddColumnMath['L3'].v);
            }
        }

    }
    if (ProcessingStepsAddColumnMath['F' + AddColumnMath].v === 'YES' || ProcessingStepsAddColumnMath['F' + AddColumnMath].v === 'Yes') {
        element(by.xpath("//span[contains(text(),'Make Absolute')]/parent::div/preceding-sibling::button//mat-icon[contains(@aria-label,'checkbox')]")).click();
    }
    element(by.xpath("//mat-select[@name='selectedRoundValue']")).click();
    browser.sleep(1000);
    element(by.xpath("//mat-option/span[normalize-space(text())='" + ProcessingStepsAddColumnMath['E' + AddColumnMath].v + "']")).click();
    browser.sleep(1000);
    if (ProcessingStepsAddColumnMath['E' + AddColumnMath].v !== 'None') {
        // browser.sleep(1000);
        //let scrolltoRDP=element(by.xpath("(//mat-label[contains(text(),'Round Decimal Places')])[2]"));
        //browser.executeScript("arguments[0].scrollIntoView()", scrolltoRDP);
        browser.sleep(1000);
        let decimalPlaces = element(by.xpath("//span[contains(text(),'-1')]"))//.click();
        browser.executeScript("arguments[0].scrollIntoView()", decimalPlaces);
        browser.executeScript("arguments[0].click()", decimalPlaces);
        browser.sleep(1000);
        element(by.xpath("//mat-option/span[normalize-space(text())='" + ProcessingStepsAddColumnMath['M' + AddColumnMath].v + "']")).click();
        browser.sleep(1000);
    }


}
if (ProcessingStepsAddColumnMainSheet['C' + AddColumnVariable].v === 'Text') {
    element(by.xpath("//div[contains(text(),'" + ProcessingStepsAddColumnText['C' + AddColumnText].v + "')]/preceding-sibling::button/span/mat-icon")).click();
    browser.sleep(1000);
    element(by.xpath("//mat-select[@aria-label='Text Formula Type']")).click();
    browser.sleep(1000);
    element(by.xpath("//mat-option/span[contains(text(),'" + ProcessingStepsAddColumnText['D' + AddColumnText].v + "')]")).click();
    if (ProcessingStepsAddColumnText['D' + AddColumnText].v === 'CONCAT') {
        if (ProcessingStepsAddColumnText['F' + AddColumnText].v !== 'YES') {
            let value1 = element(by.xpath("(//div[contains(text(),'Value')]/preceding::mat-icon[contains(@aria-label,'checkbox')])"))//.click();
            browser.executeScript("arguments[0].scrollIntoView()", value1);
            browser.executeScript("arguments[0].click()", value1);
            browser.sleep(1000);
            element(by.xpath("(//input[@aria-label='Click and Select Column from Grid' or @aria-label='Value'])")).sendKeys(ProcessingStepsAddColumnText['G' + AddColumnText].v);
        }
        else {
            let value1 = element(by.xpath("(//input[@aria-label='Click and Select Column from Grid' or @aria-label='Value'])"));
            browser.executeScript("arguments[0].scrollIntoView()", value1);
            browser.executeScript("arguments[0].click()", value1);
            //.click();
            browser.sleep(1000);
            element(by.xpath("//div[@class='ag-grid-custom-header-bar']//div[contains(@class,'ag-grid-custom-header')][contains(text(),'" + ProcessingStepsAddColumnText['E' + AddColumnText].v + "')]")).click();
            browser.sleep(10000);

        }
    }
    if ((ProcessingStepsAddColumnText['D' + AddColumnText].v === 'LTRIM' || ProcessingStepsAddColumnText['D' + AddColumnText].v === 'RTRIM') || (ProcessingStepsAddColumnText['D' + AddColumnText].v === 'TRIM' || ProcessingStepsAddColumnText['D' + AddColumnText].v === 'LEN')) {
        element(by.xpath("(//input[@aria-label='Click and Select Column from Grid' or @aria-label='Value'])")).click();
        browser.sleep(1000);
        element(by.xpath("//div[@class='ag-grid-custom-header-bar']//div[contains(@class,'ag-grid-custom-header')][contains(text(),'" + ProcessingStepsAddColumnText['E' + AddColumnText].v + "')]")).click();

    }
    if (ProcessingStepsAddColumnText['D' + AddColumnText].v === 'LEFT' || ProcessingStepsAddColumnText['D' + AddColumnText].v === 'RIGHT') {
        element(by.xpath("//input[@aria-label='Click and Select Column from Grid']")).click();
        browser.sleep(1000);
        element(by.xpath("//div[@class='ag-grid-custom-header-bar']//div[contains(@class,'ag-grid-custom-header')][contains(text(),'" + ProcessingStepsAddColumnText['E' + AddColumnText].v + "')]")).click();
        browser.sleep(1000);
        element(by.xpath("//div[contains(text(),'" + ProcessingStepsAddColumnText['I' + AddColumnText].v + "')]/preceding-sibling::button/span")).click();
        browser.sleep(1000);
        element(by.xpath("//input[@aria-label='Value']")).clear();
        browser.sleep(1000);
        element(by.xpath("//input[@aria-label='Value']")).sendKeys(ProcessingStepsAddColumnText['J' + AddColumnText].v);

    }
    if (ProcessingStepsAddColumnText['D' + AddColumnText].v === 'REPLACE') {
        element(by.xpath("//input[@aria-label='Click and Select Column from Grid']")).click();
        browser.sleep(1000);
        element(by.xpath("//div[@class='ag-grid-custom-header-bar']//div[contains(@class,'ag-grid-custom-header')][contains(text(),'" + ProcessingStepsAddColumnText['E' + AddColumnText].v + "')]")).click();
        browser.sleep(1000);
        element(by.xpath("(//input[@aria-label='Value'])[1]")).sendKeys(ProcessingStepsAddColumnText['K' + AddColumnText].v);
        browser.sleep(1000);
        element(by.xpath("(//input[@aria-label='Value'])[2]")).sendKeys(ProcessingStepsAddColumnText['L' + AddColumnText].v);

    }
    if (ProcessingStepsAddColumnText['D' + AddColumnText].v === 'MID') {
        element(by.xpath("//input[@aria-label='Click and Select Column from Grid']")).click();
        browser.sleep(1000);
        element(by.xpath("//div[@class='ag-grid-custom-header-bar']//div[contains(@class,'ag-grid-custom-header')][contains(text(),'" + ProcessingStepsAddColumnText['E' + AddColumnText].v + "')]")).click();
        browser.sleep(1000);
        //start index
        element(by.xpath("(//input[@aria-label='Value'])[1]")).sendKeys(ProcessingStepsAddColumnText['M' + AddColumnText].v);
        browser.sleep(1000);
        //take number of char
        element(by.xpath("(//input[@aria-label='Value'])[2]")).sendKeys(ProcessingStepsAddColumnText['N' + AddColumnText].v);
    }
}
if (ProcessingStepsAddColumnMainSheet['C' + AddColumnVariable].v === 'Fixed') {
    element(by.xpath("//div[contains(text(),'" + ProcessingStepsAddColumnFixed['C' + AddColumnFixed].v + "')]/preceding-sibling::button/span/mat-icon")).click();
    browser.sleep(1000);
    element(by.xpath("(//mat-select[@aria-label='Fixed Value Formula Type'])")).click();
    browser.sleep(1000);
    element(by.xpath("//mat-option/span[contains(text(),'" + ProcessingStepsAddColumnFixed['D' + AddColumnFixed].v + "')]")).click();
    browser.sleep(1000);
    if (ProcessingStepsAddColumnFixed['D' + AddColumnFixed].v === 'Column') {
        element(by.xpath("//input[@aria-label='Click and Select Column from Grid']")).click();
        browser.sleep(1000);
        element(by.xpath("//div[@class='ag-grid-custom-header-bar']//div[contains(@class,'ag-grid-custom-header')][contains(text(),'" + ProcessingStepsAddColumnText['E' + AddColumnFixed].v + "')]")).click();
        browser.sleep(1000);
    }
    if (ProcessingStepsAddColumnFixed['D' + AddColumnFixed].v === 'Static Value1') {

        if (ProcessingStepsAddColumnFixed['F' + AddColumnFixed].v === 'TEXT') {
            element(by.xpath("//mat-select[@aria-label='Data Type']")).click();
            browser.sleep(1000);
            element(by.xpath("//mat-option/span[contains(text(),'" + ProcessingStepsAddColumnFixed['F' + AddColumnFixed].v + "')]")).click();
            browser.sleep(1000);
            element(by.xpath("//input[@aria-label='Value']")).sendKeys(ProcessingStepsAddColumnFixed['G' + AddColumnFixed].v);
            browser.sleep(1000);
        }
        if (ProcessingStepsAddColumnFixed['F' + AddColumnFixed].v === 'NUMBER1') {
            element(by.xpath("//mat-select[@aria-label='Data Type']")).click();
            browser.sleep(1000);
            element(by.xpath("//mat-option/span[contains(text(),'" + ProcessingStepsAddColumnFixed['F' + AddColumnFixed].v + "')]")).click();
            browser.sleep(1000);
            element(by.xpath("//input[@aria-label='Value']")).sendKeys(ProcessingStepsAddColumnFixed['G' + AddColumnFixed].v);
            browser.sleep(1000);
        }
        if (ProcessingStepsAddColumnFixed['F' + AddColumnFixed].v === 'DECIMAL') {
            element(by.xpath("//mat-select[@aria-label='Data Type']")).click();
            browser.sleep(1000);
            element(by.xpath("//mat-option/span[contains(text(),'" + ProcessingStepsAddColumnFixed['F' + AddColumnFixed].v + "')]")).click();
            browser.sleep(1000);
            element(by.xpath("//input[@aria-label='Value']")).sendKeys(ProcessingStepsAddColumnFixed['G' + AddColumnFixed].v);
            browser.sleep(1000);

        }
        if (ProcessingStepsAddColumnFixed['F' + AddColumnFixed].v === 'DATE') {
            element(by.xpath("//mat-select[@aria-label='Data Type']")).click();
            browser.sleep(1000);
            element(by.xpath("//mat-option/span[contains(text(),'" + ProcessingStepsAddColumnFixed['F' + AddColumnFixed].v + "')]")).click();
            browser.sleep(1000);
            element(by.xpath("//input[@aria-label='Value']")).sendKeys(ProcessingStepsAddColumnFixed['G' + AddColumnFixed].v);
            browser.sleep(1000);

        }
    }
    if (ProcessingStepsAddColumnFixed['D' + AddColumnFixed].v === 'Aggregate1') {

    }
    if (ProcessingStepsAddColumnFixed['D' + AddColumnFixed].v === 'Keyword1') {

    }
}
if (ProcessingStepsAddColumnMainSheet['C' + AddColumnVariable].v === 'Update Data Type') {
    element(by.xpath("//div[contains(text(),'" + ProcessingStepsAddColumnUpdateDataType['C' + AddColumnUpdateDataType].v + "')]/preceding-sibling::button/span/mat-icon")).click();
    browser.sleep(1000);
    element(by.xpath("//input[@aria-label='Click and Select Column from Grid']")).click();
    browser.sleep(1000);
    element(by.xpath("//div[@class='ag-grid-custom-header-bar']//div[contains(@class,'ag-grid-custom-header')][contains(text(),'" + ProcessingStepsAddColumnUpdateDataType['D' + AddColumnUpdateDataType].v + "')]")).click();
    browser.sleep(1000);
    element(by.xpath("//mat-select[@aria-label='Data Type']")).click();
    browser.sleep(1000);
    element(by.xpath("//mat-option/span[contains(text(),'" + ProcessingStepsAddColumnUpdateDataType['E' + AddColumnUpdateDataType].v + "')]")).click();
    if (ProcessingStepsAddColumnUpdateDataType['E' + AddColumnUpdateDataType].v === 'DECIMAL') {
        element(by.xpath("//mat-select[@name='selectedRoundValue']")).click();
        browser.sleep(1000);
        element(by.xpath("//mat-option/span[contains(text(),'" + ProcessingStepsAddColumnUpdateDataType['F' + AddColumnUpdateDataType].v + "')]")).click();
        if (ProcessingStepsAddColumnUpdateDataType['F' + AddColumnUpdateDataType].v === 'None') {
            if (ProcessingStepsAddColumnUpdateDataType['G' + AddColumnUpdateDataType].v === 'YES') {
                //Mark Absolute
                element(by.xpath("//app-checkbox//button[@aria-label='check_box_outline_blank checkbox']/span/mat-icon")).click();
            }
        }
        if ((ProcessingStepsAddColumnUpdateDataType['F' + AddColumnUpdateDataType].v === 'Round' || ProcessingStepsAddColumnUpdateDataType['F' + AddColumnUpdateDataType].v === 'Round up') || ProcessingStepsAddColumnUpdateDataType['F' + AddColumnUpdateDataType].v === 'Round down') {
            if (ProcessingStepsAddColumnUpdateDataType['G' + AddColumnUpdateDataType].v === 'YES') {
                //Mark Absolute
                let MarkAbsolute = element(by.xpath("//app-checkbox//button[@aria-label='check_box_outline_blank checkbox']/span/mat-icon"));//.click();
                browser.executeScript("arguments[0].scrollIntoView()", MarkAbsolute);
                browser.executeScript("arguments[0].click()", MarkAbsolute);
                browser.sleep(3000);
            }
            element(by.xpath("//mat-select[@name='decimalPlaces']")).click();
            browser.sleep(1000);
            element(by.xpath("//mat-option/span[contains(text(),'" + ProcessingStepsAddColumnUpdateDataType['H' + AddColumnUpdateDataType].v + "')]")).click();
            browser.sleep(2000).then(() => {
                let decimalPlaces = element(by.xpath("//mat-select[@name='decimalPlaces']//preceding::span[contains(text(),'-1')]"));//.click();
                browser.executeScript("arguments[0].scrollIntoView()", decimalPlaces);
                browser.executeScript("arguments[0].click()", decimalPlaces);
            });
            browser.sleep(4000);
            element(by.xpath("//mat-option/span[normalize-space(text())='" + ProcessingStepsAddColumnUpdateDataType['H' + AddColumnUpdateDataType].v + "']")).click();


        }
        browser.sleep(10000);


    }
    if (ProcessingStepsAddColumnUpdateDataType['E' + AddColumnUpdateDataType].v === 'DATE') {
        element(by.xpath("//input[@placeholder='Enter Date Format']")).sendKeys('02-05-2022');
    }

}
if (ProcessingStepsAddColumnMainSheet['C' + AddColumnVariable].v === 'Row Number') {
    let MYARR = ProcessingStepsAddColumnRowNumber['E' + AddColumnRowNumber].v.split(',');
    console.log('MYARR : ', MYARR);
    let ASCorDESC = 6, ROWnum = 3;
    element(by.xpath("//div[contains(text(),'" + ProcessingStepsAddColumnRowNumber['C' + AddColumnRowNumber].v + "')]/preceding-sibling::button/span/mat-icon")).click();
    browser.sleep(1000);

    for (let itr4 = 1; itr4 <= ProcessingStepsAddColumnRowNumber['D' + AddColumnRowNumber].v; itr4++) {
        element(by.xpath("//div[@class='ag-grid-custom-header-bar']//div[contains(@class,'ag-grid-custom-header')][contains(text(),'" + MYARR[itr4 - 1] + "')]")).click();
        browser.sleep(1000);
        let scrolltoSortBy = element(by.xpath("(//mat-select[@aria-label='Sort By'])[" + itr4 + "]"))//.click();
        browser.executeScript("arguments[0].scrollIntoView()", scrolltoSortBy);
        browser.executeScript("arguments[0].click()", scrolltoSortBy);
        browser.sleep(1000);
        console.log('cvt(ASCorDESC)+ROWnum : ', cvt(ASCorDESC) + ROWnum);
        let rownumber = element(by.xpath("(//mat-option/span[contains(text(),'" + ProcessingStepsAddColumnRowNumber[cvt(ASCorDESC) + ROWnum].v + "')])"))
        browser.executeScript("arguments[0].scrollIntoView()", rownumber);
        browser.executeScript("arguments[0].click()", rownumber);
        //.click();
        browser.sleep(1000);
        ASCorDESC++;
        //ROWnum++
    }

}
/***88888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888 */