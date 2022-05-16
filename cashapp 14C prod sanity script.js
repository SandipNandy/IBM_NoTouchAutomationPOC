var XLSX = require('xlsx');
var workbook = XLSX.readFile('./cashapp 14C PROD SANITY excel.xlsx');
var statusworkbook = XLSX.readFile('./Statusstamping for CA.xlsx');
var columnselection = require('./ColmnSelection.js');
var res = new Date().toString();
var d = res.substr(0, 25);

var statusstamping = statusworkbook.Sheets['Sheet1'];
//console.log("Status Status stamping" + statusstamping);
var WorksheetLogin = workbook.Sheets['Sheet1'];
var WorksheetDios = workbook.Sheets[browser.params.env.name];
var controlstatments = workbook.Sheets['Sheet3'];
var FRow = controlstatments['A2'].v;
var LRow = controlstatments['B2'].v;
var cashappspom = require('./Cashapps_POM1.js');
var EC = protractor.ExpectedConditions;
var res = new Date().toString();
var rectypesarray = [];
var statusarray = [];


function FiletransferToFolder() {
    const fs = require('fs')
    const dir = 'IH/RPA2IW/IH/Input'

    //const dir = 'C:/Users/06318O744/Documents/ActionLog/SanityAllModulesGit/CashApps-Automation-Oriflame/ibm-r2r-automation/IH/RPA2IW/IH/Input'

    var h = [], zipName, ZipNamePrefix = [], zipFileName = [];
    var files = fs.readdirSync(dir)
    var stringfile = files.toString();
    var newfilename = stringfile.replace(/[,]/g, ' ');
    var words = newfilename.split(' ');

    for (let ii = 0; ii < files.length; ii++) {
        if (words[ii] != 'BOSNIA' && words[ii] != 'CROATIA' && words[ii] != 'ESTONIA' && words[ii] != 'ROMANIA' && words[ii] != 'FINLAND' && words[ii] != 'LATVIA' && words[ii] != 'LITHUANIA' && words[ii] != 'NETHERLANDS' && words[ii] != 'SERBIA' && words[ii] != 'SLOVAKIA' && words[ii] != 'SLOVENIA') {
            //console.log("words zip :" + words[ii]);
            zipFileName.push(words[ii]);
            var newZIPname = words[ii].replace(/[_]/g, ' ');
            zipName = newZIPname.split(' ');
            ZipNamePrefix.push(zipName[0].toUpperCase())
            //console.log("38 SPLITTED ZIP :" + ZipNamePrefix);
        }
        else {
            h.push(ii);
        }
    }
    console.log("SPLITTED ZIP :" + ZipNamePrefix);
    console.log(h);
    console.log("zip FileName :" + zipFileName);
    for (let Folder = 0; Folder < h.length; Folder++) {
        // console.log('My Name is Nandi');
        var myvalue = h[Folder];
        for (let ZipFile = 0; ZipFile < zipFileName.length; ZipFile++) {
            if (words[myvalue] == ZipNamePrefix[ZipFile]) {
                console.log(':' + ZipNamePrefix[ZipFile]);
                console.log(':-' + words[myvalue]);
                var moveFile = (file, dir2) => {
                    var fs = require('fs');
                    var path = require('path');

                    var f = path.basename(file);
                    var dest = path.resolve(dir2, f);
                    fs.rename(file, dest, (err) => {
                        if (err) throw err;
                        else console.log('Successfully moved');
                    });
                };
                moveFile('' + __dirname + '/IH/RPA2IW/IH/Input/' + zipFileName[ZipFile] + '', '' + __dirname + '/IH/RPA2IW/IH/Input/' + words[myvalue] + '');
            }

            else if (words[myvalue] == 'SERBIA' && ZipNamePrefix[ZipFile] == 'CASH') {
                var moveFile = (file, dir2) => {
                    var fs = require('fs');
                    var path = require('path');

                    var f = path.basename(file);
                    var dest = path.resolve(dir2, f);
                    fs.rename(file, dest, (err) => {
                        if (err) throw err;
                        else console.log('Successfully moved');
                    });
                };
                //moveFile('IH/RPA2IW/IH/Input/Cash_Apps_Summary-Summary_Panel-RHS.zip' , 'IH/RPA2IW/IH/Input/SERBIA');
                moveFile('' + __dirname + '/IH/RPA2IW/IH/Input/Cash_Apps_Summary-Summary_Panel-RHS.zip', '' + __dirname + '/IH/RPA2IW/IH/Input/SERBIA');

            }
        }
    }
}

function statusvalidation(yy, rrn) {
    //Selecting the data package
    console.log('69 line: ' + rrn);
    var DPname = element(by.xpath("//span[contains(text(),'" + WorksheetDios['A' + yy].v + "')]"));

    browser.sleep(2000);
    DPname.getText().then((DPN) => {
        console.log('DPN' + DPN);
        console.log(statusstamping['A' + yy].v);
        statusstamping['A' + rrn].v = DPN;
        XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
        var openinvoicereport1 = element.all(by.xpath("//div[@col-id='fileName' and @role='gridcell']"));
        //var openinvoicereport = element(by.xpath("//div[contains(text(),'Openinvoicereport')][@col-id='fileName']/following::div[@ref='eCenterContainer' and @role='rowgroup']/div[@role='row'][1]/div[@col-id='status']/span/span[contains(text(),'SUCCESS')]"));
        openinvoicereport1.count().then((OIVR) => {
            openinvoicereport1.getText().then((fileName) => {
                var array = [];
                for (let I = 0; I < OIVR; I++) {
                    console.log(fileName[I]);
                    var A = fileName[I];
                    var B = A.split('_');
                    array.push(B[0]);
                }
                console.log('array :' + array);
                var position1 = array.indexOf("Unmatched");
                var Fposition = position1 + 1;
                console.log('position1 :' + position1);
                console.log('Fposition :' + Fposition);
                var position2 = array.indexOf("Openinvoicereport");
                var Sposition = position2 + 1;
                console.log('position2 : ' + position2);
                console.log('Sposition : ' + Sposition);

                if (OIVR == 2) {
                    //var orgsuccesscount = element(by.xpath("//div[contains(text(),'SUCCESS')]/following::div/strong[contains(text(),'2')]"));
                    var orgsuccesscount = element(by.xpath("//div[contains(text(),'SUCCESS')]/following-sibling::div/strong"));
                    browser.wait(EC.visibilityOf(orgsuccesscount), 300000).then(() => {
                        var openinvoicereport = element(by.xpath("//div[@col-id='fileName']/following::div[@ref='eCenterContainer' and @role='rowgroup']/div[@role='row'][" + Fposition + "]/div[@col-id='status']/span[1]"));
                        filestatus(openinvoicereport, Sposition);
                    })
                }
                else if (OIVR == 3) {
                    //var orgsuccesscount1 = element(by.xpath("//div[contains(text(),'SUCCESS')]/following::div/strong[contains(text(),'2')]"));
                    var orgsuccesscount1 = element(by.xpath("//div[contains(text(),'SUCCESS')]/following-sibling::div/strong"));
                    browser.wait(EC.visibilityOf(orgsuccesscount1), 300000).then(() => {
                        var openinvoicereport = element(by.xpath("//div[@col-id='fileName']/following::div[@ref='eCenterContainer' and @role='rowgroup']/div[@role='row'][" + Fposition + "]/div[@col-id='status']/span[1]"));
                        filestatus(openinvoicereport, Sposition);
                    })
                }

                else if (OIVR == 4) {
                    //var orgsuccesscount2 = element(by.xpath("//div[contains(text(),'SUCCESS')]/following::div/strong[contains(text(),'2')]"));
                    var orgsuccesscount2 = element(by.xpath("//div[contains(text(),'SUCCESS')]/following-sibling::div/strong"));
                    browser.wait(EC.visibilityOf(orgsuccesscount2), 300000).then(() => {
                        var openinvoicereport = element(by.xpath("//div[@col-id='fileName']/following::div[@ref='eCenterContainer' and @role='rowgroup']/div[@role='row'][" + Fposition + "]/div[@col-id='status']/span[1]"));
                        filestatus(openinvoicereport, Sposition);
                    })
                }
            })
        })

        function filestatus(Fun_openinvoicereport, row) {
            var unmatched = element(by.xpath("//div[contains(text(),'Unmatched')][@col-id='fileName']/following::div[@ref='eCenterContainer' and @role='rowgroup']/div[@role='row'][" + row + "]/div[@col-id='status']/span[1]"));
            //var unmatched = element(by.xpath("//div[contains(text(),'Unmatched')][@col-id='fileName']/following::div[@ref='eCenterContainer' and @role='rowgroup']/div[@role='row'][1]/div[@col-id='status']/span/span[contains(text(),'SUCCESS')]"));
            Fun_openinvoicereport.getText().then((OIR) => {
                unmatched.getText().then((UM) => {
                    if (OIR == 'SUCCESS' && UM == 'SUCCESS') {
                        statusstamping['B' + rrn].v = 'Success';
                        XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');

                        //Clicking on Org unit grids
                        cashappspom.orgunit_grids();
                        browser.sleep(4000);
                    }
                    else if (OIR == 'FAILURE' || UM == 'FAILURE') {
                        statusstamping['B' + rrn].v = 'Failure';
                        XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                        //Clicking on Org unit grids
                        cashappspom.orgunit_grids();
                        browser.sleep(4000);

                    }
                    else if (OIR == 'SKIPPED' || UM == 'SKIPPED') {
                        statusstamping['B' + rrn].v = 'Skipped';
                        XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                        //Clicking on Org unit grids
                        cashappspom.orgunit_grids();
                        browser.sleep(4000);

                    }
                    else if (OIR == 'EXCEPTION' || UM == 'EXCEPTION') {
                        statusstamping['B' + rrn].v = 'Exception';
                        XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                        //Clicking on Org unit grids
                        cashappspom.orgunit_grids();
                        browser.sleep(4000);
                    }
                })
            })
        }
    })
}

function orgunitstatusvalidation(xx, rrn) {
    //Clicking on Recording Period in orgunit grids
    cashappspom.orgunit_recordingperiod();
    browser.sleep(2000);

    //Selecting the Recording period
    cashappspom.org_recordingperiod();
    browser.sleep(5000);

    //Selecting the data package
    var DPname1 = element(by.xpath("//span[contains(text(),'" + WorksheetDios['G' + xx].v + "')]"));
    browser.sleep(2000);
    DPname1.getText().then((DPN) => {
        console.log('DPN' + DPN);
        //console.log(statusstamping['G' + xx].v);
        //statusstamping['G' + xx].v = DPN;
        //XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
        var openinvoicereport1 = element.all(by.xpath("//div[@col-id='fileName' and @role='gridcell']"));
        //var openinvoicereport = element(by.xpath("//div[contains(text(),'Openinvoicereport')][@col-id='fileName']/following::div[@ref='eCenterContainer' and @role='rowgroup']/div[@role='row'][2]/div[@col-id='status']/span[1]"));
        openinvoicereport1.count().then((OIVR) => {
            openinvoicereport1.getText().then((fileName) => {
                var array = [];
                for (let I = 0; I < OIVR; I++) {
                    console.log(fileName[I]);
                    var A = fileName[I];
                    var B = A.split('_');
                    array.push(B[0]);
                }
                console.log('array 187:' + array);
                var position1 = array.indexOf("Unmatched");
                var Fposition = position1 + 1;
                console.log('position1 :' + position1);
                console.log('Fposition :' + Fposition);

                var position2 = array.indexOf("Openinvoicereport");
                var Sposition = position2 + 1;
                console.log('position2 : ' + position2);
                console.log('Sposition : ' + Sposition);

                successvalidation();

                if (OIVR == 2) {
                    //var orgsuccesscount = element(by.xpath("//div[contains(text(),'SUCCESS')]/following::div/strong[contains(text(),'2')]"));
                    var orgsuccesscount = element(by.xpath("//div[contains(text(),'SUCCESS')]/following-sibling::div/strong[1]"));
                    browser.wait(EC.visibilityOf(orgsuccesscount), 300000).then(() => {
                        var openinvoicereport = element(by.xpath("//div[@col-id='fileName']/following::div[@ref='eCenterContainer' and @role='rowgroup']/div[@role='row'][" + Fposition + "]/div[@col-id='status']/span[1]"));
                        filestatus(openinvoicereport, Sposition);
                    })
                }
                else if (OIVR == 3) {
                    var orgsuccesscount1 = element(by.xpath("//div[contains(text(),'SUCCESS')]/following::div/strong[1]"));
                    browser.wait(EC.visibilityOf(orgsuccesscount1), 300000).then(() => {
                        var openinvoicereport = element(by.xpath("//div[@col-id='fileName']/following::div[@ref='eCenterContainer' and @role='rowgroup']/div[@role='row'][" + Fposition + "]/div[@col-id='status']/span[1]"));
                        filestatus(openinvoicereport, Sposition);
                    })
                }

                else if (OIVR == 4) {
                    var orgsuccesscount2 = element(by.xpath("//div[contains(text(),'SUCCESS')]/following::div/strong[1]"));
                    browser.wait(EC.visibilityOf(orgsuccesscount2), 300000).then(() => {
                        var openinvoicereport = element(by.xpath("//div[@col-id='fileName']/following::div[@ref='eCenterContainer' and @role='rowgroup']/div[@role='row'][" + Fposition + "]/div[@col-id='status']/span[1]"));
                        filestatus(openinvoicereport, Sposition);
                    })
                }
            })
        })
        function filestatus(Fun_openinvoicereport, row) {
            var unmatched = element(by.xpath("//div[contains(text(),'Unmatched')][@col-id='fileName']/following::div[@ref='eCenterContainer' and @role='rowgroup']/div[@role='row'][" + row + "]/div[@col-id='status']/span[1]"));
            Fun_openinvoicereport.getText().then((OIR1) => {
                //var unmatched = element(by.xpath("//div[contains(text(),'Unmatched')][@col-id='fileName']/following::div[@ref='eCenterContainer' and @role='rowgroup']/div[@role='row'][1]/div[@col-id='status']/span[1]"));
                //openinvoicereport.getText().then((OIR1) => {
                unmatched.getText().then((UM1) => {
                    if (OIR1 == 'SUCCESS' && UM1 == 'SUCCESS') {
                        statusstamping['C' + rrn].v = 'Success';
                        XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');

                        //Clicking on Program grids
                        cashappspom.program_grids();
                        browser.sleep(10000);

                        //orgunitgridsindios(yy);
                    }
                    else if (OIR1 == 'FAILURE' || UM1 == 'FAILURE') {
                        statusstamping['C' + rrn].v = 'Failure';
                        XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                        //Clicking on Program grids
                        cashappspom.program_grids();
                        browser.sleep(10000);

                    }
                    else if (OIR1 == 'SKIPPED' || UM1 == 'SKIPPED') {
                        //Clicking on Initiate org unit grids
                        element(by.xpath("//span[contains(text(),'Initiate Org Unit Mapping')]")).click().then(function () {
                            console.log('Successfully Find the Locators');
                        }, function (err) {
                            console.error('Locators Error ' + err);
                            statusstamping['R2'].v = 'Locators Error : ' + err + '';
                            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                            //throw err;
                        });
                        browser.sleep(3000);

                        //Clicking on Initiate
                        element(by.xpath("//span[contains(text(),'initiate')]")).click().then(function () {
                            console.log('Successfully Find the Locators');
                        }, function (err) {
                            console.error('Locators Error ' + err);
                            statusstamping['R2'].v = 'Locators Error : ' + err + '';
                            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                            //throw err;
                        });
                        browser.sleep(5000);

                        //Clicking on Autorefresh
                        //cashappspom.auto_refresh1();
                        //browser.sleep(5000);

                        statusstamping['C' + rrn].v = 'Skipped & Initiated';
                        XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');

                        //Clicking on Program grids
                        cashappspom.program_grids();
                        browser.sleep(10000);

                    }
                    else if (OIR1 == 'EXCEPTION' || UM1 == 'EXCEPTION') {
                        statusstamping['C' + rrn].v = 'Exception';
                        XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');

                        //Clicking on Program grids
                        cashappspom.program_grids();
                        browser.sleep(10000);
                    }
                })
            })
        }
    })
}

function programgridsstatusvalidation(zz, rrn) {
    //Selecting the data package
    var DPname2 = element(by.xpath("//span[contains(text(),'" + WorksheetDios['G' + zz].v + "')]"));
    browser.sleep(2000);
    DPname2.getText().then((DPN) => {
        console.log('DPN' + DPN);
        var openinvoicereport2 = element.all(by.xpath("//div[@col-id='fileName' and @role='gridcell']"));
        openinvoicereport2.count().then((OIVR1) => {
            browser.sleep(2000);
            openinvoicereport2.getText().then((fileName) => {
                var array = [];
                for (let I = 0; I < OIVR1; I++) {
                    console.log(fileName[I]);
                    var A = fileName[I];
                    array.push(A);
                }
                console.log('array 284:' + array);

                var found1 = array.find(function (element) {
                    let re1 = new RegExp('unmatched');
                    return element.match(re1);
                })
                console.log('found1:' + found1);
                var found2 = array.find(function (element) {
                    let re2 = new RegExp('open');
                    //let re3 = new RegExp('openinvoicereport');
                    return element.match(re2);
                })
                console.log('found2:' + found2);
                var position1 = found1.indexOf(found1);
                var Fposition = position1 + 1;
                console.log('position1 :' + position1);
                console.log('Fposition : ' + Fposition);
                var position2 = array.indexOf(found2);
                var Sposition = position2 + 1;
                console.log('position2 : ' + position2);
                console.log('Sposition : ' + Sposition);
                successvalidation();

                openinvoicereport2.count().then((OIVR1) => {
                    if (OIVR1 == 2) {
                        //var pgmsuccesscount = element(by.xpath("//div[contains(text(),'SUCCESS')]/following::div/strong[contains(text(),'2')]"));
                        var pgmsuccesscount = element(by.xpath("//div[contains(text(),'SUCCESS')]/following::div/strong[1]"));
                        browser.wait(EC.visibilityOf(pgmsuccesscount), 300000).then(() => {
                            var openinvoicereport = element(by.xpath("//div[@col-id='fileName']/following::div[@ref='eCenterContainer' and @role='rowgroup']/div[@role='row'][" + Fposition + "]/div[@col-id='status']/span[1]"));
                            filestatus(openinvoicereport, Sposition);
                        })
                    }
                    else if (OIVR1 == 3) {
                        var pgmsuccesscount1 = element(by.xpath("//div[contains(text(),'SUCCESS')]/following::div/strong[1]"));
                        browser.wait(EC.visibilityOf(pgmsuccesscount1), 300000).then(() => {
                            var openinvoicereport = element(by.xpath("//div[@col-id='fileName']/following::div[@ref='eCenterContainer' and @role='rowgroup']/div[@role='row'][" + Fposition + "]/div[@col-id='status']/span[1]"));
                            filestatus(openinvoicereport, Sposition);
                        })
                    }
                    else if (OIVR1 == 4) {
                        var pgmsuccesscount2 = element(by.xpath("//div[contains(text(),'SUCCESS')]/following::div/strong[1]"));
                        browser.wait(EC.visibilityOf(pgmsuccesscount2), 300000).then(() => {
                            var openinvoicereport = element(by.xpath("//div[@col-id='fileName']/following::div[@ref='eCenterContainer' and @role='rowgroup']/div[@role='row'][" + Fposition + "]/div[@col-id='status']/span[1]"));
                            filestatus(openinvoicereport, Sposition);
                        })
                    }
                })
            })
        })
        function filestatus(Func_openinvoicereport, row) {
            var unmatched = element(by.xpath("//div[contains(text(),'unmatched')][@col-id='fileName']/following::div[@ref='eCenterContainer' and @role='rowgroup']/div[@role='row'][" + row + "]/div[@col-id='status']/span[1]"));
            Func_openinvoicereport.getText().then((OIR2) => {
                //var unmatched = element(by.xpath("//div[contains(text(),'unmatched')][@col-id='fileName']/following::div[@ref='eCenterContainer' and @role='rowgroup']/div[@role='row'][2]/div[@col-id='status']/span[1]"));
                unmatched.getText().then((UM2) => {
                    if (OIR2 == 'SUCCESS' && UM2 == 'SUCCESS') {
                        statusstamping['D' + rrn].v = 'Success';
                        XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');

                        var winhandles = browser.getAllWindowHandles();
                        winhandles.then(function (handles) {
                            var parentwindow = handles[0];
                            var childwindow = handles[1];

                            browser.switchTo().window(parentwindow);
                            browser.sleep(8000);

                            //Clicking on manual execution
                            cashappspom.manual_execution();
                            browser.sleep(5000);

                            //Uncheck the Multi Select Rec Types
                            var multirectypes = element(by.xpath("//mat-icon[contains(text(),'check_box')]"));
                            var multirectypes1 = element(by.xpath("//mat-icon[contains(text(),'check_box_outline_blank')]"));
                            multirectypes1.isPresent().then((MR) => {
                                if (MR == false) {
                                    browser.executeScript("arguments[0].click()", multirectypes);
                                }
                            });

                            //Filter button 3 lines for rec type
                            cashappspom.filter_rectypes();
                            browser.sleep(2000);

                            //Inside filter
                            cashappspom.inside_filter();
                            browser.sleep(2000);

                            //Uncheck the selectall checkbox
                            var uncheckselectallinmanualexec = element(by.xpath("//label[@ref='eSelectAllContainer']/div[@ref='eSelectAll']/span[contains(@class,'ag-icon-checkbox-checked')]"));
                            browser.executeScript("arguments[0].click()", uncheckselectallinmanualexec);
                            browser.sleep(2000);

                            //Enter the cash apps country name
                            cashappspom.rectypeselectionin_manualexec(WorksheetDios['C' + zz].v);
                            browser.sleep(5000);

                            //Click on Select All checkbox
                            var checkselectallinmanualexec = element(by.css("div[class='ag-filter-header-container'] label[ref='eSelectAllContainer'] div[ref='eSelectAll'] span"));
                            browser.executeScript("arguments[0].click()", checkselectallinmanualexec);

                            //Click on Apply filter
                            var applyfilterinmanualexec = element(by.xpath("//div[@ref='eButtonsPanel']/button[contains(text(),'Apply Filter')]"));
                            browser.executeScript("arguments[0].click()", applyfilterinmanualexec);

                            //Click on program
                            cashappspom.clickon_program();
                            browser.sleep(1000);

                            //Selecting the rectype
                            var selectrectype = element(by.xpath("//span[contains(text(),'" + WorksheetDios['C' + zz].v + "')]"));
                            browser.executeScript("arguments[0].click()", selectrectype);
                            browser.sleep(5000);
                            //Selecting the Processing Period in Manual Execution
                            var manualexecprocessingperiod = element(by.xpath("//b[contains(text(),'" + WorksheetDios['D2'].v + "')]"));
                            manualexecprocessingperiod.isPresent().then((a) => {
                                if (a == true) {
                                    browser.executeScript("arguments[0].click()", manualexecprocessingperiod);
                                }
                                else {
                                    cashappspom.ClickArrow_InManualExecution_ToSelectMonth();
                                    browser.executeScript("arguments[0].click()", manualexecprocessingperiod);
                                }
                            });

                            //Selecting executed/non-executed
                            var executed = element(by.xpath("//div[contains(text(),'Executed')]"));
                            executed.isPresent().then((b) => {
                                if (b == true) {
                                    //console.log('351 line crossed '+b);
                                    browser.sleep(3000).then(() => {
                                        var crcb = cashappspom.clickon_recgroupcheckbox();
                                        browser.sleep(5000);
                                        crcb.isPresent().then((CRCB) => {
                                            if (CRCB == true) {
                                                console.log('356 line crossed ' + CRCB);
                                                cashappspom.clickon_recgroupcheckbox1();
                                                //browser.executeScript("arguments[0].click()", CRCB);
                                                browser.sleep(2000);
                                                cashappspom.full_rerun();
                                                browser.sleep(5000);
                                                //Clicking on Yes
                                                cashappspom.yes_button();
                                                browser.sleep(6000);

                                            }
                                            else {
                                                var nonexecuted1 = element(by.xpath("//div[contains(text(),'Non-Executed')]"));
                                                browser.executeScript("arguments[0].click()", nonexecuted1);
                                                var crcb = cashappspom.clickon_recgroupcheckbox();
                                                crcb.isPresent().then((CRCB) => {
                                                    if (CRCB == true) {
                                                        browser.executeScript("arguments[0].click()", crcb);
                                                        browser.sleep(2000);
                                                        element(by.xpath("//div/button/span[contains(text(),'Execute')]")).click();
                                                        browser.sleep(5000);
                                                        //Clicking on Yes
                                                        cashappspom.yes_button();
                                                        browser.sleep(10000);


                                                    }
                                                    else {
                                                        statusstamping['J' + rrn].v = 'No rec groups in Manual execution';
                                                        XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');

                                                    }

                                                })
                                            }

                                        })
                                    })
                                    //cashappspom.clickon_recgroupcheckbox();

                                }
                                else {
                                    cashappspom.non_executed();
                                    browser.sleep(5000);
                                    cashappspom.clickon_recgroupcheckbox();
                                    browser.sleep(3000);

                                    var execute = element(by.xpath("//div/button/span[contains(text(),'Execute')]"));
                                    browser.executeScript("arguments[0].click()", execute);
                                    browser.sleep(5000);

                                    //Clicking on Yes
                                    cashappspom.yes_button();
                                    browser.sleep(10000);
                                }
                            });

                            //Bulk execution status
                            cashappspom.bulkexecution_status().click().then(function () {
                                console.log('Successfully Find the Locators');
                            }, function (err) {
                                console.error('Locators Error ' + err);
                                statusstamping['R2'].v = 'Locators Error : ' + err + '';
                                XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                                //throw err;
                            });
                            browser.sleep(7000);

                            //Fetching the Qpositions of the CasApp country in bulk execution and print in stamping sheet to fetech the proper Data 'Complete' Status
                            //div[contains(text(),'25 Feb 22, 10:37:29 am')]/preceding::*[@col-id='jobQSeqNo' and @role='gridcell']
                            var ReqStartDate = element.all(by.xpath("//div[@col-id='entityPersistDT' and @role='gridcell']")).get(0);
                            ReqStartDate.getText().then((RST) => {
                                console.log('504 LINE :-', RST);
                                statusstamping['P' + rrn].v = RST;
                                XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');

                            }).then(() => {
                                //var Qposition = element.all(by.xpath("//*[@col-id='jobQSeqNo' and @role='gridcell']")).get(0);
                                var Qposition = element(by.xpath("//div[contains(text(),'" + statusstamping['P' + rrn].v + "')]/preceding::*[@col-id='jobQSeqNo' and @role='gridcell']"));
                                Qposition.getText().then((QP) => {
                                    console.log('512 LINE :-', QP);
                                    statusstamping['I' + rrn].v = QP;
                                    XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');

                                });
                            });

                            //Clicking on Rec Plans
                            cashappspom.rec_plans();
                            browser.sleep(4000);
                            //Clicking on Exec button
                            var execbutton = element(by.xpath("//app-radio-group[1]/div[1]/div[1]/button[1]/span[1]"));
                            browser.executeScript("arguments[0].click()", execbutton);
                            browser.sleep(3000);

                            //Uncheck the Bosnia checkbox
                            cashappspom.uncheck_accrualsgeneral();
                            browser.sleep(3000);

                            browser.switchTo().window(childwindow);
                            browser.sleep(8000);
                        });
                        //orgunitgridsindios(yy);
                    }
                    else if (OIR2 == 'FAILURE' || UM2 == 'FAILURE') {
                        statusstamping['D' + rrn].v = 'Failure';
                        XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');

                    }
                    else if (OIR2 == 'SKIPPED' || UM2 == 'SKIPPED') {
                        //Clicking on Autorefresh
                        cashappspom.auto_refresh1();
                        browser.sleep(5000);

                        statusstamping['D' + rrn].v = 'Skipped & Initiated';
                        XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');

                    }
                    else if (OIR2 == 'EXCEPTION' || UM2 == 'EXCEPTION') {
                        statusstamping['D' + rrn].v = 'Exception';
                        XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                    }
                })
            })
        }
    })
}

function dateselection() {
    var res = new Date();
    var year = res.getFullYear();
    year = year.toString();
    var month = res.getMonth();
    //console.log(month);
    day = res.getDate().toString().padStart(2, "0");
    var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
    var monthName = months[res.getMonth()];
    //var shortmonth = monthName.substr(0, 3);
    var shortmonth = months.indexOf(monthName) + 1;

    let MyDate = (day + shortmonth + year);
    return MyDate;

}

function copy_files() {
    var res = new Date().toString();
    var d = res.substr(0, 21);
    var dd = d.replace(':', '_');
    console.log(dd);
    const fs = require('fs');
    fs.copyFile('./Statusstamping for CA.xlsx', './copy/Statusstamping for ' + dd + '.xlsx', (err) => {

        if (err) throw err;
        console.log('File was copied to destination');
    });
}

function successvalidation() {
    for (let j = 1; j <= 60; j++) {
        var receivedcount = element.all(by.xpath("//div[@col-id='fileName']/following::div[@ref='eCenterContainer' and @role='rowgroup']/div[@role='row']/div[@col-id='status']/span[1]"));
        receivedcount.getText().then((RF) => {
            console.log('125 line RF:' + RF);
            let successcount = RF.filter(s => s.includes('SUCCESS'));
            if (successcount.length < 2) {
                element(by.xpath("//mat-icon[contains(text(),'refresh')]")).click().then(function () {
                    console.log('Successfully Find the Locators-Refesh');
                }, function (err) {
                    console.error('Refresh Locators Error ' + err);
                    statusstamping['R2'].v = 'Locators Error : ' + err + '';
                    XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                    //throw err;
                });
                browser.sleep(2000);

            }
            else {
                j = 60
            }
        });
    }
}


describe('Cash apps', function () {

    function SynchronizationProcess() {
        browser.waitForAngularEnabled(false);
        browser.ignoreSynchronization = true;
    }

    it('Data Clear in Status Stamping Sheet', function () {
        for (let col = 0; col <= 17; col++) {
            //if (col != 4 & col != 7) {
            if (col != 4) {
                var ColumID = columnselection.identName(col);
                for (let row = 2; row <= 22; row++) {
                    // console.log('Cell : ' + ColumID + row)
                    statusstamping[ColumID + row].v = '-';
                    XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                }
            }
        }
    });

    it('Login with valid username and valid password', function () {

        SynchronizationProcess();
        //browser.get("" + WorksheetLogin['A2'].v + "", 4000);
        cashappspom.Get(WorksheetLogin['A2'].v);
        browser.getCurrentUrl().then((url) => {
            if (url !== 'https://oriflame-prod-sanity.prod.ame.gps.ihost.com/progAdmin/programs') {
                statusstamping['Q2'].v = 'Invalid URL';
                XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            }

        }).then(() => {
            if (statusstamping['Q2'].v !== 'Invalid URL') {
                browser.refresh();
                browser.manage().window().maximize();
                browser.manage().timeouts().implicitlyWait(60000);

                // element(by.id('username')).sendKeys('' + WorksheetLogin['B2'].v + '');
                cashappspom.enterUserName(WorksheetLogin['B2'].v);

                //Enter the password 
                //element(by.id('password')).sendKeys('' + WorksheetLogin['C2'].v + '');
                // element(by.id('ssword')).sendKeys('' + WorksheetLogin['C2'].v + '');

                cashappspom.enterPassword(WorksheetLogin['C2'].v);


                browser.executeScript('window.scrollTo(100,100);');

                //Click on Login
                //element(by.id('kc-login')).sendKeys(protractor.Key.ENTER);
                cashappspom.enterLogin();
                browser.sleep(3000);
                var InvalidStatus = element(by.css("span[class='kc-feedback-text']"));
                InvalidStatus.isPresent().then((IS) => {
                    if (IS == true) {

                        statusstamping['K2'].v = 'Invalid username or password.';
                        XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                    }
                });
            }
        });


    });

    it('Selecting the Oriflame Program', function () {
        if (statusstamping['K2'].v !== 'Invalid username or password.' && statusstamping['Q2'].v !== 'Invalid URL') {

            //Searching for the program
            cashappspom.search_prog(WorksheetLogin['D2'].v);
            var NoProgramName = element(by.xpath("//div[contains(@class,'no-data-msg')]"));

            NoProgramName.isPresent().then((NPN) => {

                if (NPN == true) {
                    NoProgramName.getText().then((t) => {
                        console.log(t);
                    });
                    statusstamping['L2'].v = 'No Programs Found';
                    XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');

                }
                else {

                    //Clicking on the program
                    element(by.xpath("//div[contains(text(),'" + WorksheetLogin['D2'].v + "')]")).click().then(function () {
                        console.log('Successfully Find the Locators');
                    }, function (err) {
                        console.error('Locators Error ' + err);
                        statusstamping['R2'].v = 'Locators Error : ' + err + '';
                        XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                        //throw err;
                    });

                }
            });
        }

    });

    it('Enter into Dios and doing the Manual Execution', function () {
        if (statusstamping['K2'].v !== 'Invalid username or password.' && statusstamping['Q2'].v !== 'Invalid URL') {
            var childbutton = element(by.xpath("//a[@aria-label=' Data Planning']"));
            childbutton.click().then(function () {
                console.log('Successfully Find the Locators-Data Planning');
            }, function (err) {
                console.error('Data Planning Locators Error ' + err);
                statusstamping['R2'].v = 'Locators Error : ' + err + '';
                XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                //throw err;
            });
            var winhandles = browser.getAllWindowHandles();
            winhandles.then(function (handles) {
                var parentwindow = handles[0];
                var childwindow = handles[1];
                browser.switchTo().window(childwindow);
                browser.sleep(15000);

                //Clicking on monitor
                cashappspom.click_monitor();
                browser.sleep(2000);

                //Clicking on View runtime status
                cashappspom.view_runtimestatus();
                browser.sleep(2000);
                var countycount = 0;
                //var reportrowno = 2;
                for (let y = FRow; y <= LRow; y++) {
                    if (WorksheetDios['F' + y].v == 'Y') {

                        countycount = countycount + 1;

                        //Clicking on Received files
                        cashappspom.received_files();
                        browser.sleep(2000);

                        //Clicking on Select orgunit
                        cashappspom.select_orgunit();
                        browser.sleep(2000);

                        //Selecting the orgunit
                        var EOU = cashappspom.enter_orgunit(WorksheetDios['I' + y].v);
                        EOU.isPresent().then((eou) => {
                            if (eou == true) {
                                var e1ou = cashappspom.enter_orgunit(WorksheetDios['I' + y].v);
                                browser.executeScript("arguments[0].click()", e1ou);
                            }
                            else {
                                statusstamping['M' + y].v = 'Org Unit is Wrong';
                                XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');

                            }
                        }).then(() => {
                            browser.sleep(2000);
                            if (statusstamping['M' + y].v !== 'Org Unit is Wrong') {
                                //clicking channel
                                cashappspom.select_channel();
                                browser.sleep(2000);

                                var ECh = cashappspom.enter_channel(WorksheetDios['J' + y].v);
                                ECh.isPresent().then((ech) => {
                                    if (ech == true) {
                                        var e1ch = cashappspom.enter_channel(WorksheetDios['J' + y].v);
                                        browser.executeScript("arguments[0].click()", e1ch);
                                    }
                                    else {
                                        statusstamping['N' + y].v = 'Channel is Wrong';
                                        XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');

                                    }
                                }).then(() => {
                                    browser.sleep(2000);
                                    if (statusstamping['N' + y].v !== 'Channel is Wrong') {
                                        //Clicking on select data package
                                        cashappspom.select_dp();
                                        browser.sleep(4000);

                                        //Selecting the data package
                                        var dp = element(by.xpath("//mat-option//span[contains(text(),'" + WorksheetDios['A' + y].v + "')]"));
                                        dp.isPresent().then((DP) => {
                                            if (DP == true) {
                                                var d1p = element(by.xpath("//mat-option//span[contains(text(),'" + WorksheetDios['A' + y].v + "')]"));
                                                browser.executeScript("arguments[0].click()", d1p);
                                            }
                                            else {
                                                statusstamping['O' + y].v = 'No Data Package Found';
                                                XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');

                                            }
                                        }).then(() => {
                                            browser.sleep(2000);
                                            if (statusstamping['O' + y].v !== 'No Data Package Found') {
                                                //Clicking on Autorefresh
                                                //cashappspom.auto_refresh1();
                                                //browser.sleep(5000);

                                                //Validating the date
                                                var res = new Date();
                                                var year = res.getFullYear();
                                                year = year.toString().substr(-2);
                                                day = res.getDate().toString().padStart(2, "0");
                                                var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
                                                var monthName = months[res.getMonth()];
                                                var shortmonth = monthName.substr(0, 3);

                                                let MyDate = (day + "-" + shortmonth + "-" + year);
                                                console.log(MyDate);

                                                var envdate = element(by.xpath("//span[contains(text(),'.zip')]"));
                                                envdate.getText().then((Edate) => {
                                                    var Edate1 = Edate.split(' ');
                                                    console.log('Edate1[0]' + Edate1[0]);
                                                    if (Edate1[0] == MyDate) {

                                                        //Click on Auto-refresh
                                                        //cashappspom.auto_refresh1();
                                                        browser.sleep(3000);

                                                        successvalidation();

                                                        //var receivedsuccesscount = element(by.xpath("//div[contains(text(),'SUCCESS')]/following::div/strong[contains(text(),'2')]"));
                                                        //browser.wait(EC.visibilityOf(receivedsuccesscount), 300000).then(() => {
                                                        browser.sleep(1000).then(() => {

                                                            //Selecting the data package
                                                            // statusvalidation(y, reportrowno);
                                                            statusvalidation(y, y);
                                                            console.log("734  line");
                                                            //console.log('643line:' + reportrowno);
                                                        }).then(() => {

                                                            //Org unit grids status validation
                                                            // orgunitstatusvalidation(y, reportrowno);
                                                            orgunitstatusvalidation(y, y);
                                                        }).then(() => {

                                                            //Program grids status validation
                                                            //programgridsstatusvalidation(y, reportrowno);
                                                            programgridsstatusvalidation(y, y);
                                                        });
                                                    }
                                                    else {
                                                        //Click on Configure
                                                        browser.sleep(2000);
                                                        var CICT = cashappspom.configure_inputchanneltext();
                                                        CICT.isPresent().then((cict) => {
                                                            if (cict != true) {
                                                                browser.sleep(2000);
                                                                cashappspom.configure_dios();
                                                                console.log("nikhitha");
                                                            }

                                                        }).then(() => {

                                                            browser.sleep(2000);

                                                            //Click on Configure Input Channels
                                                            cashappspom.configure_inputchannels();
                                                            browser.sleep(5000);

                                                            //Click on bridge
                                                            cashappspom.bridge_dios(WorksheetDios['I' + y].v).click().then(function () {
                                                                console.log('Successfully Find the Locators-Bridge Dios');
                                                            }, function (err) {
                                                                console.error('Bridge Dios Locators Error ' + err);
                                                                statusstamping['R2'].v = 'Locators Error : ' + err + '';
                                                                XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                                                                //throw err;
                                                            });
                                                            browser.sleep(4000);

                                                            //Click on Channel 

                                                            cashappspom.channel_dios(WorksheetDios['J' + y].v).click().then(function () {
                                                                console.log('Successfully Find the Locators-Channel Dios');
                                                            }, function (err) {
                                                                console.error('Channel Dios Locators Error ' + err);
                                                                statusstamping['R2'].v = 'Locators Error : ' + err + '';
                                                                XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                                                                //throw err;
                                                            });
                                                            browser.sleep(4000);

                                                            //Click on Data package
                                                            element(by.xpath("//div[@title='" + WorksheetDios['G' + y].v + "']/ancestor::div[@aria-label='templates']/button/span/mat-icon[contains(text(),'more_vert')]")).click().then(function () {
                                                                console.log('Successfully Find the Locators-Data Package');
                                                            }, function (err) {
                                                                console.error('Data Package Locators Error ' + err);
                                                                statusstamping['R2'].v = 'Locators Error : ' + err + '';
                                                                XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                                                                //throw err;
                                                            });
                                                            browser.sleep(4000);

                                                            //Click on version history
                                                            element(by.xpath("//span[contains(text(),'Version History')]")).click().then(function () {
                                                                console.log('Successfully Find the Locators-Version History');
                                                            }, function (err) {
                                                                console.error('Version History Locators Error ' + err);
                                                                statusstamping['R2'].v = 'Locators Error : ' + err + '';
                                                                XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                                                                //throw err;
                                                            });
                                                            browser.sleep(4000);

                                                            //Click on De-register
                                                            //element(by.xpath("//span[contains(text(),'DE-REGISTER')]")).click();
                                                            element(by.xpath("//div[@role='gridcell']/span/span[contains(text(),'REGISTER')]")).click().then(function () {
                                                                console.log('Successfully Find the Locators-De-register');
                                                            }, function (err) {
                                                                console.error('De-register Locators Error ' + err);
                                                                statusstamping['R2'].v = 'Locators Error : ' + err + '';
                                                                XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                                                                //throw err;
                                                            });
                                                            browser.sleep(10000);


                                                            //Click on Data package
                                                            element(by.xpath("//div[@title='" + WorksheetDios['G' + y].v + "']/ancestor::div[@aria-label='templates']/button/span/mat-icon[contains(text(),'more_vert')]")).click().then(function () {
                                                                console.log('Successfully Find the Locators-Data Package');
                                                            }, function (err) {
                                                                console.error('Data Package Locators Error ' + err);
                                                                statusstamping['R2'].v = 'Locators Error : ' + err + '';
                                                                XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                                                                //throw err;
                                                            });
                                                            browser.sleep(4000);

                                                            //Click on version history
                                                            element(by.xpath("//span[contains(text(),'Version History')]")).click().then(function () {
                                                                console.log('Successfully Find the Locators-Data Package');
                                                            }, function (err) {
                                                                console.error('Data Package Locators Error ' + err);
                                                                statusstamping['R2'].v = 'Locators Error : ' + err + '';
                                                                XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                                                                //throw err;
                                                            });
                                                            browser.sleep(4000);

                                                            //Click on register
                                                            element(by.xpath("//div[@role='gridcell']/span/span[contains(text(),'REGISTER')]")).click().then(function () {
                                                                console.log('Successfully Find the Locators-Data Package');
                                                            }, function (err) {
                                                                console.error('Data Package Locators Error ' + err);
                                                                statusstamping['R2'].v = 'Locators Error : ' + err + '';
                                                                XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                                                                //throw err;
                                                            });
                                                            browser.sleep(6000);

                                                            //Click on ok
                                                            var clickOK = element(by.xpath("//span[contains(text(),'OK')]"));
                                                            browser.executeScript("arguments[0].scrollIntoView()", clickOK);
                                                            browser.executeScript("arguments[0].click()", clickOK);
                                                            browser.sleep(6000);

                                                            //Clicking on View runtime status
                                                            cashappspom.view_runtimestatus();
                                                            browser.sleep(2000);

                                                            //Clicking on Received files
                                                            cashappspom.received_files();
                                                            browser.sleep(2000);

                                                            //Clicking on Select orgunit
                                                            cashappspom.select_orgunit();
                                                            browser.sleep(2000);

                                                            //Selecting the orgunit
                                                            var EOU1 = cashappspom.enter_orgunit(WorksheetDios['I' + y].v);
                                                            EOU1.isPresent().then((eou) => {
                                                                if (eou == true) {
                                                                    var e1ou = cashappspom.enter_orgunit(WorksheetDios['I' + y].v);
                                                                    browser.executeScript("arguments[0].click()", e1ou);
                                                                }
                                                                else {
                                                                    statusstamping['M' + y].v = 'Org Unit is Wrong';
                                                                    XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');

                                                                }
                                                            }).then(() => {
                                                                browser.sleep(2000);
                                                                if (statusstamping['M' + y].v !== 'Org Unit is Wrong') {
                                                                    //clicking channel
                                                                    cashappspom.select_channel();
                                                                    browser.sleep(2000);

                                                                    var ECh = cashappspom.enter_channel(WorksheetDios['J' + y].v);
                                                                    ECh.isPresent().then((ech) => {
                                                                        if (ech == true) {
                                                                            var e1ch = cashappspom.enter_channel(WorksheetDios['J' + y].v);
                                                                            browser.executeScript("arguments[0].click()", e1ch);
                                                                        }
                                                                        else {
                                                                            statusstamping['N' + y].v = 'Channel is Wrong';
                                                                            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');

                                                                        }
                                                                    });
                                                                }
                                                            }).then(() => {
                                                                if (statusstamping['N' + y].v !== 'Channel is Wrong') {
                                                                    browser.sleep(2000);

                                                                    //Clicking on select data package
                                                                    cashappspom.select_dp();
                                                                    browser.sleep(4000);


                                                                    //Selecting the data package
                                                                    var dp1 = element(by.xpath("//mat-option//span[contains(text(),'" + WorksheetDios['A' + y].v + "')]"));
                                                                    dp1.isPresent().then((DP) => {
                                                                        if (DP == true) {
                                                                            var d1p1 = element(by.xpath("//mat-option//span[contains(text(),'" + WorksheetDios['A' + y].v + "')]"));
                                                                            browser.executeScript("arguments[0].click()", d1p1);
                                                                        }
                                                                        else {
                                                                            statusstamping['O' + y].v = 'No Data Package Found';
                                                                            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');

                                                                        }
                                                                    });
                                                                    browser.sleep(5000);

                                                                }
                                                            });
                                                        }).then(() => {
                                                            if (statusstamping['O' + y].v !== 'No Data Package Found') {
                                                                //Clicking on Autorefresh
                                                                //cashappspom.auto_refresh1();
                                                                //browser.sleep(5000);

                                                                //statusvalidation(y, reportrowno);
                                                                // statusvalidation(y, y);
                                                                if (Edate1[0] == MyDate) {

                                                                    //Click on Auto-refresh
                                                                    //cashappspom.auto_refresh1();
                                                                    browser.sleep(3000);

                                                                    successvalidation();

                                                                    //var receivedsuccesscount = element(by.xpath("//div[contains(text(),'SUCCESS')]/following::div/strong[contains(text(),'2')]"));
                                                                    //browser.wait(EC.visibilityOf(receivedsuccesscount), 300000).then(() => {
                                                                    browser.sleep(1000).then(() => {
                                                                        //Selecting the data package
                                                                        // statusvalidation(y, reportrowno);
                                                                        statusvalidation(y, y);
                                                                        console.log("lakshmi nikhitha");
                                                                        //console.log('643line:' + reportrowno);
                                                                    }).then(() => {

                                                                        //Org unit grids status validation
                                                                        // orgunitstatusvalidation(y, reportrowno);
                                                                        orgunitstatusvalidation(y, y);
                                                                    }).then(() => {

                                                                        //Program grids status validation
                                                                        //programgridsstatusvalidation(y, reportrowno);
                                                                        programgridsstatusvalidation(y, y);
                                                                    });
                                                                }
                                                                else {
                                                                    statusstamping['A' + y].v = WorksheetDios['A' + y].v;
                                                                    statusstamping['B' + y].v = 'NO DATA';
                                                                    statusstamping['C' + y].v = 'NO DATA';
                                                                    statusstamping['D' + y].v = 'NO DATA';
                                                                    statusstamping['F' + y].v = 'NO DATA';
                                                                    statusstamping['G' + y].v = 'NO DATA';
                                                                    statusstamping['H' + y].v = 'No file Download/Upload';
                                                                    XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                                                                }
                                                            }
                                                        });

                                                    }
                                                });
                                            }
                                        });
                                    }
                                });
                            }
                        });

                    }
                    else if (WorksheetDios['F' + y].v == 'P') {
                        statusstamping['A' + y].v = WorksheetDios['A' + y].v;
                        statusstamping['B' + y].v = 'Already Processed';
                        statusstamping['C' + y].v = 'Already Processed';
                        statusstamping['D' + y].v = 'Already Processed';
                        statusstamping['F' + y].v = 'Already Processed';
                        statusstamping['G' + y].v = 'Already Processed';
                        statusstamping['H' + y].v = 'No file Download/Upload';
                        XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                    }
                    else {
                        statusstamping['A' + y].v = WorksheetDios['A' + y].v;
                        statusstamping['B' + y].v = 'NO DATA';
                        statusstamping['C' + y].v = 'NO DATA';
                        statusstamping['D' + y].v = 'NO DATA';
                        statusstamping['F' + y].v = 'NO DATA';
                        statusstamping['G' + y].v = 'NO DATA';
                        statusstamping['H' + y].v = 'No file Download/Upload';
                        XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                    }
                }
                console.log('countrycount :' + countycount);
                browser.switchTo().window(parentwindow);
                browser.sleep(10000);

                //Clicking on manual execution
                //cashappspom.manual_execution();
                //browser.sleep(5000);

                //Bulk execution status
                var bes = cashappspom.bulkexecution_status();
                bes.isPresent().then((BES) => {
                    if (BES == true) {
                        cashappspom.bulkexecution_status().click().then(function () {
                            console.log('Successfully Find the Locators-Data Package');
                        }, function (err) {
                            console.error('Data Package Locators Error ' + err);
                            statusstamping['R2'].v = 'Locators Error : ' + err + '';
                            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                            //throw err;
                        });

                        browser.sleep(7000);
                        //Rows per page
                        //element(by.xpath("//mat-paginator[1]/div[1]/div[1]/div[1]/mat-form-field[1]/div[1]/div[1]/div[1]/mat-select[1]/div[1]/div[2]/div[1]")).click();
                        //browser.sleep(1000);
                        //element(by.xpath("//body[1]/div[1]/div[2]/div[1]/div[1]/div[1]/mat-option[4]/span[contains(text(),'200')]")).click();
                        //browser.sleep(10000);

                        //element.all(by.xpath("//div[@col-id='jobQSeqNo'and@role='gridcell']")).getAttribute('title').then((Qpositionarray) => {
                        //  console.log("line 88 :" + Qpositionarray)
                        //Qpositionarray.push(v);
                        var Ycount = 0;
                        var YArray = [];
                        for (let K = FRow; K <= LRow; K++) {
                            if (WorksheetDios['F' + K].v == 'Y') {
                                Ycount++;
                                YArray.push(K);
                            }
                        }
                        //for (let l = 1; l <= 2; l++) {
                        if (Ycount <= 10) {
                            //Manual_Execution_Process(Ycount, 2);
                            Manual_Execution_Process(Ycount, YArray);
                            //break;
                        }
                        else {
                            element(by.xpath("//button[@aria-label='Next page']/span[1]")).click().then(function () {
                                console.log('Successfully Find the Locators-Next Page');
                            }, function (err) {
                                console.error('Next Page Locators Error ' + err);
                                statusstamping['R2'].v = 'Locators Error : ' + err + '';
                                XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                                //throw err;
                            });
                            var Ycount2 = Ycount - 10;
                            browser.sleep(2000);
                            var NewArray1 = YArray;
                            var NewArray = NewArray1.splice(0, Ycount2);

                            //Manual_Execution_Process(Ycount2, 11);
                            //Manual_Execution_Process(Ycount2, 2);
                            //Manual_Execution_Process(Ycount2, YArray);
                            Manual_Execution_Process(Ycount2, NewArray);
                            browser.sleep(5000);
                            element(by.xpath("//button[@aria-label='Previous page']/span[1]")).click().then(function () {
                                console.log('Successfully Find the Locators-Previous Page');
                            }, function (err) {
                                console.error('Previous Page Locators Error ' + err);
                                statusstamping['R2'].v = 'Locators Error : ' + err + '';
                                XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                                //throw err;
                            });
                            browser.sleep(5000);
                            var Ycount3 = Ycount - Ycount2;
                            //  var Ycount4 = Ycount2 + 2;
                            // YArray.splice(0, Ycount2);
                            // Manual_Execution_Process(Ycount3, Ycount4);
                            Manual_Execution_Process(Ycount3, YArray);
                            //break;
                        }
                    }
                });
                //}

                function Manual_Execution_Process(Ycount1, StatusNo) {

                    var columnnameBES = ['Requester Name', 'Exe Type', 'Request Mode', 'Exe State', 'Exe Start Date Time', 'Exe End Date Time', 'Job Type']
                    for (let cc = 0; cc < 7; cc++) {
                        var column = element(by.xpath("//span[contains(text(),'" + columnnameBES[cc] + "')]"));
                        browser.executeScript("arguments[0].scrollIntoView()", column);
                    }

                    console.log('Ycount:' + Ycount1);

                    var completionSTATUSno = StatusNo;
                    var zz = Ycount1 - 1;
                    for (let J = 0; J < Ycount1; J++) {

                        var jobexestatus = element.all(by.xpath("//div[@col-id='jobState' and @role='gridcell']"));

                        jobexestatus.get(J).getText().then((status) => {

                            if (status == 'Submitted' || status == 'Queued') {
                                for (let n = 0; n <= 35; n++) {
                                    var jobexestatus2 = element.all(by.xpath("//div[@col-id='jobState' and @role='gridcell']"));

                                    jobexestatus2.getText().then((status) => {
                                        console.log("991 line:" + status);
                                        //let results = status.filter(t => { return t !== "Completed" && t !== "GeneratingPreparerData" && t !== "PreparationFailure" && t !== "ExecutionFailure" && t !== "FlowRefreshmentFailure" && t != "PreparerDataGenerationFailure" });
                                        let results = status.filter(t => { return t !== "Completed" && t !== "PreparationFailure" && t !== "ExecutionFailure" && t !== "FlowRefreshmentFailure" && t != "PreparerDataGenerationFailure" });
                                        console.log('995 line :' + results.length);
                                        if (results.length != 0 || (results.length == 1 && results.includes("GeneratingPreparerData"))) {
                                            //if (status[0] != 'Completed' & status[1] != 'Completed' & status[2] != 'Completed' & status[3] != 'Completed' & status[4] != 'Completed' & status[5] != 'Completed' & status[6] != 'Completed') {
                                            browser.sleep(2000);
                                            element(by.xpath("//mat-icon[contains(text(),'refresh')]")).click().then(() => {
                                                browser.sleep(1000);
                                                ITERATION(completionSTATUSno);
                                            }).then(() => {
                                                browser.sleep(7000);
                                            });
                                        }
                                    });
                                }
                                var jobexestatus = element.all(by.xpath("//div[@col-id='jobState' and @role='gridcell']"));

                                jobexestatus.get(J).getText().then((status1) => {
                                    if (status1 == 'Completed' || status1 == 'Submitted' || status1 == 'ExecutionFailure' || status1 == 'FlowRefreshmentFailure' || status1 == 'PreparationFailure') {
                                        browser.sleep(5000);
                                        //element.all(by.xpath("//div[@col-id='jobCompletionDTOnSparkEngg' and @role='gridcell']")).get(J).getText().then((Date1) => {
                                        //browser.sleep(1000);
                                        console.log('1013 line zz:' + zz);
                                        element.all(by.xpath("//div[@col-id='jobPostedDTOnSparkEngg' and @role='gridcell']")).get(zz).click().then(() => {
                                            browser.sleep(2000);
                                            element(by.xpath("//div[@col-id='recType' and @role='gridcell']")).getText().then((x) => {
                                                browser.sleep(2000);
                                                element.all(by.xpath("//button//mat-icon[@aria-label='close']")).get(1).click();
                                                //rectypesarray.push(x)
                                                //statusarray.push(status1)
                                                statusstamping['F' + completionSTATUSno[J]].v = x;
                                                statusstamping['G' + completionSTATUSno[J]].v = status1;
                                                //statusstamping['H' + yy].v = Date1;
                                                XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                                                //  completionSTATUSno = completionSTATUSno + 1;
                                                zz--;
                                            })
                                        });
                                    }
                                })

                            }
                            else {
                                browser.sleep(5000);
                                //element.all(by.xpath("//div[@col-id='jobCompletionDTOnSparkEngg' and @role='gridcell']")).get(J).getText().then((Date1) => {
                                //browser.sleep(1000);
                                element.all(by.xpath("//div[@col-id='jobPostedDTOnSparkEngg' and @role='gridcell']")).get(zz).click().then(() => {
                                    //browser.sleep(2000);
                                    console.log('1039 line J:' + zz);
                                    element(by.xpath("//div[@col-id='recType' and @role='gridcell']")).getText().then((x) => {
                                        browser.sleep(2000);
                                        element.all(by.xpath("//button//mat-icon[@aria-label='close']")).get(1).click();
                                        rectypesarray.push(x)
                                        statusarray.push(status)
                                        statusstamping['F' + completionSTATUSno[J]].v = x;
                                        statusstamping['G' + completionSTATUSno[J]].v = status;
                                        //statusstamping['H' + yy].v = Date1;

                                        XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                                        //  completionSTATUSno = completionSTATUSno + 1;
                                        zz--;
                                    })
                                });
                            }

                        })
                    }
                }
                function ITERATION(Itr) {
                    element(by.xpath("//div[contains(text(),'REC PLANS')]")).click().then(function () {
                        console.log('Successfully Find the Locators-Rec Plans');
                    }, function (err) {
                        console.error('Rec Plans Locators Error ' + err);
                        statusstamping['R2'].v = 'Locators Error : ' + err + '';
                        XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                        //throw err;
                    });
                    browser.sleep(3000);
                    element(by.xpath("//div[contains(text(),'BULK EXECUTION STATUS')]")).click().then(function () {
                        console.log('Successfully Find the Locators-Bulk Execution Status');
                    }, function (err) {
                        console.error('Bulk Execution Status Data Package Locators Error ' + err);
                        statusstamping['R2'].v = 'Locators Error : ' + err + '';
                        XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                        //throw err;
                    });
                    browser.sleep(20000);

                    element(by.xpath("//*[text()='Q Position']/parent::div/preceding-sibling::span[@ref='eMenu']")).click().then(function () {
                        console.log('Successfully Find the Locators-Q Position');
                    }, function (err) {
                        console.error('Q Position Status Data Package Locators Error ' + err);
                        statusstamping['R2'].v = 'Locators Error : ' + err + '';
                        XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                        //throw err;
                    });
                    element(by.xpath("//div[@ref='tabHeader']/span[2]/span")).click().then(function () {
                        console.log('Successfully Find the Locators-Tab Header');
                    }, function (err) {
                        console.error('Tab Header Data Package Locators Error ' + err);
                        statusstamping['R2'].v = 'Locators Error : ' + err + '';
                        XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
                        //throw err;
                    });
                    browser.manage().timeouts().implicitlyWait(5000);
                    var UnSelectAll = element(by.xpath("//label[@ref='eSelectAllContainer']/div[@ref='eSelectAll']/span[contains(@class,'checkbox-checked')]"));
                    browser.executeScript("arguments[0].click()", UnSelectAll);
                    browser.sleep(2000);
                    element(by.xpath("//*[@placeholder='Search...']")).sendKeys("" + statusstamping['I' + Itr].v + "");
                    var SelectAll = element(by.xpath("//label[@ref='eSelectAllContainer']/div[@ref='eSelectAll']/span[contains(@class,'checkbox-unchecked')]"));
                    browser.executeScript("arguments[0].click()", SelectAll);
                    browser.sleep(2000);
                    var ApplyFilter = element(by.xpath("//*[@ref='eApplyButton' and @type='button']"));
                    browser.executeScript("arguments[0].click()", ApplyFilter);
                    var columnnameBES = ['Requester Name', 'Exe Type', 'Request Mode', 'Exe State', 'Exe Start Date Time', 'Exe End Date Time', 'Job Type']
                    for (let cc = 0; cc < 7; cc++) {
                        var column = element(by.xpath("//span[contains(text(),'" + columnnameBES[cc] + "')]"));
                        browser.executeScript("arguments[0].scrollIntoView()", column);
                    }

                }
            });
        }

    });

    it('Enter into All Recs', function () {
        if ((statusstamping['K2'].v !== 'Invalid username or password.' && statusstamping['L2'].v !== 'No Programs Found') && statusstamping['Q2'].v !== 'Invalid URL') {
            if ((statusstamping['M2'].v !== 'Org Unit is Wrong' && statusstamping['N2'].v !== 'Channel is Wrong') && statusstamping['O2'].v !== 'No Data Package Found') {

                browser.sleep(12000);
                // var ExeStatus = 2;
                for (let x = LRow; x >= FRow; x--) {
                    if (WorksheetDios['F' + x].v == 'Y' && (statusstamping['B' + x].v == 'Success' && statusstamping['C' + x].v == 'Success' && statusstamping['D' + x].v == 'Success')) {
                        var ExeStatus = x;
                        //Click on All Recs option
                        cashappspom.all_recs();
                        browser.sleep(12000);

                        //Click on Recording/Processing period dropdown
                        cashappspom.recprocess_period();
                        browser.sleep(4000);

                        //Selecting the period
                        element(by.xpath("//body/div[1]/div[2]/div[1]/div[1]/div[1]/mat-option/span/b[contains(text(),'" + WorksheetDios['D2'].v + "')]")).click().then(function () {
                            console.log('Successfully Find the Locators-Period');
                        }, function (err) {
                            console.error('Period Locators Error ' + err);
                            //throw err;
                        });
                        browser.sleep(8000);
                        console.log("worksheet dios :" + WorksheetDios['F' + x].v);
                        // if (WorksheetDios['F' + x].v == 'Y' && statusstamping['G' + x].v == 'Completed') {
                        if (WorksheetDios['F' + x].v == 'Y') {
                            //ExeStatus = 2;
                            console.log("execution status :" + ExeStatus);
                            console.log("G stamping  :" + statusstamping['G' + ExeStatus].v);

                            if (statusstamping['G' + ExeStatus].v == 'Completed') {
                                console.log("G stamping 1  :" + statusstamping['G' + ExeStatus].v);
                                //Click on rows per page
                                cashappspom.rowsper_page();
                                browser.sleep(4000);

                                //Selecting 200 rows per page
                                cashappspom.select200rows();
                                browser.sleep(2000);

                                //Clicking on Rec group name
                                cashappspom.rec_groupnameinallrecs();
                                browser.sleep(8000);

                                //Filter button 3 lines for rec group name
                                cashappspom.filter_recgroupnameinallrecs();
                                browser.sleep(2000);

                                //Inside filter
                                cashappspom.inside_filterinallrecs();
                                browser.sleep(2000);

                                //Uncheck the selectall checkbox
                                var uncheckselectallinallrecs = element(by.xpath("//label[@ref='eSelectAllContainer']/div[@ref='eSelectAll']/span[contains(@class,'ag-icon-checkbox-checked')]"));
                                browser.executeScript("arguments[0].click()", uncheckselectallinallrecs);

                                //Enter the Rec group name
                                cashappspom.searchrecgroupin_allrecs(WorksheetDios['E' + x].v);
                                browser.sleep(2000);

                                //Select the Select All checkbox
                                var checkselectallinallrecs = element(by.css("div[class='ag-filter-header-container'] label[ref='eSelectAllContainer'] div[ref='eSelectAll'] span"));
                                browser.executeScript("arguments[0].click()", checkselectallinallrecs);
                                browser.sleep(1000);

                                //Click on Apply filter
                                var applyfilterinallrecs = element(by.xpath("//div[@ref='eButtonsPanel']/button[contains(text(),'Apply Filter')]"));
                                browser.executeScript("arguments[0].click()", applyfilterinallrecs);
                                browser.sleep(1000);

                                //Click on program
                                cashappspom.clickon_program();
                                browser.sleep(1000);

                                //Clicking on the arrow before the rec
                                cashappspom.clickonarrow_beforerec();
                                browser.sleep(2000);

                                //Selecting the rec
                                cashappspom.select_rec();
                                browser.sleep(8000);

                                if (WorksheetDios['E' + x].v == 'Finland') {
                                    cashappspom.download_excel();
                                    browser.sleep(1000).then(() => {
                                        var EC = protractor.ExpectedConditions;
                                        var Download_Notification = element(by.xpath("//notifier-container/ul/li/notifier-notification"))
                                        browser.wait(EC.invisibilityOf(Download_Notification), 60000);
                                        // var glob = require("glob");                                 
                                        // var filesArray = glob.sync('./IH/RPA2IW/IH/Input'+'/*.zip');
                                        // console.log("Downloaded file count :- ",filesArray.length);
                                        Download_Notification.isPresent().then((DN) => {
                                            if (DN == true) {
                                                statusstamping['H' + x].v = 'Files are Downloaded';
                                                XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');

                                            }
                                            else {
                                                statusstamping['H' + x].v = 'NO Downloaded Files';
                                                XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');

                                            }
                                        });
                                    });

                                }
                                else {
                                    var firsttab = element.all(by.xpath("//div[contains(@class,'p-1 summary-left primary-bg-lighter')]"));
                                    firsttab.count().then((tab) => {
                                        for (let z = 0; z < tab; z++) {
                                            browser.sleep(16000);
                                            //var EC = protractor.ExpectedConditions;
                                            // var circle = element(by.xpath("//mat-spinner[@role='progressbar']"));
                                            //browser.wait(EC.invisibilityOf(circle), 100000);
                                            firsttab.get(z).click().then(function () {
                                                console.log('Successfully Find the Locators First Tab');
                                            }, function (err) {
                                                console.error('First Tab Locators Error ' + err);
                                                //throw err;
                                            });
                                            browser.sleep(12000);
                                            cashappspom.download_excel();
                                            browser.sleep(1000).then(() => {
                                                var EC = protractor.ExpectedConditions;
                                                var Download_Notification = element(by.xpath("//notifier-container/ul/li/notifier-notification"))
                                                browser.wait(EC.invisibilityOf(Download_Notification), 60000);
                                                // var glob = require("glob");                                 
                                                // var filesArray = glob.sync('./IH/RPA2IW/IH/Input'+'/*.zip');
                                                // console.log("Downloaded file count :- ",filesArray.length);
                                                Download_Notification.isPresent().then((DN) => {
                                                    if (DN == true) {
                                                        statusstamping['H' + x].v = 'Files are Downloaded';
                                                        XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');

                                                    }
                                                    else {
                                                        statusstamping['H' + x].v = 'NO Downloaded Files';
                                                        XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');

                                                    }
                                                });
                                            })

                                        }
                                    });
                                }

                                /*var MyFile, file;
                                MyFile = new ActiveXObject("Scripting.FileSystemObject");
                                file = myObject.GetFile("c:\\test.txt");
                                file.Move("C:\\");
                                /*browser.sleep(15000).then(() => {
                                    var child_process = require('child_process');
                                    child_process.exec('C:\\Users\\LakshmiNikhithaVanga\\Desktop\\Data_Upload.bat', function (error, stdout, stderr) {
                                        console.log(stdout);
                                    });
                                })*/

                                //Clicking on back arrow
                                cashappspom.back_arrow();
                                browser.sleep(5000);

                                //Clicking on Bulk Approval
                                element(by.xpath("//a[@aria-label=' Bulk Approval']")).click().then(function () {
                                    console.log('Successfully Find the Locators- Bulk Approval');
                                }, function (err) {
                                    console.error('Bulk Approval Locators Error ' + err);
                                    //throw err;
                                });
                                browser.sleep(5000);

                                //}
                            }
                            ExeStatus++;
                        }
                        //ExeStatus++
                    }
                };
            }
        }
    });

    it('Copy Stamping file in to another folder', function () {
        if ((statusstamping['K2'].v !== 'Invalid username or password.' && statusstamping['L2'].v !== 'No Programs Found') && statusstamping['Q2'].v !== 'Invalid URL') {
            if ((statusstamping['M2'].v !== 'Org Unit is Wrong' && statusstamping['N2'].v !== 'Channel is Wrong') && statusstamping['O2'].v !== 'No Data Package Found') {

                copy_files();
            }
        }
    });

    it('Upload files in SFTP SERVER', function () {
        if (statusstamping['K2'].v !== 'Invalid username or password.' && statusstamping['L2'].v !== 'No Programs Found' && statusstamping['Q2'].v !== 'Invalid URL') {
            if ((statusstamping['M2'].v !== 'Org Unit is Wrong' && statusstamping['N2'].v !== 'Channel is Wrong') && statusstamping['O2'].v !== 'No Data Package Found') {

                FiletransferToFolder();
                browser.sleep(15000).then(() => {
                    var child_process = require('child_process');
                    //child_process.exec('C:\\Users\\06318O744\\Documents\\ActionLog\\SanityAllModulesGit\\CashApps-Automation-Oriflame\\ibm-r2r-automation\\Data_Upload.bat', function (error, stdout, stderr) {

                    child_process.exec('Data_Upload.bat', function (error, stdout, stderr) {

                        if (error) {
                            console.error(error);
                            return;
                        }
                        if (stderr) {
                            console.error(stderr);
                            return;
                        }
                        console.info(stdout);
                    });
                });
            }

        }
    });

    it('Completion status Report', function () {
        if (statusstamping['K2'].v !== 'Invalid username or password.' && statusstamping['L2'].v !== 'No Programs Found' && statusstamping['Q2'].v !== 'Invalid URL') {
            if ((statusstamping['M2'].v !== 'Org Unit is Wrong' && statusstamping['N2'].v !== 'Channel is Wrong') && statusstamping['O2'].v !== 'No Data Package Found') {

                browser.sleep(5000).then(() => {
                    /*var child_process = require('child_process');
                    child_process.exec("allure serve",{app:'chrome'}, function (error, stdout, stderr) {
                        console.log(stdout);
                    });*/
                    var child_process = require('child_process');
                    child_process.execFile('Popup.bat', function (error, stdout, stderr) {

                        if (error) {
                            console.error(error);
                            return;
                        }
                        if (stderr) {
                            console.error(stderr);
                            return;
                        }
                        console.info(stdout);
                    });

                })
            }

        }
    });

});
