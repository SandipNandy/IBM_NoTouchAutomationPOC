var XLSX = require('xlsx');
//var workbook = XLSX.readFile('C:/Users/JahnaviTunuguntla/Desktop/Cashapps aut.xlsx');
var workbook = XLSX.readFile('./cashapp 14C PROD SANITY excel.xlsx');
var statusworkbook = XLSX.readFile('./Statusstamping for CA.xlsx');
var statusstamping = statusworkbook.Sheets['Sheet1'];

var WorksheetLogin = workbook.Sheets['Sheet1'];
var WorksheetDios = workbook.Sheets[browser.params.env.name];
let x;

let Cash_apps = function () {

    this.Get = function (url) {
        browser.get(url);
    };

    let username1 = element(by.id('username'));
    //let username1 = element(by.id('userame'));
    this.enterUserName = function (U) {
        username1.sendKeys(U).then(function() {
            console.log('Successfully Find the Locators- User Name');
        }, function(err) {
            console.error('User Name Locators Error ' + err)
            statusstamping['R2'].v = 'Locators Error : '+err+'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx'); 
            //throw err;
        
        });
    };
    let password1 = element(by.id('password'));
    //let password1 = element(by.id('ssword'));

    this.enterPassword = function (P) {
        
        password1.sendKeys(P).then(function() {
            console.log('Successfully Find the Locators- Password');
        }, function(err) {
            console.error('Password Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
        
    };
    let login = element(by.id('kc-login'));
    this.enterLogin = function () {
        login.sendKeys(protractor.Key.ENTER).then(function() {
            console.log('Successfully Find the Locators- Log In');
        }, function(err) {
            console.error('Log In Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let clickcredential = element(by.css("button span[class='mat-button-wrapper'] div img[class='avatar']"));
    this.click_credential = function () {
        clickcredential.click().then(function() {
            console.log('Successfully Find the Locators - Credentials Logo');
        }, function(err) {
            console.error('Credentials Logo Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let selectrole = element(by.css("div[class='cdk-overlay-pane'] div[role='menu']")).element(by.css("div[class='mat-menu-content'] div mat-form-field"));
    this.select_role = function () {
        selectrole.click().then(function() {
            console.log('Successfully Find the Locators- Select Role');
        }, function(err) {
            console.error('Select Role Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let searchprog = element(by.css("input[placeholder='Search For Programs...']"));
    this.search_prog = function (Prog) {
        searchprog.sendKeys(Prog).then(function() {
            console.log('Successfully Find the Locators- Search Program');
        }, function(err) {
            console.error('Search Program Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let monitor = element(by.xpath("//mat-expansion-panel[3]/mat-expansion-panel-header/span/mat-panel-title[contains(text(),' Monitor ')]"));
    this.click_monitor = function () {
        monitor.click().then(function() {
            console.log('Successfully Find the Locators- Monitor');
        }, function(err) {
            console.error('Monitor Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let viewruntimestatus = element(by.xpath("//mat-nav-list[(@aria-label='View Runtime Status')]"));
    this.view_runtimestatus = function () {
        viewruntimestatus.click().then(function() {
            console.log('Successfully Find the Locators- View RunTime Status');
        }, function(err) {
            console.error('View RunTime Status Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let receivedfiles = element(by.xpath("//mat-nav-list[(@aria-label='Received files')]"));
    this.received_files = function () {
        receivedfiles.click().then(function() {
            console.log('Successfully Find the Locators- Received File');
        }, function(err) {
            console.error('Received Files Locators Error ' + err);
            //throw err;
        });
    };
    let selectorgunit = element(by.xpath("//mat-drawer-content/div/div/div[2]/mat-form-field[1]/div/div/div/mat-select/div/div[2]"));
    this.select_orgunit = function () {
        selectorgunit.click().then(function() {
            console.log('Successfully Find the Locators-Select Org Unit');
        }, function(err) {
            console.error('Select Org Unit Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    //let enterorgunit = element(by.xpath("//body/div[3]/div[2]/div[1]/div[1]/div[1]/mat-option/span[contains(text(),' Oriflame Automation Sanity ')]"));
    this.enter_orgunit = function (OrgUnit) {
        //enterorgunit.click();
        return element(by.xpath("//body/div[3]/div[2]/div[1]/div[1]/div[1]/mat-option/span[contains(text(),' "+ OrgUnit +" ')]"));//.click();
    };
    let selectchannel = element(by.xpath("//body/app-root[1]/main[1]/div[1]/div[1]/div[1]/div[2]/mat-drawer-container[1]/mat-drawer-content[1]/app-dios-dashboard[1]/div[1]/app-run-fmea[1]/div[1]/mat-drawer-container[1]/mat-drawer-content[1]/div[1]/div[1]/div[2]/mat-form-field[2]/div[1]/div[1]/div[1]"));
    this.select_channel = function () {
        selectchannel.click().then(function() {
            console.log('Successfully Find the Locators- Select Channel');
        }, function(err) {
            console.error('Select Channel Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };

   // let enterchannel = element(by.xpath("//body/div[3]/div[2]/div[1]/div[1]/div[1]/mat-option/span[contains(text(),' Oriflame CA ')]"));
    this.enter_channel = function (Channel) {
        //enterchannel.click();
        return element(by.xpath("//body/div[3]/div[2]/div[1]/div[1]/div[1]/mat-option/span[contains(text(),'"+Channel+"')]"));//.click();
    };
    let selectdp = element(by.xpath("//mat-drawer-content/div/div/div[2]/mat-form-field[3]/div/div/div/mat-select/div/div[2]"));
    this.select_dp = function () {
        selectdp.click().then(function() {
            console.log('Successfully Find the Locators- Select Data-Package');
        }, function(err) {
            console.error('Select Data-Package Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let rfautorefresh = element(by.xpath("//app-icon/img"));
    this.auto_refresh1 = function () {
        rfautorefresh.click().then(function() {
            console.log('Successfully Find the Locators- Auto Refresh');
        }, function(err) {
            console.error('Auto Refresh Locators Error ' + err);
            //throw err;
        });
    };
    let configure = element(by.xpath("//mat-panel-title[contains(text(),'Configure')]"));
    this.configure_dios = function () {
        configure.click().then(function() {
            console.log('Successfully Find the Locators-Configure Dios');
        }, function(err) {
            console.error('Configure Dios Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let configureinputchannels = element(by.xpath("//mat-nav-list[@aria-label='Configure Input Channels']"));
    this.configure_inputchannels = function () {
        configureinputchannels.click().then(function() {
            console.log('Successfully Find the Locators- Configure Input Channels');
        }, function(err) {
            console.error('Configure Input Channels Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let configureinputchanneltext = element(by.xpath("//div[contains(text(),'Configure Input Channels')]"));
    this.configure_inputchanneltext = function () {
        return configureinputchanneltext;
    };



    //let bridge = element(by.xpath("//div[contains(text(),'Oriflame Automation Sanity')]"));
    this.bridge_dios = function (BridgeDios) {
        // bridge.click().then(function() {
        //     console.log('Successfully Find the Locators');
        // }, function(err) {
        //     console.error('Locators Error ' + err);
        //     //throw err;
        // });
        return element(by.xpath("//div[contains(text(),'"+BridgeDios+"')]"));

    };
    //let channel = element(by.xpath("//div[contains(text(),'Oriflame CA')]"));
    this.channel_dios = function (ChannelDios) {
        // channel.click().then(function() {
        //     console.log('Successfully Find the Locators');
        // }, function(err) {
        //     console.error('Locators Error ' + err);
        //     //throw err;
        // });
        return element(by.xpath("//div[contains(text(),'"+ChannelDios+"')]"));
    };
    let successmsg = element.all(by.xpath("//span[contains(text(),'SUCCESS')]"));
    this.success_msg = function () {
        successmsg.click().then(function() {
            console.log('Successfully Find the Locators');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let orgunitgrids = element(by.xpath("//mat-nav-list[(@aria-label='Org unit Grids')]"));
    this.orgunit_grids = function () {
        orgunitgrids.click().then(function() {
            console.log('Successfully Find the Locators');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let orgunitrecordingperiod = element(by.xpath("//mat-drawer-content/div/div/div[2]/mat-form-field[5]/div/div/div/mat-select/div/div[2]"));
    this.orgunit_recordingperiod = function () {
        orgunitrecordingperiod.click().then(function() {
            console.log('Successfully Find the Locators');
        }, function(err) {
            console.error('Locators Error ' + err);
            //throw err;
        });
    };
    let orgrecordingperiod = element(by.xpath("//span[contains(text(),'" + WorksheetDios['B2'].v + "')]"));
    this.org_recordingperiod = function () {
        orgrecordingperiod.click().then(function() {
            console.log('Successfully Find the Locators- Org Recording Period');
        }, function(err) {
            console.error('Org Recording Period Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let programgrids = element(by.xpath("//mat-nav-list[(@aria-label='Program Grids')]"));
    this.program_grids = function () {
        programgrids.click().then(function() {
            console.log('Successfully Find the Locators');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let recordingperiod = element(by.xpath("//mat-drawer-content/div/div/div[2]/mat-form-field[6]/div/div/div/mat-select/div/div[2]"));
    this.recording_period = function () {
        recordingperiod.click().then(function() {
            console.log('Successfully Find the Locators');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            
            //throw err;
        });
    };
    let diosrecordingperiod = element(by.xpath("//span[contains(text(),'" + WorksheetDios['C2'].v + "')]"));
    this.dios_recordingperiod = function () {
        diosrecordingperiod.click().then(function() {
            console.log('Successfully Find the Locators');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let pgautorefresh = element(by.xpath("//app-icon/img"));
    this.auto_refresh2 = function () {
        pgautorefresh.click().then(function() {
            console.log('Successfully Find the Locators');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let manualexecution = element(by.xpath("//a[@aria-label='radio_button_unchecked Manual Execution']"));
    this.manual_execution = function () {
        manualexecution.click().then(function() {
            console.log('Successfully Find the Locators');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let rectypes = element(by.xpath("//div[@ref='eLabel']/span[contains(text(),'Rec Types ')]"));
    this.rec_types = function () {
        rectypes.click().then(function() {
            console.log('Successfully Find the Locators');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let filter1 = element(by.xpath("//span[@ref='eMenu']/span"));
    this.filter_rectypes = function () {
        filter1.click().then(function() {
            console.log('Successfully Find the Locators');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let insidefilter = element(by.xpath("//div[@class='ag-menu ag-ltr']/div[1]/div[@ref='tabHeader']/span[2]/span[@class='ag-icon ag-icon-filter']"));
    this.inside_filter = function () {
        insidefilter.click().then(function() {
            console.log('Successfully Find the Locators-Inside Filter');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let rectypeselectioninmanualexec = element(by.xpath("//div[@class='ag-input-wrapper']/input[@placeholder='Search...']"));
    this.rectypeselectionin_manualexec = function (rectypeselection) {
        rectypeselectioninmanualexec.sendKeys(rectypeselection).then(function() {
            console.log('Successfully Find the Locators-Rec Types Selection in Manual Execution');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let clickonprogram = element(by.xpath("//div/p[contains(text(),'PROGRAM')]"));
    this.clickon_program = function () {
        clickonprogram.click().then(function() {
            console.log('Successfully Find the Locators- Click on Program');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    var nonexecuted = element(by.xpath("//app-radio-group[1]/div[1]/div[2]/button[1]/span[1]/mat-icon[1]"));
    this.non_executed = function () {
        nonexecuted.click().then(function() {
            console.log('Successfully Find the Locators- Non executed');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let fullrerun = element(by.xpath("//span[contains(text(),'Full Re-Run')]"));
    this.full_rerun = function () {
        fullrerun.click().then(function() {
            console.log('Successfully Find the Locators- Full Re-run');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let yesbutton = element(by.xpath("//button/span[contains(text(),'YES')]"));
    this.yes_button = function () {
        yesbutton.click().then(function() {
            console.log('Successfully Find the Locators- YES button');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let bulkexecutionstatus = element(by.xpath("//app-manual-rec-exe/div/div/div/mat-button-toggle-group/mat-button-toggle[2]"));
    this.bulkexecution_status = function () {
        return bulkexecutionstatus;
    };
    let recplans = element(by.xpath("//app-manual-rec-exe/div/div/div/mat-button-toggle-group/mat-button-toggle[1]"));
    this.rec_plans = function () {
        recplans.click().then(function() {
            console.log('Successfully Find the Locators- Rec plans');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let uncheckaccrualsgeneral = element(by.xpath("//ag-grid-angular[1]/div[1]/div[2]/div[1]/div[3]/div[2]/div[1]/div[1]/div[1]/div/div[1]/span[1]/span[1]"));
    this.uncheck_accrualsgeneral = function () {
        uncheckaccrualsgeneral.click().then(function() {
            console.log('Successfully Find the Locators-Uncheck Accrual General');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let clickonrecgroupcheckbox = element(by.xpath("//div[@role='gridcell' and @col-id='groupName']"));
    this.clickon_recgroupcheckbox = function () {
        return clickonrecgroupcheckbox;
    };

    let clickonrecgroupcheckbox1 = element(by.xpath("//ag-grid-angular[1]/div[1]/div[2]/div[1]/div[3]/div[1]/div[1]/div[1]/div[1]/span[1]/span[2]"));
    this.clickon_recgroupcheckbox1 = function () {
        clickonrecgroupcheckbox1.click().then(function() {
            console.log('Successfully Find the Locators-Click on Rec Group Checkbox');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };

    let uncheck = element(by.xpath("//label[@ref='eSelectAllContainer']/div[@ref='eSelectAll']/span[contains(@class,'ag-icon-checkbox-checked')]"));
    this.uncheck_checkbox = function () {
        uncheck.click().then(function() {
            console.log('Successfully Find the Locators- Uncheck');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let applyfilter = element(by.xpath("//div[@ref='eButtonsPanel']/button[contains(text(),'Apply Filter')]"));
    this.apply_filter = function () {
        applyfilter.click().then(function() {
            console.log('Successfully Find the Locators- Apply Filter');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    /*let selectcheckbox = element(by.xpath("//span[contains(text(),'" + WorksheetDios['C' + x].v + "')]"));
    this.select_checkbox = function () {
        selectcheckbox.click();
    };*/
    let processingperiod = element(by.xpath("//nav/div[2]/div/div/a[8]/span/div/b[contains(text(),'Apr-21')]"));
    this.processing_period = function () {
        processingperiod.click().then(function() {
            console.log('Successfully Find the Locators- Processing Period');
        }, function(err) {
            console.error('Locators Error ' + err);
            //throw err;
        });
    };
    let selectingrec = element(by.xpath("//ag-grid-angular[1]/div[1]/div[2]/div[1]/div[3]/div[1]/div[1]/div[1]/div[1]/span[1]/span[2]"));
    this.selecting_rec = function () {
        selectingrec.click().then(function() {
            console.log('Successfully Find the Locators- Selecting Recs');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let allrecs = element(by.xpath("//span[contains(text(),'All Recs')]"));
    this.all_recs = function () {
        allrecs.click().then(function() {
            console.log('Successfully Find the Locators-All recs');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    /*let recprocessperiod = element(by.xpath("//app-all-recs-dashboard[1]/div[1]/div[1]/div[1]/div[2]/div[1]/mat-form-field[1]/div[1]/div[1]/div[1]/mat-select[1]/div[1]/div[2]/div[1]"));
    this.recprocess_period = function () {
        recprocessperiod.click();
    };*/
    let recprocessperiod = element(by.xpath("//body/app-root[1]/main[1]/div[1]/div[1]/div[1]/div[2]/mat-drawer-container[1]/mat-drawer-content[1]/app-dashboard[1]/div[1]/mat-drawer-container[1]/mat-drawer-content[1]/div[1]/app-program-admin-dashboard[1]/div[1]/app-admin-active-program[1]/div[1]/div[1]/app-recon-program-config[1]/div[1]/div[1]/app-all-recs-dashboard[1]/div[1]/div[1]/div[1]/div[2]/div[1]/app-period-selection[1]/mat-form-field[1]/div[1]/div[1]/div[1]/mat-select[1]/div[1]/div[1]/span[1]"));
    this.recprocess_period = function () {
        recprocessperiod.click().then(function() {
            console.log('Successfully Find the Locators-Rec Process Period');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };

    let rowsperpage = element(by.xpath("//mat-paginator[1]/div[1]/div[1]/div[1]/mat-form-field[1]/div[1]/div[1]/div[1]/mat-select[1]/div[1]/div[2]/div[1]"));
    this.rowsper_page = function () {
        rowsperpage.click().then(function() {
            console.log('Successfully Find the Locators-Rows Per Page');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let selectnoofrows = element(by.xpath("//span[contains(text(),'200')]"));
    this.select200rows = function () {
        selectnoofrows.click().then(function() {
            console.log('Successfully Find the Locators-Select No. of Rows');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let recgroupnameinallrecs = element(by.xpath("//div[@ref='eLabel']/span[contains(text(),'Rec Group Name')]"));
    this.rec_groupnameinallrecs = function () {
        recgroupnameinallrecs.click().then(function() {
            console.log('Successfully Find the Locators-Rec Group Name in All Recs');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let filterinallrecs = element(by.xpath("//div[@col-id='recgroupname']//div[3]/span/span"));
    this.filter_recgroupnameinallrecs = function () {
        filterinallrecs.click().then(function() {
            console.log('Successfully Find the Locators-Rec Group Name in All Recs');
        }, function(err) {
            console.error('Locators Error ' + err);
            //throw err;
        });
    };
    let insidefilterinallrecs = element(by.xpath("//div[@class='ag-menu ag-ltr']/div[1]/div[@ref='tabHeader']/span[2]/span[@class='ag-icon ag-icon-filter']"));
    this.inside_filterinallrecs = function () {
        insidefilterinallrecs.click().then(function() {
            console.log('Successfully Find the Locators-Inside FilterAll Reces');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let clearrecgroupinallrecs = element(by.xpath("//div[@class='ag-input-wrapper']/input[@placeholder='Search...']"));
    this.clearrecgroupin_allrecs = function () {
        clearrecgroupinallrecs.clear().then(function() {
            console.log('Successfully Find the Locators-Clear RecGroup All Reces');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let searchrecgroupinallrecs = element(by.xpath("//div[@class='ag-input-wrapper']/input[@placeholder='Search...']"));
    this.searchrecgroupin_allrecs = function (recgroup) {
        searchrecgroupinallrecs.sendKeys(recgroup).then(function() {
            console.log('Successfully Find the Locators-Search Rec Group in All Recs');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    this.searchrecgroupin_allrecs1 = function () {
        searchrecgroupinallrecs.clear().then(function() {
            console.log('Successfully Find the Locators-Search Rec Group in All Recs');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let clickonarrowbeforerec = element(by.xpath("//body/app-root[1]/main[1]/div[1]/div[1]/div[1]/div[2]/mat-drawer-container[1]/mat-drawer-content[1]/app-dashboard[1]/div[1]/mat-drawer-container[1]/mat-drawer-content[1]/div[1]/app-program-admin-dashboard[1]/div[1]/app-admin-active-program[1]/div[1]/div[1]/app-recon-program-config[1]/div[1]/div[1]/app-all-recs-dashboard[1]/div[1]/div[1]/div[3]/app-workflow-grid[1]/div[1]/div[1]/ag-grid-angular[1]/div[1]/div[2]/div[1]/div[3]/div[1]/div[1]/div[1]/span[1]/span[2]/span[1]"));
    this.clickonarrow_beforerec = function () {
        clickonarrowbeforerec.click().then(function() {
            console.log('Successfully Find the Locators-Click On Arrow Before Reces');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let selecttherec = element(by.xpath("//ag-grid-angular[1]/div[1]/div[2]/div[1]/div[3]/div[4]/div[1]/div[1]/div[2]/div[1]/div[2]/div[1]/div[3]/div[2]/div[1]/div[1]/div[1]/div[3]"));
    this.select_rec = function () {
        selecttherec.click().then(function() {
            console.log('Successfully Find the Locators');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let downloadexcel = element(by.xpath("//app-table-actions/div//div[4]/button[@aria-label='Export to Excel']/span/i"));
    this.download_excel = function () {
        downloadexcel.click().then(function() {
            console.log('Successfully Find the Locators-DownLoad Excel');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let backarrow = element(by.xpath("//nav/div[2]/div/div/button/span/mat-icon[contains(text(),'arrow_back')]"));
    this.back_arrow = function () {
        backarrow.click().then(function() {
            console.log('Successfully Find the Locators-Back Arrow');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let selectrectype = element(by.xpath("//span[contains(text(),'" + WorksheetDios['C2'].v + "')]"));
    this.select_rectype = function () {
        selectrectype.click().then(function() {
            console.log('Successfully Find the Locators-Select Rec types');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
    let mousemoverecTypes = element(by.xpath("//div[@ref='eLabel']/span[contains(text(),'Rec Types')]"));
    //span[contains(text(),'REC TYPES')]
    this.mousemove_recTypes = function () {
        return mousemoverecTypes;
    };
    let ClickArrowInManualExecutionToSelectMonth = element(by.xpath("//app-global-title-header/mat-toolbar[1]/div[2]/nav[1]/div[3]/div[1]"));
    this.ClickArrow_InManualExecution_ToSelectMonth = function () {
        ClickArrowInManualExecutionToSelectMonth.click().then(function() {
            console.log('Successfully Find the Locators-Click Arrow In ManualExecution');
        }, function(err) {
            console.error('Locators Error ' + err);
            statusstamping['R2'].v = 'Locators Error : ' + err +'';
            XLSX.writeFile(statusworkbook, './Statusstamping for CA.xlsx');
            //throw err;
        });
    };
}
module.exports = new Cash_apps();