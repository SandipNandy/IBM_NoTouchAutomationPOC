describe('Data Updates', function () {
         var child_process = require('child_process');
         it('Trigger SFTP server to get the country name and date updates', function () {
            //var child_process = require('child_process');
             browser.sleep(2000).then(() => {
                child_process.exec('Date_Update.bat');
             }).then(() => {
                browser.sleep(60000)
            });
        });
    
        it('Input data sheet updates with "Y" or "N" ', function () {
             browser.sleep(15000).then(() => {
                 //readSFTPTime();
                 child_process.exec('readSFTPTime.bat');
             }).then(() => {
                browser.sleep(6000)
            });
    
         });
    
     })
    