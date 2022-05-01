exports.config = {
  directConnect: true,

  //seleniumAddress:'http://localhost:4444/wd/hub',
  suites: {
    //Data_Update: ['DataUpdateMain.js'],
    Cashapp_14C: ['cashapp 14C prod sanity script.js'],
  },

  'capabilities': {
    'browserName': 'chrome',
    'shardTestFiles': true,
    'chromeOptions': {
      'prefs': {
        'download': {
          'prompt_for_download': false,
          'directory_upgrade': true,
          'default_directory': '' + __dirname + '\\IH\\RPA2IW\\IH\\Input',

        }
      }
    }
  },

  /*capabilities: {
    browserName: 'chrome',
    'shardTestFiles': true,
     chromeOptions: {
     args: ["--headless","--no-sandbox","--disable-gpu","--lang=en-US","--window-size=1920,1280" , "--disable-dev-shm-usage"]
    },
  },*/

  params: {
    env: {
      name: 'PRODUCTION'
    }
  },

  // Framework to use. Jasmine is recommended.
  framework: 'jasmine',
  //specs: ['cashapp 14C prod sanity script.js'],
  // Options to be passed to Jasmine.
  jasmineNodeOpts: {
    defaultTimeoutInterval: 2000000
  },

  onPrepare: function () {
    var AllureReporter = require('jasmine-allure-reporter');
    jasmine.getEnv().addReporter(new AllureReporter({
      resultsDir: 'allure-results'
    }));
  },
  onPrepare: function () {
    var AllureReporter = require('jasmine-allure-reporter');
    jasmine.getEnv().addReporter(new AllureReporter());
    jasmine.getEnv().afterEach(function (done) {
      browser.takeScreenshot().then(function (png) {
        allure.createAttachment('Screenshot', function () {
          return new Buffer(png, 'base64')
        }, 'image/png')();
        done();
      })
    });
  }

};