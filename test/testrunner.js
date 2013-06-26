var testrunner = require("qunit");

testrunner.run([
    {
        code : __dirname + "/MsOfficeWordHTMLWriter.js",
        tests: __dirname + "/MsOfficeWordHTMLWriter.js"
    }
]);