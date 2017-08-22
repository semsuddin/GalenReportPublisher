const fs = require('fs');
const publishConfig = JSON.parse(fs.readFileSync('publish.config'));
const json = JSON.parse(fs.readFileSync('execution.json'));

function getDateTime(shortTimeFormat){
    const date = new Date();
    let hour = date.getHours();
    hour = (hour < 10 ? "0" : "") + hour;
    let min = date.getMinutes();
    min = (min < 10 ? "0" : "") + min;
    let sec = date.getSeconds();
    sec = (sec < 10 ? "0" : "") + sec;
    const year = date.getFullYear();
    let month = date.getMonth() + 1;
    month = (month < 10 ? "0" : "") + month;
    let day = date.getDate();
    day = (day < 10 ? "0" : "") + day;
    if (shortTimeFormat != null)
    {
        return month + '/' + day + '/' + year + ' ' + hour + ':' + min;
    }
    else {
        return year + "_" + month + "_" + day + "_" + hour + "_" + min + "_" + sec;
    }
}

function getAppBuildNumber(cb) {
    const request = require('request');
    const cheerio = require('cheerio');
    const fs = require('fs');

    // Fetching url from execution.json (url has to be exported while running the test since its a separate process
    // and we have no way to access parameters directly since those are defined as a galen page object.

    let config = JSON.parse(fs.readFileSync("execution.json"));
    if (!config.Configuration.url){
        console.log("Url missing from configuration, e-mail sending skipped...");
        return;
    }
    let url = config.Configuration.url._text;
    sysAdminUrl = url.replace('wp','sysadmin');
    console.log('Getting Build Number by scrapping: ' + sysAdminUrl);
    request(sysAdminUrl, function(err, res, body) {
        if (err) return console.error(err);
        else if (res.statusCode === 200) {
            let $ = cheerio.load(body);
            let footer = $('#divCopyright');

            // Getting build number
            // Footer content has to contain build number and it has to be at the last place for example "© 2017 company 2.5.0.62"

            global.buildNumber = footer.text().split(" ").pop();
            console.log ('Build: ' + global.buildNumber);
            if (typeof cb === 'function')
                cb();
        }
    });
}

function createHtmlReport(cb){
    const fs = require("fs");
    const jsdom = require('jsdom');
    const reportJson = JSON.parse(fs.readFileSync('reports/report.json'));
    let htmlContent;

    jsdom.env(
        '<html></html>',
        ['http://code.jquery.com/jquery-1.5.min.js'],

        function (err, window) {
            if (err) console.error(err);

            const $ = window.jQuery;
            let html = $('html');
            let body = $('body');
            let head = $('head');

            head.append("<style>" +
                    //"body{background-color: #666666; color: white}" +
                "table.description{border-collapse: collapse; font-size: medium; font-family: 'Baskerville Old Face',monospace; border: none;}" +
                //"table.description tr td:nth-child(1){font-weight: bold;}" +
                "table.tests{border-collapse: collapse; border-color: white}" +
                "h1.title{font-size: xx-large; font-family: 'Segoe Marker' ,'Gungsuh', 'Harlow Solid Italic', 'Baskerville Old Face', 'Dotum', monospace; color: darkblue}" +
                "table.tests tr, table.tests td{border: 1px solid white; font-family: 'Dotum'; padding-left:5px; padding-right:5px;} " +
                "table.tests tr.headerRow{font-weight: bold; font-family: 'Segoe Marker'; background-color: #666666; color: white }" +
                //"tr.test" + (reportJson.tests.length-2) + "{border-bottom-style: solid;}" +
                "td.passed{color: limegreen}" +
                "td.failed{color: orangered}" +
                "td.information{padding-left:10px;}" +
                "</style>");

            body.append("<h1 class='title'>Device layout test report</h1>");
            body.append("<table class='description'>");
            //body.append(" <!--[if mso]><v:roundrect xmlns:v='urn:schemas-microsoft-com:vml' xmlns:w='urn:schemas-microsoft-com:office:word' href='http://www.EXAMPLE.com/' style='height:40px;v-text-anchor:middle;width:300px;' arcsize='10%' stroke='f' fillcolor='#d62828'><w:anchorlock/><center style='color:#ffffff;font-family:sans-serif;font-size:16px;font-weight:bold;'>Button Text Here!</center></v:roundrect><![endif]-->");
            let table = $('table.description');
            let location = publishConfig.LatestArchiveLocation._text;
            location = location.replace(location.split('\\').pop(), '');
            table.append("<tr><td><b>Environment: </b></td><td class='information'>" + json.Configuration.customer._text.toUpperCase() + "-" + json.Configuration.Env._text.toUpperCase() + " " + json.Configuration.appType._text + "</td></tr>");
            table.append("<tr><td><b>Results: </b></td><td class='information'><a href='file://" + publishConfig.LatestArchiveLocation._text + "'>Download</a>" + ",  " + "<a href='file://" + location + "'>Location</a>" + ",  " + "<a href='file://" + location + "/reports/report.html'>HTML</a></td></tr>");
            table.append("<tr><td><b>Time and Date: </b></td><td class='information'>" + getDateTime(true) + "</td></tr>");
            table.append("<tr><td><b>Build #: </b></td><td class='information'>" + global.buildNumber + "</td></tr>");

            body.append("<br>");

            body.append("<table class='tests'>");

            table = $('table.tests');

            table.append("<tr class='headerRow'><td>#</td><td>Test name and device tested</td><td>Total verifications processed</td><td>Passed</td><td>Failed</td><td>Duration</td></tr>");
            // $('table').append("<tr style='font-weight: bold'><td style='border: 1px solid black'>Test name and device</td><td style='border: 1px solid black'>Total verifications processed</td><td style='border: 1px solid black'>Passed</td><td style='border: 1px solid black'>Failed</td></tr><tr>");

            for (let i = 0; i < reportJson.tests.length; i++) {
                table.append("<tr class = 'test" + i + "'>");
                let tableRow = $('tr.test' + i);
                let duration = new Date(reportJson.tests[i].duration);
                let minutes = duration.getMinutes();
                let seconds = duration.getSeconds();
                if (minutes > 0)
                    duration = minutes + 'm ' + seconds + 's';
                else
                    duration = seconds + 's';
                let passedPercentage = reportJson.tests[i].statistic.passed / reportJson.tests[i].statistic.total * 100;
                let failedPercentage = reportJson.tests[i].statistic.failed / reportJson.tests[i].statistic.total * 100;


                tableRow.append("<td>" + (i+1) + ".</td>");
                tableRow.append("<td>" + reportJson.tests[i].name + "</td>");
                tableRow.append("<td class = 'meter'>" + reportJson.tests[i].statistic.total + "</td>");
                // let rowMeter = $('tr.meter');
                // rowMeter.append("<td style='width:" + passedPercentage + "%; background: #71E854; height: 10px;border: none;color: #229605;'></td>" +
                //                 "<td style='width:" + failedPercentage + "%; background: #FA644D; height: 10px;border: none;color: #D60F26;'></td>");
                tableRow.append("<td class = 'passed'>" + reportJson.tests[i].statistic.passed + "</td>");
                tableRow.append("<td class = 'failed'>" + reportJson.tests[i].statistic.errors + "</td>");
                tableRow.append("<td>" + duration + "</td>");
                // if (reportJson.tests[i].exceptionMessage != null)
                //     tableRow.append("<td><a href='" + reportJson.tests[i].statistic.exceptionMessage + ' ' + reportJson.tests[i].statistic.exceptionStacktrace + "'>Stacktrace</a></td>");
                // else
                //     tableRow.append("<td>None</td>");
            }


            htmlContent = html[0].outerHTML;

            fs.writeFile("reports/emailReport.html", htmlContent, function(err){
                if (err) return console.log(err);
                if (typeof cb === 'function')
                    cb(htmlContent);
            })
        });
}

function sendEmail(){
    let nodemailer = require('nodemailer');
    let fs = require('fs');

    // Please check publish.config file for mail server configurations
    // For example:

    // localSmtpConfig = {
    // host: 'smtp.server.local',
    // port: 587,
    // secure: false
    // };

    // remoteSmtpConfig = {
    // host: 'smtp.office365.com',
    // port: 587,
    // secure: false,
    // auth: {
    // user: 'user@company.com',
    // pass: 'password'
    // }
    // };

    let smtpConfig = publishConfig.Configuration.MailServer;

    const transporter = nodemailer.createTransport(smtpConfig);

    transporter.verify(function(error, success) {
        if (error) {
            console.log(error);
        }
        if (json.Configuration.email._text == null)
        {
            console.error('ERROR: Missing e-mail recipients from configuration file, please check that <email> node is not empty.');
        }
        else {
            console.log('Mail server is ready to take our messages. Sending e-mail...');
        }
    });

    createHtmlReport(
        function callback(htmlContent, err) {
            if (err) console.error(err);
            transporter.sendMail({
                from: 'Galen Publisher <galenpublisher@noreply.com>',
                to: json.Configuration.email._text,
                subject: 'Device Layout Testing Report',
                html: htmlContent
                //attachments: [{
                //    path: './reports/report.json'
                //}]
            }, function notifyDone(err) {
                if (err) return console.error(err);
                console.log("E-mail sent to: " + json.Configuration.email._text);
        });
    })
}

function archiveDirectory(dirLocation, callback){
    const fs = require('fs');
    const archiver = require('archiver');
    let dirName;

    if (dirLocation.indexOf("/") > -1)
        dirName = dirLocation.split("/").pop();
    else if (dirLocation.indexOf("\\") > -1)
        dirName = dirLocation.split("\\").pop();
    else
        dirName = dirLocation;

    console.log('Directory to archive: ' + dirName);

    const output = fs.createWriteStream(dirName + '.zip');
    const archive = archiver('zip', {
        zlib: {level: 9}
    });

    output.on('close', function (){
        console.log('Report has been successfully archived. Size: ' + (archive.pointer()/1024000).toFixed(3) + ' MB.');
        if (typeof callback === "function")
            callback();
    });

    archive.on('error', function(err) {
        console.log(err);
        throw err;
    });

    archive.pipe(output);
    archive.glob(dirLocation + '/**/*');
    archive.finalize();
}

function copyReport(galenResultsLocation){
    const fs = require('fs');
    const ncp = require('ncp').ncp;
    const os = require('os');
    const currentDateTime = getDateTime();
    const reportDateFolder = 'reports_' + currentDateTime + '\\';
    const reportDateName = 'reports_' + currentDateTime + '.zip';
    const reportDirLocation = galenResultsLocation + reportDateFolder + 'reports\\';
    let localArchiveLocation = galenResultsLocation + reportDateFolder + reportDateName;

    ncp.limit = 16;

    if (localArchiveLocation.indexOf('C:') > -1)
        localArchiveLocation = localArchiveLocation.replace('C:', os.hostname());
    //Export Remote Report location to execution json
    publishConfig.LatestArchiveLocation = {};
    publishConfig.LatestArchiveLocation._text = localArchiveLocation;

    console.log("Copying report to " + galenResultsLocation + reportDateFolder);

    fs.renameSync('reports.zip', reportDateName);

    if (!fs.existsSync(galenResultsLocation))
        fs.mkdirSync(galenResultsLocation);
    if (!fs.existsSync(galenResultsLocation + reportDateFolder))
        fs.mkdirSync(galenResultsLocation + reportDateFolder);

    fs.mkdirSync(reportDirLocation);

    ncp('reports', reportDirLocation, function (err) {
        if (err) {
            return console.error(err);
        }
        console.log('Report folder contents copied.');
    });

    ncp(reportDateName, galenResultsLocation + reportDateFolder + reportDateName, function (err) {
        if (err) {
            return console.error(err);
        }
        console.log('Report zip file copied.');

        if (publishConfig.Configuration.SendMail) {
            getAppBuildNumber( function callback(err){
                if (err) return console.error(err);
                sendEmail();
            });
        }
    });
}

function publishResults(){
    const fs = require('fs');
    const ncp = require('ncp').ncp;
    const glob = require('glob');
    ncp.limit = 16;

    const galenResultsLocation = publishConfig.Configuration.ArchiveReportLocation;

    //Create an archive of the report related to the time when the test was run and copy it to
    if (publishConfig.Configuration.ArchiveReport)
        archiveDirectory('reports', function () {
            copyReport(galenResultsLocation);
        });

    //Delete any old reports from the solution folder
    glob('reports_*.zip', function (er, files) {
        if (er) console.error(er);
        files.forEach(file => {
            fs.unlinkSync(file);
            console.log('Old report: ' + file + ' has been deleted from the solution root.');
        });
    })
}

module.exports = {
    publish: function () {
        publishResults();
    },
    getDateTime: function (dateShortFormat) {
        getDateTime(dateShortFormat);
    },
    archiveDirectory: function (dirLocation, callback) {
        archiveDirectory(dirLocation, callback);
    }
};
