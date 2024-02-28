const { Client, Events, GatewayIntentBits } = require('discord.js');
const client = new Client({ intents: [GatewayIntentBits.Guilds, GatewayIntentBits.GuildMessages, GatewayIntentBits.MessageContent] });

const xlsx = require("xlsx");
const fs = require("fs");
const { SPPull } = require('sppull');
require('dotenv').config();

var wb;
var ws;
var xlsxData;
var dataSize;
var searchedClasses = ["c2c", "c2b", "c4c"];

client.login(process.env.DISCORDTOKEN);

client.on(Events.ClientReady, client => {
    client.channels.cache.get("1209202117529702400");

    download();
    //processXLSX();
});

function download() {
    const context = {
        siteUrl: process.env.URL,
        creds: {
            username: process.env.SITEUSERNAME,
            password: process.env.SITEPASSWORD,
            online: true
        }
    };

    const options = {
        spRootFolder: process.env.SITEROOTFOLDER,
        strictObjects: [process.env.SITEORIGINALFILENAME],
        dlRootFolder: "./data/"
    };

    SPPull.download(context, options).then(function () {
        fs.rename("./data/" + process.env.SITEORIGINALFILENAME, "./data/data.xlsx", () => {});
        processXLSX();
    }).catch(function (err) {
        console.log(err);
    });
}

function processXLSX() {
    let date = new Date();

    wb = xlsx.readFile("./data/data.xlsx");

    let dateFormat = (date.getDate() + 1) + "" + (date.getMonth() + 1) + "" + date.getFullYear();
    for (let i = wb.SheetNames.length - 1; i > 0; i--) {
        if (wb.SheetNames[i].replace(/[^0-9]/g, '') == dateFormat) {
            ws = wb.Sheets[wb.SheetNames[i]];
            break;
        }
    }

    xlsx.utils.sheet_add_aoa(ws, [["class", "h1", "h2", "h3", "h4", "h5", "h6", "h7", "h8", "h9", "h10"]], { origin: "A1" });
    xlsxData = xlsx.utils.sheet_to_json(ws);
    dataSize = xlsxData.length;

    for (let i = 0; i < dataSize; i++) {
        for (let j = 0; j < searchedClasses.length; j++) {
            if (xlsxData[i].class != undefined) {
                if (JSON.stringify(xlsxData[i].class).toLocaleLowerCase().match(searchedClasses[j])) {
                    console.log(i);
                    console.log(searchedClasses[j]);
                    console.log(JSON.stringify(xlsxData[i]));
                }
            }
        }
    }
}