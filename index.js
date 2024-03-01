const { Client, Events, GatewayIntentBits } = require("discord.js");
const client = new Client({ intents: [GatewayIntentBits.Guilds, GatewayIntentBits.GuildMessages, GatewayIntentBits.MessageContent] });

const xlsx = require("xlsx");
const fs = require("fs");
const { SPPull } = require("sppull");
require("dotenv").config();

var searchedClasses = process.env.CLASSARRAY.split(" ");
var classMap = new Map();

const sppullContext = {
    siteUrl: process.env.URL,
    creds: {
        username: process.env.SITEUSERNAME,
        password: process.env.SITEPASSWORD,
        online: true
    }
};

const sppullOptions = {
    spRootFolder: process.env.SITEROOTFOLDER,
    strictObjects: [process.env.SITEORIGINALFILENAME],
    dlRootFolder: "./data/"
};

client.login(process.env.DISCORDTOKEN);

client.on(Events.ClientReady, client => {
    let channelIDs = process.env.CHATID.split(" ");
    for (let i = 0; i < searchedClasses.length; i++) {
        classMap.set(searchedClasses[i], client.channels.cache.get(channelIDs[i]));
    }

    download();
});

function download() {
    SPPull.download(sppullContext, sppullOptions).then(function () {
        fs.rename("./data/" + process.env.SITEORIGINALFILENAME, "./data/data.xlsx", () => {
            processXLSX();
        });
    }).catch(function (err) {
        console.log(err);
    });
}

function processXLSX() {
    let xlsxData, wb, ws;
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

    for (let i = 0; i < xlsxData.length; i++) {
        for (let j = 0; j < searchedClasses.length; j++) {
            if (xlsxData[i].class != undefined) {
                let currentClass = searchedClasses[j];
                if (JSON.stringify(xlsxData[i].class).toLocaleLowerCase().match(currentClass)) {
                    let out = "";
                    for (let j = 1; j < 10; j++) {
                        if (xlsxData[i]['h' + j] !== undefined) {
                            out += "***hodina " + j + ":***\n" + xlsxData[i]['h' + j] + "\n\n";
                        }
                    }

                    if (out != "") {
                        classMap.get(currentClass).send("Změny rozvrhu pro třídu " + currentClass + " na den: " + (date.getDate() + 1) + "." + (date.getMonth() + 1) + ".\n" + out);
                    }
                }
            }
        }
    }
}