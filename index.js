const { Client, Events, GatewayIntentBits } = require("discord.js");
const client = new Client({ intents: [GatewayIntentBits.Guilds,
        GatewayIntentBits.GuildMessages, GatewayIntentBits.MessageContent] });
const config = require("./config.json");

const xlsx = require("xlsx");
const fs = require("fs");
const { SPPull } = require("sppull");

let searchedClasses = config.forClasses.split(" ");
let classMap = new Map();

const sppullContext = {
    siteUrl: config.url,
    creds: {
        username: config.siteUsername,
        password: config.sitePassword,
        online: true
    }
};

const sppullOptions = {
    spRootFolder: config.siteRootFolder,
    strictObjects: [config.siteFileName],
    dlRootFolder: "./data/"
};

client.login(config.discordToken);

client.on(Events.ClientReady, client => {
    let channelIDs = config.chatId.split(" ");
    for (let i = 0; i < searchedClasses.length; i++) {
        classMap.set(searchedClasses[i], client.channels.cache.get(channelIDs[i]));
    }

    start();
});

function start() {

    let date = new Date();

    if (date.getHours() === 18 && date.getMinutes() >= 50) {

        run();

    } else {

        let interval1 = setInterval(function () {

            let date2 = new Date();

            if (date2.getHours() === 18 && date2.getMinutes() >= 50) {

                run();
                clearInterval(interval1);

            }

        }, 600000);
        //1200000
    }

}

function run() {

    let date = new Date();
    let day = date.getDay();

    if (day === 0 || day === 1 || day === 3 || day === 4 || day === 2) {
        f1();
        setInterval(function () {
            f1();
        }, 86400000);

    } else {
        setInterval(function () {
            f1();
        }, 86400000);
       
    }
    function f1() {
        let date2 = new Date();
        let day = date2.getDay();
        if (day === 0 || day === 1 || day === 3 || day === 4 || day === 2) {
            download();
        }
    }
}

function download() {
    SPPull.download(sppullContext, sppullOptions).then(function () {
        fs.rename("./data/" + config.siteFileName, "./data/data.xlsx", () => {
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
        if (wb.SheetNames[i].replace(/[^0-9]/g, '') === dateFormat) {
            ws = wb.Sheets[wb.SheetNames[i]];
            break;
        }
    }

    xlsx.utils.sheet_add_aoa(ws, [["class", "h1", "h2", "h3", "h4", "h5", "h6", "h7", "h8", "h9", "h10"]], { origin: "A1" });
    xlsxData = xlsx.utils.sheet_to_json(ws);

    for (let i = 0; i < xlsxData.length; i++) {
        for (let j = 0; j < searchedClasses.length; j++) {
            if (xlsxData[i].class !== undefined) {
                let currentClass = searchedClasses[j];
                if (JSON.stringify(xlsxData[i].class).toLocaleLowerCase().match(currentClass)) {
                    let out = "";
                    for (let j = 1; j < 10; j++) {
                        if (xlsxData[i]['h' + j] !== undefined) {
                            out += "***hodina " + j + ":***\n" + xlsxData[i]['h' + j] + "\n\n";
                        }
                    }

                    if (out !== "") {
                        classMap.get(currentClass).send("Změny rozvrhu pro třídu " + currentClass + " na den: " + (date.getDate() + 1) + "." + (date.getMonth() + 1) + ".\n" + out);
                    }
                }
            }
        }
    }
}