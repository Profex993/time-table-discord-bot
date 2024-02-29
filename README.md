discord bot for time tabel, our school uses a excel file for changes in time table and this bot is supposed to return the changes into a selected chat

*work in progress*

add folder named "data" and file named ".env" into the directory and edit the content like this:

URL="url of Microsoft SharePoint" `<br>`
SITEUSERNAME="Microsoft username" `<br>`
SITEPASSWORD="Microsoft password" `<br>`
SITEROOTFOLDER="internal directory of the file in Microsoft SharePoint" `<br>`
SITEORIGINALFILENAME="name of the file.xlsx" `<br>`
DISCORDTOKEN="token for discord bot"`<br>`
CLASSARRAY="keywoards of the row you want to find in the xlsx document, split with ";" "<br>
CHATID="discord channel IDs of the text channels you want to send the message to, split by ";" "<br>

modules used: `<br>`
discord.js : npm i discord.js `<br>`
SPPull : npm install sppull --save-dev `<br>`
xlsx : npm install xlsx `<br>`
dotenv : npm install dotenv --save `<br>`
