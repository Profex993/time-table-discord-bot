discord bot for time tabel, our school uses a excel file for changes in time table and this bot is supposed to return the changes into a selected chat

*work in progress*

add folder named "data" and file named ".env" into the directory and edit the content like this:

URL="url of Microsoft SharePoint"
SITEUSERNAME="Microsoft username"
SITEPASSWORD="Microsoft password"
SITEROOTFOLDER="internal directory of the file in Microsoft SharePoint"
SITEORIGINALFILENAME="name of the file.xlsx"
DISCORDTOKEN="token for discord bot"


modules used:
discord.js : npm i discord.js
SPPull : npm install sppull --save-dev
xlsx : npm install xlsx
dotenv : npm install dotenv --save
