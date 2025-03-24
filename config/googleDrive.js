const { google } = require("googleapis");
require("dotenv").config();

const auth = new google.auth.GoogleAuth({
  keyFile: "credentials.json", // File kredensial dari Google Cloud
  scopes: ["https://www.googleapis.com/auth/drive"],
});

const drive = google.drive({ version: "v3", auth });

module.exports = drive;
