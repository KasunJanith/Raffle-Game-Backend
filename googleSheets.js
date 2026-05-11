const { google } = require("googleapis");

const auth = new google.auth.GoogleAuth({
  keyFile: "credentials.json", // your service account file
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

const sheets = google.sheets({
  version: "v4",
  auth,
});

const SPREADSHEET_ID = "1gXlGXkaSbPQx9y8X_hO4wh4MNMQldRu0Km8N9Ax1R8w";

module.exports = { sheets, SPREADSHEET_ID };