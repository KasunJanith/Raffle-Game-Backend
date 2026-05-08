import { sheets, SPREADSHEET_ID } from "./googleSheets";

async function saveGame2(data) {
  const { mobile, name, answer } = data;

  const time = new Date().toISOString();

  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: "Game2!A:D",
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [[mobile, name, answer, time]],
    },
  });
}

export default saveGame2;