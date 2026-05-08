import { sheets, SPREADSHEET_ID } from "./googleSheets";

async function saveGame3(data) {
  const { mobile, name, answer } = data;

  const time = new Date().toISOString();

  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: "Game3!A:D",
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [[mobile, name, answer, time]],
    },
  });
}

export default saveGame3;