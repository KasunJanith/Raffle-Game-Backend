// Raffle Endpoints for the new raffle draw system
import express from "express";

// Helper: get the internal sheet ID for a given sheet name
async function getSheetId(sheets, spreadsheetId, sheetName) {
  const res = await sheets.spreadsheets.get({
    spreadsheetId,
    ranges: [],
    includeGridData: false,
  });
  const sheet = res.data.sheets.find(s => s.properties.title === sheetName);
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found`);
  return sheet.properties.sheetId;
}

export function setupRaffleRoutes(app, sheets, spreadsheetId, normalizePhone) {
  // ========================
  // GET PARTICIPANTS
  // ========================
  app.get("/api/raffle/participants", async (req, res) => {
    try {
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: "Players!A2:C1000",
      });

      const rows = response.data.values || [];
      const participants = rows.map((row, index) => ({
        id: index,
        name: row[0] || "Unknown",
        phone: row[1] || "",
        email: row[2] || "",
      }));

      console.log("✅ Fetched participants count:", participants.length);

      res.json({
        success: true,
        participants,
      });
    } catch (error) {
      console.error("❌ Error fetching participants:", error.message);
      res.status(500).json({
        success: false,
        message: "Failed to fetch participants",
        error: error.message,
      });
    }
  });

  // ========================
  // SAVE WINNER (and remove from Players)
  // ========================
  app.post("/api/raffle/save-winner", async (req, res) => {
    try {
      const { participantId, participantName, participantPhone, timestamp } = req.body;

      if (!participantName || !participantPhone) {
        return res.status(400).json({
          success: false,
          message: "Participant name and phone are required",
        });
      }

      // 1. Append winner to Winners sheet
      await sheets.spreadsheets.values.append({
        spreadsheetId,
        range: "Winners!A:D",
        valueInputOption: "USER_ENTERED",
        requestBody: {
          values: [
            [
              participantName,
              participantPhone,
              new Date(timestamp).toLocaleString(),
              "WINNER",
            ],
          ],
        },
      });

      console.log("✅ Winner saved:", participantName);

      // 2. Find and remove the winner from the Players sheet
      try {
        // Get all players data (name and phone columns only)
        const playersRes = await sheets.spreadsheets.values.get({
          spreadsheetId,
          range: "Players!A2:B1000",
        });
        const rows = playersRes.data.values || [];
        let rowIndexToDelete = -1;
        for (let i = 0; i < rows.length; i++) {
          const rowName = rows[i][0] || "";
          const rowPhone = rows[i][1] || "";
          if (
            rowName.trim().toLowerCase() === participantName.trim().toLowerCase() &&
            rowPhone.trim() === participantPhone.trim()
          ) {
            rowIndexToDelete = i; // 0‑based, relative to A2
            break;
          }
        }

        if (rowIndexToDelete >= 0) {
          const sheetId = await getSheetId(sheets, spreadsheetId, "Players");
          // Row numbers in Google Sheets are 1‑based; data starts at row 2 (index 1)
          const startRowIndex = rowIndexToDelete + 1; // because header is row 1
          const endRowIndex = startRowIndex + 1;

          await sheets.spreadsheets.batchUpdate({
            spreadsheetId,
            requestBody: {
              requests: [
                {
                  deleteDimension: {
                    range: {
                      sheetId: sheetId,
                      dimension: "ROWS",
                      startIndex: startRowIndex,
                      endIndex: endRowIndex,
                    },
                  },
                },
              ],
            },
          });
          console.log(`🗑️ Removed ${participantName} from Players`);
        } else {
          console.warn("⚠️ Could not find player in Players sheet to delete");
        }
      } catch (deleteError) {
        // Log but don’t fail the whole request – the winner is already saved
        console.error("❌ Error removing player from sheet:", deleteError.message);
      }

      res.json({
        success: true,
        message: "Winner saved successfully",
      });
    } catch (error) {
      console.error("❌ Error saving winner:", error.message);
      res.status(500).json({
        success: false,
        message: "Failed to save winner",
        error: error.message,
      });
    }
  });

  // ========================
  // GET PREVIOUS WINNERS
  // ========================
  app.get("/api/raffle/winners", async (req, res) => {
    try {
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: "Winners!A2:D100",
      });

      const rows = response.data.values || [];
      const winners = rows.map((row) => ({
        name: row[0] || "Unknown",
        phone: row[1] || "",
        timestamp: row[2] || "",
        status: row[3] || "WINNER",
      }));

      res.json({
        success: true,
        winners,
      });
    } catch (error) {
      console.error("❌ Error fetching winners:", error.message);
      res.status(500).json({
        success: false,
        message: "Failed to fetch winners",
        error: error.message,
      });
    }
  });

  // ========================
  // ADD NEW PARTICIPANT
  // ========================
  app.post("/api/raffle/add-participant", async (req, res) => {
    try {
      const { name, phone, email } = req.body;

      if (!name || !phone) {
        return res.status(400).json({
          success: false,
          message: "Name and phone are required",
        });
      }
      await sheets.spreadsheets.values.append({
        spreadsheetId,
        range: "Players!A:C",
        valueInputOption: "USER_ENTERED",
        requestBody: {
          values: [[name, normalizePhone(phone), email || ""]],
        },
      });

      console.log("✅ Participant added:", name);

      res.json({
        success: true,
        message: "Participant added successfully",
      });
    } catch (error) {
      console.error("❌ Error adding participant:", error.message);
      res.status(500).json({
        success: false,
        message: "Failed to add participant",
        error: error.message,
      });
    }
  });
}