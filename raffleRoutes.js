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
  // Players columns: A = Ticket Display ID, B = Buyer's Name
  // ========================
  app.get("/api/raffle/participants", async (req, res) => {
    try {
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: "Players!A2:B1000",
      });

      const rows = response.data.values || [];
      const participants = rows
        .filter(row => row[0] || row[1]) // skip completely empty rows
        .map((row, index) => ({
          id: index,
          phone: row[0] || "Unknown",        // Ticket Display ID → phone (frontend expects this)
          name: row[1] || "Unknown",          // Buyer's Name → name
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
  // Winners columns: A = Winner Name, B = Ticket Display ID, C = Date/Time, D = Status
  // ========================
  app.post("/api/raffle/save-winner", async (req, res) => {
    try {
      const { participantName, participantPhone, timestamp } = req.body;

      if (!participantName || !participantPhone) {
        return res.status(400).json({
          success: false,
          message: "Participant name and Ticket Display ID are required",
        });
      }

      // 1. Append winner to Winners sheet
      // Columns: A = Winner Name, B = Ticket Display ID, C = Date/Time, D = Status
      await sheets.spreadsheets.values.append({
        spreadsheetId,
        range: "Winners!A:D",
        valueInputOption: "USER_ENTERED",
        requestBody: {
          values: [
            [
              participantName,                                    // A: Winner Name
              participantPhone,                                   // B: Ticket Display ID
              new Date(timestamp).toLocaleString(),               // C: Date/Time
              "WINNER",                                           // D: Status
            ],
          ],
        },
      });

      console.log("✅ Winner saved:", participantName, "| ID:", participantPhone);

      // 2. Find and remove the winner from the Players sheet
      // Players: A = Ticket Display ID, B = Buyer's Name
      try {
        const playersRes = await sheets.spreadsheets.values.get({
          spreadsheetId,
          range: "Players!A2:B1000",
        });
        const rows = playersRes.data.values || [];
        let rowIndexToDelete = -1;
        for (let i = 0; i < rows.length; i++) {
          const rowId = (rows[i][0] || "").trim();
          const rowName = (rows[i][1] || "").trim();
          if (
            rowId === participantPhone.trim() &&
            rowName.toLowerCase() === participantName.trim().toLowerCase()
          ) {
            rowIndexToDelete = i; // 0‑based, relative to A2
            break;
          }
        }

        if (rowIndexToDelete >= 0) {
          const sheetId = await getSheetId(sheets, spreadsheetId, "Players");
          const startRowIndex = rowIndexToDelete + 1; // header is row 1
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
          console.log(`🗑️ Removed ${participantName} (ID: ${participantPhone}) from Players`);
        } else {
          console.warn("⚠️ Could not find player in Players sheet to delete");
        }
      } catch (deleteError) {
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
  // Winners columns: A = Winner Name, B = Ticket Display ID, C = Date/Time, D = Status
  // ========================
  app.get("/api/raffle/winners", async (req, res) => {
    try {
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: "Winners!A2:D100",
      });

      const rows = response.data.values || [];
      const winners = rows
        .filter(row => row[0] || row[1])
        .map((row) => ({
          name: row[0] || "Unknown",           // A: Winner Name
          phone: row[1] || "",                  // B: Ticket Display ID (kept as phone for frontend)
          timestamp: row[2] || "",              // C: Date/Time
          status: row[3] || "WINNER",           // D: Status
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
  // Players columns: A = Ticket Display ID, B = Buyer's Name
  // ========================
  app.post("/api/raffle/add-participant", async (req, res) => {
    try {
      const { name, phone } = req.body;

      if (!name || !phone) {
        return res.status(400).json({
          success: false,
          message: "Name and Ticket Display ID are required",
        });
      }
      await sheets.spreadsheets.values.append({
        spreadsheetId,
        range: "Players!A:B",
        valueInputOption: "USER_ENTERED",
        requestBody: {
          values: [[phone, name]],   // A = Ticket Display ID, B = Buyer's Name
        },
      });

      console.log("✅ Participant added:", name, "| ID:", phone);

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