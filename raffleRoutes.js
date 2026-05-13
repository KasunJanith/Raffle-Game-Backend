// Raffle Endpoints for the new raffle draw system
import express from "express";

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
  // SAVE WINNER
  // ========================
  app.post("/api/raffle/save-winner", async (req, res) => {
    try {
      const { participantId, participantName, participantPhone, timestamp } =
        req.body;

      if (!participantName || !participantPhone) {
        return res.status(400).json({
          success: false,
          message: "Participant name and phone are required",
        });
      }

      // Append winner to Winners sheet
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
      }      await sheets.spreadsheets.values.append({
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

