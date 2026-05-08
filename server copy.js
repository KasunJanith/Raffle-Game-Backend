const express = require("express");
const cors = require("cors");
const { google } = require("googleapis");

const app = express();
app.use(cors());
app.use(express.json());

const PORT = 5000;

// GOOGLE AUTH
const auth = new google.auth.GoogleAuth({
  keyFile: "./service-account.json",
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

const sheets = google.sheets({ version: "v4", auth });

process.on("uncaughtException", (err) => {
  console.error("UNCAUGHT EXCEPTION:", err);
});

process.on("unhandledRejection", (err) => {
  console.error("UNHANDLED PROMISE:", err);
});

// SHEET ID
const spreadsheetId = "1a9ZmrJFkQr5O9z4oMXrJYUYEKJwNi_bm7ennxUzy_Z8";

// ===============================
// HELPERS
// ===============================
const normalizePhone = (phone = "") => String(phone).trim();

const getRows = async (range) => {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range,
  });
  return response.data.values || [];
};

app.get("/", (req, res) => {
  res.send("Backend running");
});

// ===============================
// READ GOOGLE SHEET
// ===============================
app.get("/sheet-data", async (req, res) => {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: "Game1!A1:Z100",
    });

    res.json({
      success: true,
      data: response.data.values,
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message,
    });
  }
});

// ===============================
// ADD ROW
// ===============================
app.post("/add-row", async (req, res) => {
  try {
    const { name, email, phone } = req.body;

    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: "Game1!A:C",
      valueInputOption: "RAW",
      requestBody: {
        values: [[name, email, phone]],
      },
    });

    res.json({
      success: true,
      message: "Row added",
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message,
    });
  }
});

// ===============================
// UPDATE CELL
// ===============================
app.put("/update", async (req, res) => {
  try {
    const { cell, value } = req.body;

    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `Game1!${cell}`,
      valueInputOption: "RAW",
      requestBody: {
        values: [[value]],
      },
    });

    res.json({
      success: true,
      message: "Cell updated",
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message,
    });
  }
});

// ===============================
// DELETE ROW
// ===============================
app.delete("/delete-row", async (req, res) => {
  try {
    const { row } = req.body;

    await sheets.spreadsheets.values.clear({
      spreadsheetId,
      range: `Game1!A${row}:Z${row}`,
    });

    res.json({
      success: true,
      message: "Row deleted",
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message,
    });
  }
});

// ===============================
// CHECK REGISTERED USER FROM USERS SHEET
// ===============================
app.get("/check-user/:mobile", async (req, res) => {
  try {
    let mobile = req.params.mobile;

    if (mobile.startsWith("0")) {
      mobile = "94" + mobile.substring(1);
    }

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: "Users!A1:L1000",
    });

    const rows = response.data.values || [];
    const users = rows.slice(1);

    const user = users.find((row) => row[6] === mobile);

    if (user) {
      return res.json({
        success: true,
        registered: true,
        firstName: user[0],
        lastName: user[1],
        mobile: user[6],
      });
    }

    res.json({
      success: true,
      registered: false,
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message,
    });
  }
});

// ===============================
// PLAYER CHECK
// ===============================
app.post("/api/player/check", async (req, res) => {
  try {
    const phone = normalizePhone(req.body.phone);

    if (!phone) {
      return res.status(400).json({
        success: false,
        message: "Phone number is required",
      });
    }

    const rows = await getRows("Players!A2:G10000");
    const found = rows.find((row) => normalizePhone(row[1]) === phone);

    if (!found) {
      return res.json({
        success: true,
        exists: false,
        playedBefore: false,
        player: null,
      });
    }

    return res.json({
      success: true,
      exists: true,
      playedBefore: String(found[3]).toLowerCase() === "true",
      player: {
        fullName: found[0] || "",
        phone: found[1] || "",
        language: found[2] || "en",
        playedBefore: String(found[3]).toLowerCase() === "true",
        gamePlayedCount: Number(found[4] || 0),
        lastLoginAt: found[5] || null,
        createdAt: found[6] || null,
      },
    });
  } catch (error) {
    console.error("PLAYER CHECK ERROR:", error);
    res.status(500).json({
      success: false,
      message: "Server error while checking player",
      error: error.message,
    });
  }
});

// ===============================
// PLAYER CHECK OR LOGIN
// ===============================
app.post("/api/player/check-or-login", async (req, res) => {
  try {
    const fullName = String(req.body.fullName || "").trim();
    const phone = normalizePhone(req.body.phone);
    const language = String(req.body.language || "en").trim();
    const now = new Date().toISOString();

    if (!fullName || !phone) {
      return res.status(400).json({
        success: false,
        message: "Full name and phone are required",
      });
    }

    const rows = await getRows("Players!A2:G10000");
    const foundIndex = rows.findIndex(
      (row) => normalizePhone(row[1]) === phone,
    );

    if (foundIndex !== -1) {
      const sheetRow = foundIndex + 2;
      const existing = rows[foundIndex];

      const updatedRow = [
        existing[0] || fullName,
        existing[1] || phone,
        existing[2] || language,
        existing[3] || "FALSE",
        existing[4] || 0,
        now,
        existing[6] || now,
      ];

      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `Players!A${sheetRow}:G${sheetRow}`,
        valueInputOption: "USER_ENTERED",
        requestBody: {
          values: [updatedRow],
        },
      });

      return res.json({
        success: true,
        exists: true,
        isNewUser: false,
        playedBefore: String(updatedRow[3]).toLowerCase() === "true",
        message: "Existing player logged in",
        player: {
          fullName: updatedRow[0],
          phone: updatedRow[1],
          language: updatedRow[2],
          playedBefore: String(updatedRow[3]).toLowerCase() === "true",
          gamePlayedCount: Number(updatedRow[4] || 0),
          lastLoginAt: updatedRow[5],
          createdAt: updatedRow[6],
        },
      });
    }

    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: "Players!A:G",
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [[fullName, phone, language, "FALSE", 0, now, now]],
      },
    });

    return res.status(201).json({
      success: true,
      exists: false,
      isNewUser: true,
      playedBefore: false,
      message: "New player created and logged in",
      player: {
        fullName,
        phone,
        language,
        playedBefore: false,
        gamePlayedCount: 0,
        lastLoginAt: now,
        createdAt: now,
      },
    });
  } catch (error) {
    console.error("CHECK OR LOGIN ERROR:", error);
    res.status(500).json({
      success: false,
      message: "Server error while checking or logging in player",
      error: error.message,
    });
  }
});

// ===============================
// MARK PLAYER AS PLAYED
// ===============================
app.post("/api/player/mark-played", async (req, res) => {
  try {
    const phone = normalizePhone(req.body.phone);

    if (!phone) {
      return res.status(400).json({
        success: false,
        message: "Phone number is required",
      });
    }

    const rows = await getRows("Players!A2:G10000");
    const foundIndex = rows.findIndex(
      (row) => normalizePhone(row[1]) === phone,
    );

    if (foundIndex === -1) {
      return res.status(404).json({
        success: false,
        message: "Player not found",
      });
    }

    const sheetRow = foundIndex + 2;
    const existing = rows[foundIndex];
    const currentCount = Number(existing[4] || 0);

    const updatedRow = [
      existing[0] || "",
      existing[1] || "",
      existing[2] || "en",
      "TRUE",
      currentCount + 1,
      new Date().toISOString(),
      existing[6] || new Date().toISOString(),
    ];

    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `Players!A${sheetRow}:G${sheetRow}`,
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [updatedRow],
      },
    });

    res.json({
      success: true,
      message: "Player marked as played",
      player: {
        fullName: updatedRow[0],
        phone: updatedRow[1],
        language: updatedRow[2],
        playedBefore: true,
        gamePlayedCount: Number(updatedRow[4]),
        lastLoginAt: updatedRow[5],
        createdAt: updatedRow[6],
      },
    });
  } catch (error) {
    console.error("MARK PLAYED ERROR:", error);
    res.status(500).json({
      success: false,
      message: "Server error while marking player as played",
      error: error.message,
    });
  }
});

// ===============================
// GAME 1 SUBMIT - QUIZ
// Sheet: QuizGame
// Header: phone | fullName | language | score | points | answer | playedAt
// ===============================
app.post("/game1/submit", async (req, res) => {
  try {
    const { mobile, phone, name, fullName, language, score, time, answer } =
      req.body;

    const finalPhone = normalizePhone(phone || mobile);
    const finalName = String(fullName || name || "").trim();
    const finalLanguage = String(language || "en").trim();

    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: "QuizGame!A:D",
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [[finalPhone, finalName, score * 10, time]],
      },
    });

    res.json({
      success: true,
      message: "Game1 answer saved ******",
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message,
    });
  }
});

// ===============================
// GAME 2 SUBMIT - KAVUM COUNT
// Sheet: KavumCount
// ===============================

app.post("/game2/submit", async (req, res) => {
  try {
    const { mobile, phone, name, fullName, time, score, isCorrect, answer } =
      req.body;

    const finalPhone = normalizePhone(phone || mobile);
    const finalName = String(fullName || name || "").trim();

    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: "KavumCount!A:D",
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [[finalPhone, finalName, isCorrect, time]],
      },
    });

    res.json({
      success: true,
      message: "Game2 answer saved",
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message,
    });
  }
});

// ===============================
// GAME 3 SUBMIT - HIDDEN LAMPS
// Sheet: HiddenLamps
// ===============================
app.post("/game3/submit", async (req, res) => {
  try {
    const { mobile, phone, name, fullName, time, score, isCorrect, answer } =
      req.body;

    const finalPhone = normalizePhone(phone || mobile);
    const finalName = String(fullName || name || "").trim();

    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: "HiddenLamps!A:G",
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [[finalPhone, finalName, isCorrect, time]],
      },
    });

    res.json({
      success: true,
      message: "Game3 answer saved",
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message,
    });
  }
});

// ===============================
// GAME 4 SUBMIT - RABANA
// Sheet: RabanaGame
// ===============================
app.post("/game4/submit", async (req, res) => {
  try {
    const { mobile, phone, name, fullName, score } = req.body;

    const finalPhone = normalizePhone(phone || mobile);
    const finalName = String(fullName || name || "").trim();

    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: "RabanaGame!A:D",
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [[finalPhone, finalName, score]],
      },
    });

    res.json({
      success: true,
      message: "Game4 answer saved",
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message,
    });
  }
});

// ===============================
// GAME 5 SUBMIT - CATCH KAVUM
// Sheet: CatchKavum
// ===============================
app.post("/game5/submit", async (req, res) => {
  try {
    const { mobile, phone, name, fullName, score } = req.body;

    const finalPhone = normalizePhone(phone || mobile);
    const finalName = String(fullName || name || "").trim();

    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: "CatchKavum!A:C",
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [[finalPhone, finalName, score]],
      },
    });

    res.json({
      success: true,
      message: "Game5 answer saved",
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message,
    });
  }
});

// ===============================
// GAME 6 SUBMIT - BREAK POT
// Sheet: BreakPot
// ===============================
app.post("/game6/submit", async (req, res) => {
  try {
    const { mobile, phone, name, fullName, time } = req.body;

    const finalPhone = normalizePhone(phone || mobile);
    const finalName = String(fullName || name || "").trim();

    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: "BreakPot!A:C",
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [[finalPhone, finalName, time]],
      },
    });

    res.json({
      success: true,
      message: "Game6 answer saved",
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message,
    });
  }
});

// ===============================
// CHECK IF USER ALREADY PLAYED A GAME
// game1 -> QuizGame
// game2 -> KavumCount
// game3 -> HiddenLamps
// game4 -> RabanaGame
// game5 -> CatchKavum
// game6 -> BreakPot
// ===============================
app.get("/check-played/:game/:mobile", async (req, res) => {
  try {
    const { game, mobile } = req.params;

    const gameMap = {
      game1: "QuizGame",
      game2: "KavumCount",
      game3: "HiddenLamps",
      game4: "RabanaGame",
      game5: "CatchKavum",
      game6: "BreakPot",
    };

    const sheetName = gameMap[game];

    if (!sheetName) {
      return res.status(400).json({
        success: false,
        message: "Invalid game name",
      });
    }

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetName}!A2:G10000`,
    });

    const rows = response.data.values || [];
    const found = rows.find(
      (row) => normalizePhone(row[0]) === normalizePhone(mobile),
    );

    res.json({
      success: true,
      played: !!found,
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message,
    });
  }
});

// ===============================
// LEADERBOARD FROM QUIZGAME
// ===============================
app.get("/leaderboard", async (req, res) => {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: "QuizGame!A2:G10000",
    });

    const rows = response.data.values || [];

    const players = rows.map((row) => ({
      phone: row[0] || "",
      fullName: row[1] || "",
      language: row[2] || "",
      score: row[3] || "",
      points: row[4] || "",
      answer: row[5] || "",
      playedAt: row[6] || "",
    }));

    res.json({
      success: true,
      data: players,
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message,
    });
  }
});


app.get("/api/spin-config", async (req, res) => {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: "Spin!A2:D1000",
    });

    const rows = response.data.values || [];

    const data = rows
      .filter((row) => row[0] && row[1]) // must have option + weight
      .map((row) => ({
        option: row[0] || "",
        weight: Number(row[1]) || 0,
        style: {
          backgroundColor: row[2] || "#1890ff",
          textColor: row[3] || "#ffffff",
        },
      }));

    res.json({
      success: true,
      data,
    });
  } catch (error) {
    console.error("SPIN CONFIG ERROR:", error);
    res.status(500).json({
      success: false,
      message: "Failed to fetch spin config",
      error: error.message,
    });
  }
});

app.post("/spin-result/submit", async (req, res) => {
  try {
    const { mobile, phone, name, fullName, prize } = req.body;

    const finalPhone = normalizePhone(phone || mobile);
    const finalName = String(fullName || name || "").trim();

    if (!prize) {
      return res.status(400).json({
        success: false,
        message: "Prize is required",
      });
    }

    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: "SpinResults!A:G",
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [
          [
            finalPhone,
            finalName,
            prize.option || "",
            Number(prize.weight) || 0,
            prize.style?.backgroundColor || "",
            prize.style?.textColor || "",
            new Date().toISOString(),
          ],
        ],
      },
    });

    res.json({
      success: true,
      message: "Spin result saved",
    });
  } catch (error) {
    console.error("SPIN RESULT ERROR:", error);
    res.status(500).json({
      success: false,
      error: error.message,
    });
  }
});


// ✅ SAVE GAME RESULT
app.post("/save-game-summary", async (req, res) => {
  try {
    const { name, phone, games } = req.body;

    if (!name || !phone || !games) {
      return res.json({ success: false, message: "Missing data" });
    }

    // 👉 Flatten data
    const row = [
      new Date().toISOString(),
      name,
      phone,

      // quiz
      games.quiz?.score || "",
      games.quiz?.time || "",

      // kavum
      games.kavum_count?.score || "",
      games.kavum_count?.time || "",

      // lamps
      games.hidden_lamps?.score || "",
      games.hidden_lamps?.time || "",

      // rabana
      games.rabana?.score || "",

      // catch kavum
      games.catch_kavum?.score || "",

      // break pot
      games.break_pot?.finalTime || "",
    ];

    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: "SheetFinal!A1",
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [row],
      },
    });

    res.json({ success: true });
  } catch (err) {
    console.error(err);
    res.json({ success: false, message: "Server error" });
  }
});

// Format today date (YYYY-MM-DD)
const getTodayDate = () => {
  return new Date().toISOString().split("T")[0];
};

// Read sheet

// Update remaining
const updateRemaining = async (rowIndex, newValue) => {
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `spin_inventory!D${rowIndex}`,
    valueInputOption: "RAW",
    requestBody: {
      values: [[newValue]],
    },
  });
};

// ================= PLAY SPIN =================

app.post("/api/spin/play", async (req, res) => {
  try {
    const today = getTodayDate();

    const rows = await getRows();

    // map rows with index
    const formatted = rows.map((row, i) => ({
      date: row[0],
      prize: row[1],
      total: Number(row[2]),
      remaining: Number(row[3]),
      rowIndex: i + 2, // because sheet starts from row 2
    }));

    // filter today's available prizes
    const todayAvailable = formatted.filter(
      (r) => r.date === today && r.remaining > 0
    );

    if (todayAvailable.length === 0) {
      return res.json({
        success: true,
        prize: "😢 No Chance",
      });
    }

    // RANDOM PICK
    const selected =
      todayAvailable[Math.floor(Math.random() * todayAvailable.length)];

    // DECREMENT
    await updateRemaining(selected.rowIndex, selected.remaining - 1);

    return res.json({
      success: true,
      prize: `💰 Rs.${selected.prize}`,
      rawPrize: selected.prize,
    });
  } catch (error) {
    console.error("Spin error:", error);
    return res.status(500).json({
      success: false,
      message: "Spin failed",
    });
  }
});

// ================= INIT DATA (RUN ONCE) =================

app.post("/api/spin/init", async (req, res) => {
  try {
    const days = Array.from({ length: 25 }, (_, i) => i + 1);

    const shuffle = (arr) => arr.sort(() => 0.5 - Math.random());

    const thousandDays = shuffle([...days]).slice(0, 5);
    const fiveThousandDays = shuffle(
      days.filter((d) => !thousandDays.includes(d))
    ).slice(0, 2);

    const rows = [];

    days.forEach((day) => {
      const date = new Date();
      date.setDate(date.getDate() + (day - 1));
      const formattedDate = date.toISOString().split("T")[0];

      // always
      rows.push([formattedDate, 250, 4, 4]);
      rows.push([formattedDate, 500, 1, 1]);

      // special prizes
      if (thousandDays.includes(day)) {
        rows.push([formattedDate, 1000, 1, 1]);
      } else {
        rows.push([formattedDate, 1000, 0, 0]);
      }

      if (fiveThousandDays.includes(day)) {
        rows.push([formattedDate, 5000, 1, 1]);
      } else {
        rows.push([formattedDate, 5000, 0, 0]);
      }
    });

    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: "spin_inventory!A2",
      valueInputOption: "RAW",
      requestBody: {
        values: rows,
      },
    });

    return res.json({
      success: true,
      message: "Spin inventory initialized",
    });
  } catch (error) {
    console.error("Init error:", error);
    return res.status(500).json({
      success: false,
      message: "Init failed",
    });
  }
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
