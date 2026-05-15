import express, { json } from "express";
import cors from "cors";
import { google } from "googleapis";
import { setupRaffleRoutes } from "./raffleRoutes.js";

const app = express();
app.use(cors());
app.use(json());

const PORT = 5000;

// GOOGLE AUTH
const auth = new google.auth.GoogleAuth({
  keyFile: "./service-account.json",
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

const authClient = await auth.getClient();

const sheets = google.sheets({
  version: "v4",
  auth: authClient,
});

process.on("uncaughtException", (err) => {
  console.error("UNCAUGHT EXCEPTION:", err);
});

process.on("unhandledRejection", (err) => {
  console.error("UNHANDLED PROMISE:", err);
});

// SHEET ID
const spreadsheetId = "1gXlGXkaSbPQx9y8X_hO4wh4MNMQldRu0Km8N9Ax1R8w";

// ===============================
// HELPERS
// ===============================
const normalizePhone = (phone = "") => String(phone).trim();
const updateSheet = async (rows) => {
  try {
    // convert objects → array format
    const values = [
      ["date", "prize", "remaining", "given"], // header
      ...rows.map((r) => [r.date, r.prize, r.remaining, r.given || 0]),
    ];

    // 🔥 overwrite entire sheet
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: "daily_distribution!A1",
      valueInputOption: "RAW",
      requestBody: {
        values,
      },
    });

    console.log("✅ Sheet updated successfully");
  } catch (err) {
    console.error("❌ updateSheet error:", err);
  }
};
const getRows = async (range) => {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range,
  });

  const rows = response.data.values || [];

  return rows.slice(1).map((row) => ({
    date: row[0],
    prize: Number(row[1]),
    remaining: Number(row[2]),
    given: Number(row[3]),
    unlock_time: row[4] ? row[4].trim() : "00:00", // ✅ keep as string
  }));
};

const getRawRows = async (range) => {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range,
  });
  const rows = response.data.values || [];
  return rows;
};
app.get("/", (req, res) => {
  res.send("Backend running");
});

app.get("/debug/auth", async (req, res) => {
  try {
    const response = await sheets.spreadsheets.get({
      spreadsheetId,
    });
    res.json({
      success: true,
      spreadsheetTitle: response.data.properties.title,
      spreadsheetId,
    });
  } catch (err) {
    console.error("❌ Auth Error:", err.message);
    res.status(500).json({
      success: false,
      error: err.message,
      details: err,
    });
  }
});

app.get("/debug/players", async (req, res) => {
  try {
    const rows = await getRawRows("Players!A2:G10");
    console.log("✅ Players sheet data:", rows);
    res.json({
      success: true,
      count: rows.length,
      sample: rows.slice(0, 3),
    });
  } catch (err) {
    console.error("❌ Error reading Players sheet:", err);
    res.status(500).json({
      success: false,
      error: err.message,
    });
  }
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

    const allRows = await getRawRows("Players!A2:G10000");
    const found = allRows.find((row) => row && row[1] && normalizePhone(row[1]) === phone);

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

    console.log("🔍 Login attempt:", { fullName, phone, language });

    if (!fullName || !phone) {
      return res.status(400).json({
        success: false,
        message: "Full name and phone are required",
      });
    }

    let allRows = [];
    try {
      allRows = await getRawRows("Players!A2:G10000");
      console.log("📋 Fetched rows from sheet. Count:", allRows.length);
    } catch (sheetError) {
      console.error("⚠️  Warning: Could not fetch existing rows:", sheetError.message);
      // Continue anyway - we'll create a new entry
      allRows = [];
    }

    const foundIndex = allRows.findIndex(
      (row) => row && row[1] && normalizePhone(row[1]) === phone,
    );

    if (foundIndex !== -1) {
      console.log("✅ Player found at index:", foundIndex);
      const sheetRow = foundIndex + 2;
      const existing = allRows[foundIndex];

      const updatedRow = [
        existing[0] || fullName,
        existing[1] || phone,
        existing[2] || language,
        existing[3] || "FALSE",
        existing[4] || 0,
        now,
        existing[6] || now,
      ];

      try {
        await sheets.spreadsheets.values.update({
          spreadsheetId,
          range: `Players!A${sheetRow}:G${sheetRow}`,
          valueInputOption: "USER_ENTERED",
          requestBody: {
            values: [updatedRow],
          },
        });
        console.log("✅ Player row updated");
      } catch (updateError) {
        console.error("⚠️ Could not update row:", updateError.message);
      }

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

    console.log("➕ Creating new player entry");
    try {
      await sheets.spreadsheets.values.append({
        spreadsheetId,
        range: "Players!A:G",
        valueInputOption: "USER_ENTERED",
        requestBody: {
          values: [[fullName, phone, language, "FALSE", 0, now, now]],
        },
      });
      console.log("✅ New player created");
    } catch (appendError) {
      console.error("⚠️  Could not append new player:", appendError.message);
    }

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
    console.error("Error stack:", error.stack);
    res.status(500).json({
      success: false,
      message: "Server error while checking or logging in player",
      error: error.message,
      errorCode: error.code,
      errorStatus: error.status,
    });
  }
});
const getToday = () => {
  const d = new Date();

  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");

  return `${year}-${month}-${day}`;
};
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

    const allRows = await getRawRows("Players!A2:G10000");
    const foundIndex = allRows.findIndex(
      (row) => row && row[1] && normalizePhone(row[1]) === phone,
    );

    if (foundIndex === -1) {
      return res.status(404).json({
        success: false,
        message: "Player not found",
      });
    }

    const sheetRow = foundIndex + 2;
    const existing = allRows[foundIndex];
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
const decreasePrize = async (selectedRow) => {
  const rows = await getRows("daily_distribution!A1:E");

  const index = rows.findIndex(
    (r) =>
      r.date === selectedRow.date &&
      r.prize === selectedRow.prize &&
      r.unlock_time === selectedRow.unlock_time &&
      r.remaining > 0
  );

  if (index === -1) {
    console.log("❌ Row not found for update");
    return;
  }

  rows[index].remaining = 0; // ✅ since 1 row = 1 prize
  rows[index].given = (rows[index].given || 0) + 1;

  console.log("🔻 Updated Row:", rows[index]);

  await updateSheet(rows);
};
app.get("/api/spin-results", async (req, res) => {
  try {
    const { date, phone } = req.query;

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: "spin_results!A1:D",
    });

    const rows = response.data.values || [];

    const data = rows.slice(1).map((row) => ({
      date: row[0],
      phone: row[1],
      prize: row[2],
      time: row[3],
    }));

    // 🔍 FILTER
    const filtered = data.filter((r) => {
      if (date && r.date !== date) return false;
      if (phone && r.phone !== phone) return false;
      return true;
    });

    res.json(filtered);
  } catch (err) {
    console.error("❌ GET spin results error:", err);
    res.status(500).json([]);
  }
});
app.get("/api/spin-results", async (req, res) => {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: "spin_results!A1:D",
    });

    const rows = response.data.values || [];

    if (rows.length <= 1) {
      return res.json([]);
    }

    // remove header
    const data = rows.slice(1).map((row) => ({
      date: row[0],
      phone: row[1],
      prize: row[2],
      time: row[3],
    }));

    res.json(data);
  } catch (err) {
    console.error("❌ GET spin results error:", err);
    res.status(500).json([]);
  }
});
const saveSpinResult = async ({ playerId, prize }) => {
  try {
    const now = new Date();

    const date = now.toLocaleDateString("en-CA", {
      timeZone: "Asia/Colombo",
    });

    const time = now.toLocaleTimeString("en-GB", {
      timeZone: "Asia/Colombo",
    });

    const values = [[date, playerId.phone, prize, time]];

    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: "spin_results!A2",
      valueInputOption: "RAW",
      requestBody: {
        values,
      },
    });

    console.log("📝 Spin result saved:", values);
  } catch (err) {
    console.error("❌ saveSpinResult error:", err);
  }
};
const getCurrentTime = () => {
  return new Date()
    .toLocaleTimeString("en-GB", {
      hour12: false,
      timeZone: "Asia/Colombo",
    })
    .slice(0, 5); // HH:mm
};

app.post("/spin", async (req, res) => {
  try {
    const today = getToday();
    const nowTime = getCurrentTime();
    const { playerId } = req.body;



    const rows = await getRows("daily_distribution!A1:E");

    const todayData = rows.filter((r) => r.date === today);

    if (!todayData.length) {
      console.log("❌ No data for today");
      return res.json({ success: true, prize: "No Chance" });
    }

    let pool = [];

    todayData.forEach((r) => {
      const isUnlocked = r.unlock_time <= nowTime;
      const hasStock = r.remaining > 0;

     
      if (isUnlocked && hasStock) {
        pool.push(r); // ✅ push full row
      }
    });

    console.log("🎯 Pool Size:", pool.length);

    if (pool.length === 0) {
      console.log("⚠️ No unlocked prizes → No Chance");

      await saveSpinResult({
        playerId,
        prize: "No Chance",
      });

      return res.json({
        success: true,
        prize: "No Chance",
      });
    }

    // 🎲 RANDOM PICK
    const selectedRow =
      pool[Math.floor(Math.random() * pool.length)];

    const selectedPrize = selectedRow.prize; // ✅ FIX

    console.log("🏆 Selected:", selectedPrize);

    // 🔻 UPDATE DISTRIBUTION
    await decreasePrize(selectedRow);

    // 📝 SAVE RESULT
    await saveSpinResult({
      playerId,
      prize: selectedPrize,
    });

    console.log("🎡 ===== SPIN END =====\n");

    return res.json({
      success: true,
      prize: selectedPrize,
    });
  } catch (err) {
    console.error("❌ SPIN ERROR:", err);

    return res.status(500).json({
      success: false,
      prize: "No Chance",
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

    const rows = await getRawRows("spin_inventory!A2:D1000");

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
      (r) => r.date === today && r.remaining > 0,
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
      days.filter((d) => !thousandDays.includes(d)),
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

app.use(cors());
app.use(json());

// ================= GLOBAL ERROR HANDLING =================
process.on("uncaughtException", (err) => console.error(err));
process.on("unhandledRejection", (err) => console.error(err));

const updateRow = async (range, values) => {
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range,
    valueInputOption: "USER_ENTERED",
    requestBody: { values },
  });
};

const appendRow = async (range, values) => {
  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range,
    valueInputOption: "USER_ENTERED",
    requestBody: { values },
  });
};

// 🇱🇰 Sri Lanka Date Fix
const getTodayDate = () => {
  const now = new Date();
  const sl = new Date(
    now.toLocaleString("en-US", { timeZone: "Asia/Colombo" }),
  );
  return sl.toISOString().split("T")[0];
};

// ================= BASIC =================
app.get("/", (req, res) => res.send("Backend running"));

// ================= SHEET =================
app.get("/sheet-data", async (req, res) => {
  try {
    const data = await getRawRows("Game1!A1:Z100");
    res.json({ success: true, data });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

app.post("/add-row", async (req, res) => {
  try {
    const { name, email, phone } = req.body;
    await appendRow("Game1!A:C", [[name, email, phone]]);
    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

app.put("/update", async (req, res) => {
  try {
    const { cell, value } = req.body;
    await updateRow(`Game1!${cell}`, [[value]]);
    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ success: false });
  }
});

app.delete("/delete-row", async (req, res) => {
  try {
    const { row } = req.body;
    await sheets.spreadsheets.values.clear({
      spreadsheetId,
      range: `Game1!A${row}:Z${row}`,
    });
    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ success: false });
  }
});

// ================= SPIN CONFIG =================
app.get("/api/spin-config", async (req, res) => {
  try {
    const rows = await getRawRows("Spin!A2:D1000");

    const data = rows.map((r) => ({
      option: r[0],
      weight: Number(r[1]),
      style: {
        backgroundColor: r[2] || "#faad14",
        textColor: r[3] || "#fff",
      },
    }));

    res.json({ success: true, data });
  } catch (e) {
    res.status(500).json({ success: false });
  }
});

app.post("/api/spin/play", async (req, res) => {
  try {
    const phone = normalizePhone(req.body.phone);

    const rows = await getRawRows("SpinConfig!A2:B100");

    let prizes = rows.map((r, i) => ({
      option: r[0],
      count: Number(r[1]),
      rowIndex: i + 2,
    }));

    // remove finished prizes
    prizes = prizes.filter((p) => p.count > 0);

    if (prizes.length === 0) {
      return res.json({
        success: true,
        prize: "😢 No Chance",
      });
    }

    // random pick
    const selected = prizes[Math.floor(Math.random() * prizes.length)];

    // decrement count
    await updateRow(`SpinConfig!B${selected.rowIndex}`, [[selected.count - 1]]);

    // save result
    await appendRow("SpinResults!A:C", [
      [phone, selected.option, new Date().toISOString()],
    ]);

    return res.json({
      success: true,
      prize: selected.option,
    });
  } catch (err) {
    res.status(500).json({ success: false });
  }
});

// ================= INIT =================
app.post("/api/spin/init", async (req, res) => {
  try {
    const days = Array.from({ length: 25 }, (_, i) => i + 1);
    const shuffle = (a) => a.sort(() => 0.5 - Math.random());

    const d1000 = shuffle([...days]).slice(0, 5);
    const d5000 = shuffle(days.filter((d) => !d1000.includes(d))).slice(0, 2);

    const rows = [];

    days.forEach((d) => {
      const date = new Date();
      date.setDate(date.getDate() + (d - 1));
      const dt = date.toISOString().split("T")[0];

      rows.push([dt, 250, 4, 4]);
      rows.push([dt, 500, 1, 1]);
      rows.push([
        dt,
        1000,
        d1000.includes(d) ? 1 : 0,
        d1000.includes(d) ? 1 : 0,
      ]);
      rows.push([
        dt,
        5000,
        d5000.includes(d) ? 1 : 0,
        d5000.includes(d) ? 1 : 0,
      ]);
    });

    await updateRow("spin_inventory!A2", rows);

    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ success: false });
  }
});

// ================= SAVE RESULT =================
app.post("/spin-result/submit", async (req, res) => {
  try {
    const { phone, fullName, prize } = req.body;

    await appendRow("SpinResults!A:G", [
      [
        phone,
        fullName,
        prize.option,
        prize.weight,
        prize.style?.backgroundColor,
        prize.style?.textColor,
        new Date().toISOString(),
      ],
    ]);

    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ success: false });
  }
});

// ================= SETUP RAFFLE ROUTES =================
setupRaffleRoutes(app, sheets, spreadsheetId, normalizePhone);

// ================= START =================
app.listen(PORT, () => {
  console.log(`🔥 Server running on ${PORT}`);
});
app.post("/api/setup-distribution", async (req, res) => {
  try {
    const TOTAL_DAYS = 25;

    const formatDate = (d) => d.toISOString().split("T")[0];

    const getRandomTime = () => {
      const hour = Math.floor(Math.random() * 24);
      const min = Math.floor(Math.random() * 60);
      return `${String(hour).padStart(2, "0")}:${String(min).padStart(2, "0")}`;
    };

    const getRandomDays = (count) => {
      const set = new Set();
      while (set.size < count) {
        set.add(Math.floor(Math.random() * TOTAL_DAYS));
      }
      return [...set];
    };

    const days1000 = getRandomDays(5);
    const days5000 = getRandomDays(2);

    const startDate = new Date();

    const rows = [];

    for (let i = 0; i < TOTAL_DAYS; i++) {
      const d = new Date(startDate);
      d.setDate(d.getDate() + i);
      const date = formatDate(d);

      // Rs.250 (4 per day → 4 random unlocks)
      for (let j = 0; j < 4; j++) {
        rows.push([date, 250, 1, 0, getRandomTime()]);
      }

      // Rs.500 (1 per day)
      rows.push([date, 500, 1, 0, getRandomTime()]);

      // Rs.1000
      rows.push([date, 1000, days1000.includes(i) ? 1 : 0, 0, getRandomTime()]);

      // Rs.5000
      rows.push([date, 5000, days5000.includes(i) ? 1 : 0, 0, getRandomTime()]);
    }

    // clear
    await sheets.spreadsheets.values.clear({
      spreadsheetId,
      range: "daily_distribution!A2:E",
    });

    // insert
    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: "daily_distribution!A2",
      valueInputOption: "RAW",
      requestBody: { values: rows },
    });

    return res.json({
      success: true,
      message: "🔥 Distribution with unlock_time created",
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false });
  }
});
