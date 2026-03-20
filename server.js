require("dotenv").config();

const express  = require("express");
const mongoose = require("mongoose");
const cors     = require("cors");
const ExcelJS  = require("exceljs");

const app = express();

app.use(cors({ origin: "*", methods: ["GET","POST"] }));
app.use(express.json({ limit: "20mb" }));

/* ── MONGODB ── */
let isConnected = false;
const connectDB = async () => {
  if (isConnected) return;
  await mongoose.connect(process.env.MONGO_URI);
  isConnected = true;
  console.log("✅ MongoDB connected");
};

/* ── SCHEMA ── */
const contactSchema = new mongoose.Schema({
  name:             { type: String, required: true, trim: true },
  email:            { type: String, required: true, trim: true, lowercase: true },
  subject:          { type: String, trim: true },
  message:          { type: String, required: true },
  resumeBase64:     { type: String, default: null },
  resumeFileName:   { type: String, default: null },
  resumeMimeType:   { type: String, default: null },
  createdAt:        { type: Date, default: Date.now },
});
const Contact = mongoose.models.Contact || mongoose.model("Contact", contactSchema);

/* ── ROUTES ── */
app.get("/", (req, res) => res.json({ status: "🟢 Ritesh Portfolio API running" }));

/* POST /api/contact */
app.post("/api/contact", async (req, res) => {
  try {
    await connectDB();
    const { name, email, subject, message, resumeBase64, resumeFileName, resumeMimeType } = req.body;
    if (!name || !email || !message)
      return res.status(400).json({ ok: false, error: "Name, email and message are required." });

    const doc = await Contact.create({
      name, email, subject, message,
      resumeBase64:   resumeBase64   || null,
      resumeFileName: resumeFileName || null,
      resumeMimeType: resumeMimeType || null,
    });
    console.log(`📩 New message from: ${name} <${email}>${resumeFileName ? " [Resume: "+resumeFileName+"]" : ""}`);
    res.status(201).json({ ok: true, id: doc._id });
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, error: "Server error: " + err.message });
  }
});

/* GET /api/messages */
app.get("/api/messages", async (req, res) => {
  try {
    await connectDB();
    const msgs = await Contact.find().sort({ createdAt: -1 })
      .select("-resumeBase64"); // don't send base64 in list
    res.json({ ok: true, count: msgs.length, data: msgs });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

/* GET /api/resume/:id — download resume by contact id */
app.get("/api/resume/:id", async (req, res) => {
  try {
    await connectDB();
    const doc = await Contact.findById(req.params.id).select("resumeBase64 resumeFileName resumeMimeType");
    if (!doc || !doc.resumeBase64)
      return res.status(404).json({ ok: false, error: "Resume not found" });

    const buffer = Buffer.from(doc.resumeBase64, "base64");
    res.setHeader("Content-Type", doc.resumeMimeType || "application/octet-stream");
    res.setHeader("Content-Disposition", `attachment; filename="${doc.resumeFileName || "resume.pdf"}"`);
    res.send(buffer);
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

/* GET /api/export — Excel */
app.get("/api/export", async (req, res) => {
  try {
    await connectDB();
    const msgs = await Contact.find().sort({ createdAt: -1 }).select("-resumeBase64");

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Contact Messages");

    ws.columns = [
      { header: "Sr.",     key: "sr",      width: 6  },
      { header: "Name",    key: "name",    width: 22 },
      { header: "Email",   key: "email",   width: 30 },
      { header: "Subject", key: "subject", width: 28 },
      { header: "Message", key: "message", width: 50 },
      { header: "Resume",  key: "resume",  width: 35 },
      { header: "Date",    key: "date",    width: 22 },
    ];

    /* Header styling */
    ws.getRow(1).eachCell(cell => {
      cell.fill      = { type:"pattern", pattern:"solid", fgColor:{ argb:"FF0D1B2A" } };
      cell.font      = { bold:true, color:{ argb:"FF00F5FF" }, size:11 };
      cell.border    = { bottom:{ style:"medium", color:{ argb:"FF00F5FF" } } };
      cell.alignment = { vertical:"middle", horizontal:"center" };
    });
    ws.getRow(1).height = 28;

    msgs.forEach((m, i) => {
      const resumeUrl = m.resumeFileName
        ? `${req.protocol}://${req.get("host")}/api/resume/${m._id}`
        : "No resume";

      const row = ws.addRow({
        sr:      i + 1,
        name:    m.name,
        email:   m.email,
        subject: m.subject || "—",
        message: m.message,
        resume:  resumeUrl,
        date:    new Date(m.createdAt).toLocaleString("en-IN", { timeZone:"Asia/Kolkata" }),
      });

      /* Clickable resume link */
      if (m.resumeFileName) {
        row.getCell("resume").value = {
          text: m.resumeFileName,
          hyperlink: resumeUrl,
        };
        row.getCell("resume").font = { color:{ argb:"FF00F5FF" }, underline:true, size:10 };
      }

      const bg = i % 2 === 0 ? "FF0A1628" : "FF0D1F3C";
      row.eachCell({ includeEmpty: true }, (cell, colNum) => {
        cell.fill      = { type:"pattern", pattern:"solid", fgColor:{ argb: bg } };
        if (colNum !== 6 || !m.resumeFileName)
          cell.font    = { color:{ argb:"FFE8EAF0" }, size:10 };
        cell.alignment = { vertical:"middle", wrapText: true };
        cell.border    = { bottom:{ style:"thin", color:{ argb:"FF1A2E4A" } } };
      });
      row.height = 22;
    });

    ws.views      = [{ state:"frozen", ySplit:1 }];
    ws.autoFilter = { from:"A1", to:"G1" };

    res.setHeader("Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition",
      `attachment; filename="ritesh_contacts_${Date.now()}.xlsx"`);
    await wb.xlsx.write(res);
    res.end();
    console.log(`📊 Excel exported — ${msgs.length} records`);
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, error: err.message });
  }
});

module.exports = app;

if (require.main === module) {
  const PORT = process.env.PORT || 5000;
  app.listen(PORT, () => console.log(`🚀 Server running on http://localhost:${PORT}`));
}