require("dotenv").config();

const express   = require("express");
const mongoose  = require("mongoose");
const cors      = require("cors");
const multer    = require("multer");
const path      = require("path");
const fs        = require("fs");
const ExcelJS   = require("exceljs");

const app = express();

app.use(cors({ origin: "*", methods: ["GET","POST"] }));
app.use(express.json());

/* ── UPLOADS FOLDER ── */
const uploadDir = path.join(__dirname, "uploads");
if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir);

/* ── MULTER ── */
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, uploadDir),
  filename:    (req, file, cb) => {
    const unique = Date.now() + "-" + Math.round(Math.random() * 1e6);
    cb(null, unique + path.extname(file.originalname));
  },
});
const upload = multer({
  storage,
  limits: { fileSize: 5 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const allowed = [".pdf",".doc",".docx"];
    const ext = path.extname(file.originalname).toLowerCase();
    allowed.includes(ext) ? cb(null, true) : cb(new Error("Only PDF/DOC/DOCX allowed"));
  },
});

/* ── MONGODB ── */
mongoose.connect(process.env.MONGO_URI)
  .then(() => console.log("✅ MongoDB Atlas connected"))
  .catch((err) => console.error("❌ MongoDB error:", err));

/* ── SCHEMA ── */
const contactSchema = new mongoose.Schema({
  name:               { type: String, required: true, trim: true },
  email:              { type: String, required: true, trim: true, lowercase: true },
  subject:            { type: String, trim: true },
  message:            { type: String, required: true },
  resumeFile:         { type: String, default: null },
  resumeOriginalName: { type: String, default: null },
  createdAt:          { type: Date, default: Date.now },
  read:               { type: Boolean, default: false },
});
const Contact = mongoose.model("Contact", contactSchema);

/* ── ROUTES ── */
app.get("/", (req, res) => res.json({ status: "🟢 Ritesh Portfolio API running" }));

/* POST /api/contact */
app.post("/api/contact", upload.single("resume"), async (req, res) => {
  try {
    const { name, email, subject, message } = req.body;
    if (!name || !email || !message)
      return res.status(400).json({ ok: false, error: "Name, email and message are required." });

    const doc = await Contact.create({
      name, email, subject, message,
      resumeFile:         req.file ? req.file.filename     : null,
      resumeOriginalName: req.file ? req.file.originalname : null,
    });
    console.log(`📩 New message from: ${name} <${email}>${req.file ? " [Resume attached]" : ""}`);
    res.status(201).json({ ok: true, id: doc._id });
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, error: "Server error." });
  }
});

/* GET /api/messages */
app.get("/api/messages", async (req, res) => {
  try {
    const msgs = await Contact.find().sort({ createdAt: -1 });
    res.json({ ok: true, count: msgs.length, data: msgs });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

/* GET /api/resume/:filename */
app.get("/api/resume/:filename", (req, res) => {
  const filePath = path.join(uploadDir, req.params.filename);
  if (!fs.existsSync(filePath))
    return res.status(404).json({ ok: false, error: "File not found" });
  res.download(filePath);
});

/* GET /api/export — Excel download */
app.get("/api/export", async (req, res) => {
  try {
    const msgs = await Contact.find().sort({ createdAt: -1 });

    const wb = new ExcelJS.Workbook();
    wb.creator = "Ritesh Portfolio";
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

    /* Data rows */
    msgs.forEach((m, i) => {
      const resumeUrl = m.resumeFile
        ? `${req.protocol}://${req.get("host")}/api/resume/${m.resumeFile}`
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
      if (m.resumeFile) {
        row.getCell("resume").value = {
          text: m.resumeOriginalName || "Download Resume",
          hyperlink: resumeUrl,
        };
        row.getCell("resume").font = { color:{ argb:"FF00F5FF" }, underline:true, size:10 };
      }

      const bg = i % 2 === 0 ? "FF0A1628" : "FF0D1F3C";
      row.eachCell({ includeEmpty: true }, (cell, colNum) => {
        cell.fill      = { type:"pattern", pattern:"solid", fgColor:{ argb: bg } };
        if (colNum !== 6 || !m.resumeFile)
          cell.font    = { color:{ argb:"FFE8EAF0" }, size:10 };
        cell.alignment = { vertical:"middle", wrapText: true };
        cell.border    = { bottom:{ style:"thin", color:{ argb:"FF1A2E4A" } } };
      });
      row.height = 22;
    });

    ws.views     = [{ state:"frozen", ySplit:1 }];
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

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => console.log(`🚀 Server running on http://localhost:${PORT}`));