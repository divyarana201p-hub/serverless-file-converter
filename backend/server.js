
require("dotenv").config({
    path: require("path").resolve(__dirname, "../.env")
});

const express = require("express");
const cors = require("cors");
const multer = require("multer");
const path = require("path");
const fs = require("fs");
const crypto = require("crypto");
const XLSX = require("xlsx");
const PDFDocument = require("pdfkit");
const mammoth = require("mammoth");
const { Document, Packer, Paragraph } = require("docx");
const { PDFParse } = require("pdf-parse");

const {
    S3Client,
    PutObjectCommand,
    ListObjectsV2Command,
    GetObjectCommand,
    DeleteObjectCommand
} = require("@aws-sdk/client-s3");

const app = express();
const PORT = process.env.PORT || 5000;
const UPLOAD_DIR = path.join(__dirname, "uploads");

fs.mkdirSync(UPLOAD_DIR, { recursive: true });

app.use(cors());
app.use(express.json());

const BUCKET = process.env.FILEBASE_BUCKET;

const s3 = new S3Client({
    endpoint: process.env.FILEBASE_ENDPOINT || "https://s3.filebase.io",
    region: process.env.FILEBASE_REGION || "us-east-1",
    forcePathStyle: true,
    credentials: {
        accessKeyId: process.env.FILEBASE_ACCESS_KEY,
        secretAccessKey: process.env.FILEBASE_SECRET_KEY
    }
});

const storage = multer.diskStorage({
    destination: (_req, _file, cb) => cb(null, UPLOAD_DIR),
    filename: (_req, file, cb) => {
        const ext = path.extname(file.originalname);
        cb(null, `${crypto.randomUUID()}${ext}`);
    }
});

const upload = multer({
    storage,
    limits: { fileSize: 20 * 1024 * 1024 }
});

function safeName(name) {
    return path.basename(name).replace(/[^a-zA-Z0-9._-]/g, "_");
}

function makePdf(lines, outputPath) {
    return new Promise((resolve, reject) => {
        const pdf = new PDFDocument({ margin: 40 });
        const stream = fs.createWriteStream(outputPath);
        stream.on("error", reject);
        pdf.pipe(stream);
        lines.forEach(line => pdf.text(String(line ?? ""), {
            paragraphGap: 5
        }));
        pdf.end();
        stream.on("finish", resolve);
    });
}

async function uploadToCloud(filePath, key) {
    if (!BUCKET) throw new Error("FILEBASE_BUCKET is missing in .env");

    await s3.send(new PutObjectCommand({
        Bucket: BUCKET,
        Key: key,
        Body: fs.createReadStream(filePath)
    }));
}

async function saveAndSend(req, res, outputPath, outputName, contentType) {
    const key = `${Date.now()}-${safeName(outputName)}`;
    await uploadToCloud(outputPath, key);

    res.json({
        message: "Conversion successful and saved to Filebase.",
        filename: key,
        downloadUrl: `/download/${encodeURIComponent(key)}`,
        contentType
    });
}

app.get("/", (_req, res) => {
    res.send("Serverless File Converter backend is running.");
});

app.post("/upload", upload.single("file"), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ message: "Please select a file." });
        }

        const key = `${Date.now()}-${safeName(req.file.originalname)}`;
        await uploadToCloud(req.file.path, key);

        res.json({
            message: "File uploaded to Filebase successfully.",
            filename: key
        });
    } catch (error) {
        console.error("Upload error:", error.name, error.message);
        res.status(500).json({ message: "Upload failed.", error: error.message });
    }
});

app.get("/files", async (_req, res) => {
    try {
        if (!BUCKET) throw new Error("FILEBASE_BUCKET is missing in .env");

        const result = await s3.send(new ListObjectsV2Command({
            Bucket: BUCKET
        }));

        res.json((result.Contents || []).map(item => ({
            name: item.Key,
            size: item.Size,
            storage: "Filebase",
            downloadUrl: `/download/${encodeURIComponent(item.Key)}`
        })));
    } catch (error) {
        console.error("Filebase list error:", error.name, error.message);
        res.status(500).json({
            message: "Could not retrieve cloud files.",
            error: error.message
        });
    }
});

app.get("/download/:key", async (req, res) => {
    try {
        const key = req.params.key;

        const result = await s3.send(new GetObjectCommand({
            Bucket: BUCKET,
            Key: key
        }));

        res.setHeader(
            "Content-Disposition",
            `attachment; filename="${safeName(key)}"`
        );

        result.Body.pipe(res);
    } catch (error) {
        console.error("Download error:", error.name, error.message);
        res.status(500).json({ message: "Download failed." });
    }
});

app.post("/convert/excel-to-pdf", upload.single("file"), async (req, res) => {
    let outputPath;

    try {
        if (!req.file) {
            return res.status(400).json({ message: "Select an Excel file." });
        }

        const workbook = XLSX.readFile(req.file.path);
        const lines = [];

        workbook.SheetNames.forEach(sheetName => {
            lines.push(`Sheet: ${sheetName}`);
            const rows = XLSX.utils.sheet_to_json(
                workbook.Sheets[sheetName],
                { header: 1, defval: "" }
            );

            rows.forEach(row => {
                lines.push(row.map(cell => String(cell)).join(" | "));
            });
            lines.push("");
        });

        outputPath = path.join(UPLOAD_DIR, `${crypto.randomUUID()}.pdf`);
        await makePdf(lines, outputPath);

        await saveAndSend(
            req, res, outputPath,
            `${path.parse(req.file.originalname).name}.pdf`,
            "application/pdf"
        );
    } catch (error) {
        console.error("Excel-to-PDF error:", error);
        res.status(500).json({
            message: "Excel to PDF conversion failed.",
            error: error.message
        });
    }
});

app.post("/convert/pdf-to-excel", upload.single("file"), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ message: "Select a PDF file." });
        }

        const buffer = fs.readFileSync(req.file.path);

const parser = new PDFParse({ data: buffer });
const result = await parser.getText();
const parsed = { text: result.text };

await parser.destroy();
        const rows = parsed.text.split(/\r?\n/)
            .map(line => line.trim())
            .filter(Boolean)
            .map(line => line.split(/\s{2,}|\t|\s+\|\s+/));

        const sheet = XLSX.utils.aoa_to_sheet(rows);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, sheet, "Converted");

        const outputPath = path.join(UPLOAD_DIR, `${crypto.randomUUID()}.xlsx`);
        XLSX.writeFile(workbook, outputPath);

        await saveAndSend(
            req, res, outputPath,
            `${path.parse(req.file.originalname).name}.xlsx`,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        );
    } catch (error) {
        console.error("PDF-to-Excel error:", error);
        res.status(500).json({
            message: "PDF to Excel conversion failed.",
            error: error.message
        });
    }
});

app.post("/convert/word-to-pdf", upload.single("file"), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ message: "Select a Word file." });
        }

        const result = await mammoth.extractRawText({ path: req.file.path });
        const outputPath = path.join(UPLOAD_DIR, `${crypto.randomUUID()}.pdf`);

        await makePdf(result.value.split(/\r?\n/), outputPath);

        await saveAndSend(
            req, res, outputPath,
            `${path.parse(req.file.originalname).name}.pdf`,
            "application/pdf"
        );
    } catch (error) {
        console.error("Word-to-PDF error:", error);
        res.status(500).json({
            message: "Word to PDF conversion failed.",
            error: error.message
        });
    }
});

app.post("/convert/pdf-to-word", upload.single("file"), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ message: "Select a PDF file." });
        }

        const parser = new PDFParse({
    data: fs.readFileSync(req.file.path)
});

const result = await parser.getText();
const parsed = { text: result.text };

await parser.destroy();
        const doc = new Document({
            sections: [{
                children: parsed.text.split(/\r?\n/).map(
                    line => new Paragraph({ text: line })
                )
            }]
        });

        const buffer = await Packer.toBuffer(doc);
        const outputPath = path.join(UPLOAD_DIR, `${crypto.randomUUID()}.docx`);
        fs.writeFileSync(outputPath, buffer);

        await saveAndSend(
            req, res, outputPath,
            `${path.parse(req.file.originalname).name}.docx`,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        );
    } catch (error) {
        console.error("PDF-to-Word error:", error);
        res.status(500).json({
            message: "PDF to Word conversion failed.",
            error: error.message
        });
    }
});

app.post("/convert/image-to-pdf", upload.single("file"), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ message: "Select a JPG or PNG image." });
        }

        const ext = path.extname(req.file.originalname).toLowerCase();
        if (![".jpg", ".jpeg", ".png"].includes(ext)) {
            return res.status(400).json({
                message: "Only JPG, JPEG, and PNG images are supported."
            });
        }

        const outputPath = path.join(UPLOAD_DIR, `${crypto.randomUUID()}.pdf`);

        await new Promise((resolve, reject) => {
            const pdf = new PDFDocument({ autoFirstPage: false });
            const stream = fs.createWriteStream(outputPath);
            stream.on("error", reject);
            pdf.pipe(stream);

            pdf.addPage({ size: "A4", margin: 30 });
            pdf.image(req.file.path, 30, 30, {
                fit: [535, 780],
                align: "center",
                valign: "center"
            });

            pdf.end();
            stream.on("finish", resolve);
        });

        await saveAndSend(
            req, res, outputPath,
            `${path.parse(req.file.originalname).name}.pdf`,
            "application/pdf"
        );
    } catch (error) {
        console.error("Image-to-PDF error:", error);
        res.status(500).json({
            message: "Image to PDF conversion failed.",
            error: error.message
        });
    }
});

app.use((error, _req, res, _next) => {
    console.error("Request error:", error.message);
    res.status(500).json({ message: "Request failed.", error: error.message });
});

app.listen(PORT, () => {
    console.log(`Server running at http://localhost:${PORT}`);
    console.log("Cloud storage: Filebase");
});
