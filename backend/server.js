const express = require("express");
const cors = require("cors");
const multer = require("multer");
const path = require("path");
const fs = require("fs");

const XLSX = require("xlsx");
const PDFDocument = require("pdfkit");
const { PDFParse } = require("pdf-parse");
const mammoth = require("mammoth");

const {
    Document,
    Packer,
    Paragraph,
    TextRun
} = require("docx");

const app = express();


// ==========================================
// BASIC SETTINGS
// ==========================================

app.use(cors());
app.use(express.json());


// ==========================================
// UPLOAD SETTINGS
// ==========================================

const upload = multer({
    dest: path.join(__dirname, "uploads/")
});


// ==========================================
// TEST API
// ==========================================

app.get("/", (req, res) => {

    res.json({
        message: "Serverless File Converter API is running!"
    });

});


// ==========================================
// UPLOAD FILE
// ==========================================

app.post("/upload", upload.single("file"), (req, res) => {

    if (!req.file) {

        return res.status(400).json({
            message: "No file uploaded."
        });

    }

    console.log(
        "File received:",
        req.file.originalname
    );

    res.json({

        message: "File uploaded successfully!",

        filename:
            req.file.originalname,

        filePath:
            req.file.path

    });

});


// ==========================================
// EXCEL → PDF
// ==========================================

app.post(
    "/convert/excel-to-pdf",
    upload.single("file"),
    (req, res) => {

        if (!req.file) {

            return res.status(400).json({
                message: "No Excel file uploaded."
            });

        }

        try {

            const workbook =
                XLSX.readFile(req.file.path);

            const sheetName =
                workbook.SheetNames[0];

            const worksheet =
                workbook.Sheets[sheetName];

            const data =
                XLSX.utils.sheet_to_json(
                    worksheet,
                    {
                        header: 1,
                        defval: ""
                    }
                );

            const outputName =
                path.parse(
                    req.file.originalname
                ).name + ".pdf";

            const outputPath =
                path.join(
                    __dirname,
                    "uploads",
                    outputName
                );

            const doc =
                new PDFDocument({
                    margin: 40,
                    size: "A4"
                });

            const stream =
                fs.createWriteStream(
                    outputPath
                );

            doc.pipe(stream);

            doc.fontSize(18)
                .text(
                    "Converted Excel File",
                    {
                        align: "center"
                    }
                );

            doc.moveDown();

            doc.fontSize(10);

            data.forEach(
                (row, rowIndex) => {

                    const rowText =
                        row
                            .map(
                                cell =>
                                    String(cell)
                            )
                            .join("    ");

                    if (rowIndex === 0) {

                        doc.font(
                            "Helvetica-Bold"
                        );

                    } else {

                        doc.font(
                            "Helvetica"
                        );

                    }

                    doc.text(rowText);

                    if (rowIndex === 0) {
                        doc.moveDown(0.5);
                    }

                }
            );

            doc.end();

            stream.on(
                "finish",
                () => {

                    res.json({

                        message:
                            "Excel converted to PDF successfully!",

                        downloadUrl:
                            `http://localhost:5000/download/${encodeURIComponent(outputName)}`

                    });

                }
            );

        } catch (error) {

            console.error(
                "Excel conversion error:",
                error
            );

            res.status(500).json({

                message:
                    "Excel conversion failed."

            });

        }

    }
);


// ==========================================
// PDF → EXCEL
// ==========================================

app.post(
    "/convert/pdf-to-excel",
    upload.single("file"),
    async (req, res) => {

        if (!req.file) {

            return res.status(400).json({

                message:
                    "No PDF file uploaded."

            });

        }

        try {

            const pdfBuffer =
                fs.readFileSync(
                    req.file.path
                );

            const parser =
                new PDFParse({
                    data: pdfBuffer
                });

            const pdfData =
                await parser.getText();

            const text =
                pdfData.text;

            const rows =
                text
                    .split("\n")
                    .map(
                        row => [row.trim()]
                    )
                    .filter(
                        row =>
                            row[0] !== ""
                    );

            const workbook =
                XLSX.utils.book_new();

            const worksheet =
                XLSX.utils.aoa_to_sheet(
                    rows
                );

            XLSX.utils.book_append_sheet(
                workbook,
                worksheet,
                "PDF Data"
            );

            const outputName =
                path.parse(
                    req.file.originalname
                ).name + ".xlsx";

            const outputPath =
                path.join(
                    __dirname,
                    "uploads",
                    outputName
                );

            XLSX.writeFile(
                workbook,
                outputPath
            );

            res.json({

                message:
                    "PDF converted to Excel successfully!",

                downloadUrl:
                    `http://localhost:5000/download/${encodeURIComponent(outputName)}`

            });

        } catch (error) {

            console.error(
                "PDF conversion error:",
                error
            );

            res.status(500).json({

                message:
                    "PDF conversion failed."

            });

        }

    }
);


// ==========================================
// PDF → WORD
// ==========================================

app.post(
    "/convert/pdf-to-word",
    upload.single("file"),
    async (req, res) => {

        if (!req.file) {

            return res.status(400).json({

                message:
                    "No PDF file uploaded."

            });

        }

        try {

            // Read PDF
            const pdfBuffer =
                fs.readFileSync(
                    req.file.path
                );


            // Create PDF parser
            const parser =
                new PDFParse({
                    data: pdfBuffer
                });


            // Extract text
            const pdfData =
                await parser.getText();


            const text =
                pdfData.text;


            // Split PDF text into paragraphs
            const paragraphs =
                text
                    .split("\n")
                    .map(
                        line => line.trim()
                    )
                    .filter(
                        line => line !== ""
                    );


            // Create Word document
            const children =
                paragraphs.map(
                    line =>
                        new Paragraph({

                            children: [

                                new TextRun({
                                    text: line
                                })

                            ]

                        })
                );


            const document =
                new Document({

                    sections: [

                        {

                            properties: {},

                            children:
                                children

                        }

                    ]

                });


            // Output filename
            const outputName =
                path.parse(
                    req.file.originalname
                ).name + ".docx";


            const outputPath =
                path.join(
                    __dirname,
                    "uploads",
                    outputName
                );


            // Create DOCX file
            const buffer =
                await Packer.toBuffer(
                    document
                );


            fs.writeFileSync(
                outputPath,
                buffer
            );


            res.json({

                message:
                    "PDF converted to Word successfully!",

                downloadUrl:
                    `http://localhost:5000/download/${encodeURIComponent(outputName)}`

            });

        } catch (error) {

            console.error(
                "PDF to Word conversion error:",
                error
            );

            res.status(500).json({

                message:
                    "PDF to Word conversion failed."

            });

        }

    }
);


// ==========================================
// JPG / JPEG / PNG → PDF
// ==========================================

app.post(
    "/convert/image-to-pdf",
    upload.single("file"),
    (req, res) => {

        if (!req.file) {

            return res.status(400).json({

                message:
                    "No image file uploaded."

            });

        }

        try {

            const extension =
                path.extname(
                    req.file.originalname
                ).toLowerCase();

            const allowedExtensions = [
                ".jpg",
                ".jpeg",
                ".png"
            ];

            if (
                !allowedExtensions.includes(
                    extension
                )
            ) {

                return res.status(400).json({

                    message:
                        "Only JPG, JPEG and PNG images are allowed."

                });

            }

            const outputName =
                path.parse(
                    req.file.originalname
                ).name + ".pdf";

            const outputPath =
                path.join(
                    __dirname,
                    "uploads",
                    outputName
                );

            const doc =
                new PDFDocument({

                    size: "A4",

                    margin: 40

                });

            const stream =
                fs.createWriteStream(
                    outputPath
                );

            doc.pipe(stream);

            doc.fontSize(18)
                .text(
                    "Converted Image",
                    {
                        align: "center"
                    }
                );

            doc.moveDown();

            doc.image(
                req.file.path,
                {

                    fit: [
                        515,
                        700
                    ],

                    align: "center",

                    valign: "center"

                }
            );

            doc.end();

            stream.on(
                "finish",
                () => {

                    res.json({

                        message:
                            "Image converted to PDF successfully!",

                        downloadUrl:
                            `http://localhost:5000/download/${encodeURIComponent(outputName)}`

                    });

                }
            );

        } catch (error) {

            console.error(
                "Image conversion error:",
                error
            );

            res.status(500).json({

                message:
                    "Image conversion failed."

            });

        }

    }
);


// ==========================================
// WORD / DOC / DOCX → PDF
// ==========================================

app.post(
    "/convert/word-to-pdf",
    upload.single("file"),
    async (req, res) => {

        if (!req.file) {

            return res.status(400).json({

                message:
                    "No Word file uploaded."

            });

        }

        try {

            const extension =
                path.extname(
                    req.file.originalname
                ).toLowerCase();

            const allowedExtensions = [
                ".doc",
                ".docx"
            ];

            if (
                !allowedExtensions.includes(
                    extension
                )
            ) {

                return res.status(400).json({

                    message:
                        "Only DOC and DOCX files are allowed."

                });

            }

            const result =
                await mammoth.convertToHtml({
                    path: req.file.path
                });

            const html =
                result.value;

            const plainText =
                html
                    .replace(
                        /<br\s*\/?>/gi,
                        "\n"
                    )
                    .replace(
                        /<\/p>/gi,
                        "\n\n"
                    )
                    .replace(
                        /<\/h[1-6]>/gi,
                        "\n\n"
                    )
                    .replace(
                        /<li>/gi,
                        "• "
                    )
                    .replace(
                        /<\/li>/gi,
                        "\n"
                    )
                    .replace(
                        /<[^>]*>/g,
                        ""
                    )
                    .replace(
                        /&nbsp;/g,
                        " "
                    )
                    .replace(
                        /&amp;/g,
                        "&"
                    )
                    .replace(
                        /&lt;/g,
                        "<"
                    )
                    .replace(
                        /&gt;/g,
                        ">"
                    )
                    .trim();

            const outputName =
                path.parse(
                    req.file.originalname
                ).name + ".pdf";

            const outputPath =
                path.join(
                    __dirname,
                    "uploads",
                    outputName
                );

            const doc =
                new PDFDocument({

                    size: "A4",

                    margin: 50

                });

            const stream =
                fs.createWriteStream(
                    outputPath
                );

            doc.pipe(stream);

            doc.fontSize(18)
                .font("Helvetica-Bold")
                .text(
                    "Converted Word Document",
                    {
                        align: "center"
                    }
                );

            doc.moveDown();

            doc.fontSize(11)
                .font("Helvetica")
                .text(
                    plainText,
                    {
                        align: "left",
                        lineGap: 5
                    }
                );

            doc.end();

            stream.on(
                "finish",
                () => {

                    res.json({

                        message:
                            "Word document converted to PDF successfully!",

                        downloadUrl:
                            `http://localhost:5000/download/${encodeURIComponent(outputName)}`

                    });

                }
            );

        } catch (error) {

            console.error(
                "Word conversion error:",
                error
            );

            res.status(500).json({

                message:
                    "Word to PDF conversion failed."

            });

        }

    }
);


// ==========================================
// LIST ALL FILES
// ==========================================

app.get(
    "/files",
    (req, res) => {

        const uploadFolder =
            path.join(
                __dirname,
                "uploads"
            );

        if (
            !fs.existsSync(
                uploadFolder
            )
        ) {

            return res.json([]);

        }

        const files =
            fs.readdirSync(
                uploadFolder
            );

        const fileList =
            files.map(
                filename => {

                    const filePath =
                        path.join(
                            uploadFolder,
                            filename
                        );

                    const stats =
                        fs.statSync(
                            filePath
                        );

                    return {

                        name:
                            filename,

                        size:
                            stats.size,

                        downloadUrl:
                            `http://localhost:5000/download/${encodeURIComponent(filename)}`

                    };

                }
            );

        res.json(fileList);

    }
);


// ==========================================
// DOWNLOAD FILE
// ==========================================

app.get(
    "/download/:filename",
    (req, res) => {

        const filename =
            path.basename(
                req.params.filename
            );

        const filePath =
            path.join(
                __dirname,
                "uploads",
                filename
            );

        if (
            !fs.existsSync(
                filePath
            )
        ) {

            return res.status(404).send(
                "File not found."
            );

        }

        res.download(
            filePath
        );

    }
);


// ==========================================
// ERROR HANDLER
// ==========================================

app.use(
    (err, req, res, next) => {

        console.error(
            "Server error:",
            err
        );

        res.status(500).json({

            message:
                "Something went wrong on the server."

        });

    }
);


// ==========================================
// START SERVER
// ==========================================

const PORT = 5000;

app.listen(
    PORT,
    () => {

        console.log(
            `Server running at http://localhost:${PORT}`
        );

    }
);