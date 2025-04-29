const express = require("express");
const XlsxPopulate = require("xlsx-populate");
const bodyParser = require("body-parser");

const app = express();
app.use(bodyParser.json());

// Enable CORS for dev
app.use((req, res, next) => {
    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
    res.setHeader("Access-Control-Allow-Headers", "Content-Type");
    next();
});

app.post("/generate-excel", async (req, res) => {
    const { filename = "report.xlsx", password = "", labels, data } = req.body;

    try {
        const workbook = await XlsxPopulate.fromBlankAsync();
        const sheet = workbook.sheet("Sheet1");

        // Insert headers (labels)
        const keys = Object.keys(labels);
        keys.forEach((key, index) => {
            const col = String.fromCharCode(65 + index);
            sheet.cell(`${col}1`).value(labels[key]);
        });

        data.forEach((row, rowIndex) => {
            keys.forEach((key, colIndex) => {
                const col = String.fromCharCode(65 + colIndex);
                const value = row[key] ?? "";
                sheet.cell(`${col}${rowIndex + 2}`).value(value);
            });
        });

        const buffer = await workbook.outputAsync(password ? { password } : undefined);

        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
        res.send(buffer);
    } catch (err) {
        console.error(err);
        res.status(500).send("Error generating Excel file");
    }
});

module.exports = app;