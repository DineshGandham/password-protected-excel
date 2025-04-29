// const XlsxPopulate = require('xlsx-populate');

// async function createPasswordProtectedExcel() {
//   const workbook = await XlsxPopulate.fromBlankAsync();
//   const sheet = workbook.sheet("Sheet1");

//   // Add some data
//   sheet.cell("A1").value("Name");
//   sheet.cell("B1").value("Age");
//   sheet.cell("A2").value("Alice");
//   sheet.cell("B2").value(28);

//   // Save the file with password protection
//   await workbook.toFileAsync("./protected.xlsx", { password: "mypassword123" });
//   console.log("Excel file created with password protection!");
// }

// createPasswordProtectedExcel();


// const express = require("express");
// const XlsxPopulate = require("xlsx-populate");
// const bodyParser = require("body-parser");

// const app = express();
// const port = 3001;

// app.use(bodyParser.json());

// // CORS (optional if frontend is separate)
// app.use((req, res, next) => {
//   res.setHeader("Access-Control-Allow-Origin", "*"); // or specific origin
//   res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
//   res.setHeader("Access-Control-Allow-Headers", "Content-Type");
//   next();
// });

// app.post("/generate-excel", async (req, res) => {
//   const { filename = "data.xlsx", password = "mypassword123" } = req.body;

//   try {
//     const workbook = await XlsxPopulate.fromBlankAsync();
//     const sheet = workbook.sheet("Sheet1");

//     sheet.cell("A1").value("Name");
//     sheet.cell("B1").value("Age");
//     sheet.cell("A2").value("Alice");
//     sheet.cell("B2").value(28);

//     // Create the file buffer
//     const buffer = await workbook.outputAsync({ password });

//     // Set headers to force download
//     res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
//     res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
//     res.send(buffer);
//   } catch (err) {
//     console.error(err);
//     res.status(500).send("Error generating Excel file");
//   }
// });

// app.listen(port, () => {
//   console.log(`Excel API running on http://localhost:${port}`);
// });

const express = require("express");
const XlsxPopulate = require("xlsx-populate");
const bodyParser = require("body-parser");

const app = express();
const port = 3001;

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

app.listen(port, () => {
  console.log(`Excel API running on http://localhost:${port}`);
});
