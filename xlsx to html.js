let xlsx = require('xlsx');
let fs = require('fs');
let INPUT_FILE = "output.xlsx";
let OUTPUT_FILE = "htmltable.html";
let CSS_STYLE = `
<style>
    table {
        border-collapse: collapse; 
        width: 100%;
        font-family: Arial, Helvetica, sans-serif;
    }
    th, td {
        border: 1px solid #333; 
        padding: 8px;
        text-align: left;
    }
</style>
`;

try {
    let wb = xlsx.readFile(INPUT_FILE);
    let sheetName = wb.SheetNames[0];
    let ws = wb.Sheets[sheetName];

    let rawTableHtml = xlsx.utils.sheet_to_html(ws);

    let finalHtmlData = `
<!DOCTYPE html>
<html lang="en"> 
<head>
    <meta charset="UTF-8">
    <title>25Live Room Reservations</title>
    ${CSS_STYLE}
</head>
<body>
    ${rawTableHtml}
</body>
</html>
`;

    fs.writeFileSync(OUTPUT_FILE, finalHtmlData);
    console.log("Outputted File");
} catch (e) {
    console.error("Error");
}