let xlsx = require('xlsx');
let fs = require('fs');
let inputfile = "output.xlsx";
let outputfile = "htmltable.html";
let style = `
<style>
    body {
        background-color: #009dd1;
    }
    table {
        background-color: white;
        border-collapse: collapse; 
        width: 100%;
        font-family: Arial, Helvetica, sans-serif;
    }
    th, td {
        background-color: white;
        border: 1px solid #333; 
        padding: 8px;
        text-align: left;
        color: #003552;
    }
    table tr:first-child td {
        font-weight: bold;
        color: black;
    }
    td:first-child {
        font-weight: bold;
        color: black;
    }
</style>
`;

let script = `
<script>
    document.addEventListener('DOMContentLoaded', () => {
        const dropdown = document.getElementById('RoomID_dropdown');
        const table = document.querySelector('table');
        const roomcolumn = 0;

        const rows = Array.from(table.querySelectorAll('tr')).slice(1);

        dropdown.addEventListener('change', (event) => {
            const selectedroomID = event.target.value;
            rows.forEach(row => {
                const roomcellID = row.cells[roomcolumn]; 
                if (roomcellID) {
                    const rowText = roomcellID.textContent.trim();
                    
                    if (selectedroomID === "") {
                        row.style.display = ""; 
                    } else if (rowText === selectedroomID) {
                        row.style.display = ""; 
                    } else {
                        row.style.display = "none"; 
                    }
                }
            });
        });
    });
</script>
`;

try {
    let wb = xlsx.readFile(inputfile);
    let sheetName = wb.SheetNames[0];
    let ws = wb.Sheets[sheetName];

    let rawTableHtml = xlsx.utils.sheet_to_html(ws);
    let data = xlsx.utils.sheet_to_json(ws, { header: 1 });
const uniqueroomIDs = new Set();
const roomcolumn = 0;
    for (let i = 1; i < data.length; i++) {
        let rawroomID = data[i][roomcolumn];

        if (rawroomID) {
            if (rawroomID.startsWith("Room: ")) {
                rawroomID = rawroomID.substring("Room: ".length);
            }
            uniqueroomIDs.add(String(rawroomID).trim());
        }
    }
    let dropdownOptions = '<option value="">Show All Rooms</option>';
    uniqueroomIDs.forEach(roomID => {
        dropdownOptions += `<option value="${roomID}">${roomID}</option>`;
    });
    const dropdownHtml = `
        <div style="margin-bottom: 20px; margin-top: 20px; font-family: Arial, sans-serif; font-weight: bold; color: white;">
        <label for="RoomID_dropdown">Filter by Room:</label>
    <select id="RoomID_dropdown">
        ${dropdownOptions}
    </select>
</div>
    `;
        let finalHtmlData = `
<!DOCTYPE html>
<html lang="en"> 
<head>
    <meta charset="UTF-8">
    <title>25Live Room Reservations</title>
    ${style}
</head>
<body>
    ${dropdownHtml}
    ${rawTableHtml}
    ${script}
</body>
</html>
`;

    fs.writeFileSync(outputfile, finalHtmlData);
    console.log("Outputted File");
} catch (e) {
    console.error("Error", e.message);
}