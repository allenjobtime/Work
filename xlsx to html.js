let xlsx = require("xlsx")
let fs = require('fs')
let wb = xslx.readfile(output.xlsx)
let name = wb.sheetnames[0]
let value = wb.sheets[name]
let htmldata = xlsx.utils.sheet_to_html(value);
console.log(htmldata);
fs.writeFile("htmltable.html", htmldata, function(err){
    if (err) {
        return console.log(err);
}});
