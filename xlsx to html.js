let xlsx = require('xlsx')
let fs = require('fs')
let wb = xlsx.readFile("output.xlsx")
let name = wb.SheetNames[0]
let value = wb.Sheets[name]
let htmldata = xlsx.utils.sheet_to_html(value);
console.log(htmldata);
fs.writeFile("htmltable.html", htmldata, function(err){
    if (err) {
        return console.log(err);
}});
