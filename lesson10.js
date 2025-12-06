/* const fs = require('fs');

try {
    const data = fs.readFileSync('newFile.txt', 'utf8');
    let words = data.trim().split(" ");
    let wordCount = words.length;

    console.log("the words are: ", words, "and the amount of words are:", wordCount);
} catch (err) {
    console.error("error: ", err);
} */


const xlsx = require('xlsx')

const workbook = xlsx.readFile("Book1.xlsx") ;
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const data = xlsx.utils.sheet_to_json(sheet) ;
const grades = data.map(val => val.grade);
const total = grades.reduce((acc,curr) => acc + curr,0);
const avg = total / grades.length;
console.log(avg);
