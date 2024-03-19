// import OpenAI from 'openai';
const { OpenAI } = require("openai");

const fs = require('fs');
const XLSX = require('xlsx');
const fileName = 'gpt_remaining_classification_result_2.xlsx';
let counter = 0;

const workbook = XLSX.readFile('gpt_remaining_dataset.xlsx');
const sheetName = 'gpt_remaining_dataset';
const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: ['origin', 'candidate_text', 'origin_site'] });
data.shift();

// Specify the file path
const filePath = 'prompt.txt';
let prompt = "";

let newData = [];

// Read the contents of the file as a string synchronously
try {
    fileContents = fs.readFileSync(filePath, 'utf8');
  
    prompt = fileContents;
} catch (err) {
    console.error('Error reading the file synchronously:', err);
}

let APIcall = async (text) => { 
    const openai = new OpenAI({
        apiKey: 'sk-something'
    });
    const user_input = prompt + ' ' + text;
    // console.log(user_input);
    try { 
        const chatCompletion = await openai.chat.completions.create({
            model: "gpt-3.5-turbo-0125",
            messages: [{"role": "user", "content": user_input}],
            temperature: 0
        });
        const output_text = chatCompletion.choices[0].message.content;
        // console.log(output_text); 
        return output_text;
    } catch (err) { 
        if (err.response) { 
            console.log(err.response.status); 
            console.log(err.response.data); 
        } else { 
            console.log(err.message); 
        } 
    }
};

function writeInExcel (site, page, text, label)
{
  // Check if the file exists
  if (fs.existsSync(fileName)) {

    const newRowData = [site, page, text, label];

    newData.push(newRowData);

    if (newData.length==100)
    {
        // If the file exists, read the existing workbook
        const existingWorkbook = XLSX.readFile(fileName);

        // Assume the first sheet is the target sheet
        const ws = existingWorkbook.Sheets[existingWorkbook.SheetNames[0]];

        // console.log(newRowData);

        for (let data of newData)
        {
            // Append the new row
            // console.log(data);
            XLSX.utils.sheet_add_aoa(ws, [data], { origin: -1 });
        }

        // Write the updated workbook back to the file
        XLSX.writeFile(existingWorkbook, fileName, { bookType: 'xlsx', type: 'file' });

        console.log('File updated in append mode.');

        newData = [];
    }

  } else {
    // If the file doesn't exist, create a new workbook with column names
    // Define column names
    const columns = ['origin_site', 'origin', 'candidate_text', 'gpt_label'];
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([columns]);
    XLSX.utils.book_append_sheet(wb, ws, 'GPT Classification Result');

    // Sample data to be appended
    const newRowData = [site, page, text, label];

    // Append the new row
    XLSX.utils.sheet_add_aoa(ws, [newRowData], { origin: -1 });

    XLSX.writeFile(wb, fileName, { bookType: 'xlsx', type: 'file' });

    console.log('File created with column names.');
  }
}

function sleep(milliseconds) {
    console.log("Sleeping");
    return new Promise(resolve => setTimeout(resolve, milliseconds));
}

async function main (){

    // console.log(prompt);


    for (const [index, { origin, candidate_text, origin_site }] of data.entries()) {

        console.log(index+1);

        response = await APIcall(candidate_text);

        // response = 'test';

        console.log(candidate_text + "-----------" + response);

        writeInExcel(origin_site, origin, candidate_text, response);

        console.log("done");
    }
}

main();

