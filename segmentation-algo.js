const fs = require('fs');
const path = require('path');
const puppeteer = require('puppeteer');
const XLSX = require('xlsx');

const jsFilePath = 'block_segment.js'; // Path to your JavaScript file
const fileName = 'cdp.xlsx';
let counter = 0;

// Function to process a single HTML file
async function processHtmlFile(filePath, browser) {
  console.log(++counter);
  
  const page = await browser.newPage();

  try {
    const htmlContent = fs.readFileSync(filePath, 'utf8');
    //const jsCode = fs.readFileSync(jsFilePath, 'utf8');

    //console.log(jsCode);

    // Load HTML content into Puppeteer page
    await page.setContent(htmlContent);

    // Load yourScript.js into the Puppeteer page
    //const jsCode = require(jsFilePath);
    //await page.evaluate(jsCode);

    // Call the exported function to get the results
    //let results = await page.evaluate(() => getCandidateDpTexts());

    // ******************important to explore**************************
    // const extractedText = await page.$eval('*', (el) => el.innerText);
    const results = await page.$$eval(':not(p):not(script):not(style):not(noscript):not(br):not(hr):not(:has(*))', elements => elements.map(el => el.innerText));
    
    //console.log(extractedText);

    // Log or use the results as needed
    // console.log(`Results of ${filePath}:`);

    segments = [];

    for(let result of results)
    {
      if (result != null)
      {
        try {
          let line = result.trim();
          if (line.length > 0) //&& line.length < 100)
          {
            segments.push(line);
          }
        }
        catch(error){
          continue;
        }
      }
    }

    writeInExcel(filePath, segments)

    console.log(segments.length);
  }
  catch (error) {
    console.log(filePath + ' could not be analyzed');
    console.log(error);
  }
  finally {
    await page.close();
  }
}

// Function to process all HTML files in a directory recursively
async function processAllHtmlFiles(directoryPath, browser) {
  const files = fs.readdirSync(directoryPath);

  for (const file of files) {
    const filePath = path.join(directoryPath, file);
    const stats = fs.statSync(filePath);

    if (stats.isDirectory()) {
      // Recursively process subdirectories
      if(parseInt(file)>0)
      {
        await processAllHtmlFiles(filePath, browser);
      }
    } else if (stats.isFile() && file.endsWith('.html')) {
      // Process HTML files
      await processHtmlFile(filePath, browser);
      console.log(`Processed: ${filePath}`);
    }
  }
}

// Replace 'path/to/your/html/files' with the actual path to your HTML files
const htmlFilesDirectory = 'data';

async function main ()
{
  const browser = await puppeteer.launch();
  await processAllHtmlFiles(htmlFilesDirectory, browser);
  //await processHtmlFile("C:/Users/Yasin Sazid/Downloads/eshikhon.html", browser);
  await browser.close();
}

main();

function writeInExcel (filePath, segments)
{
  // Check if the file exists
  if (fs.existsSync(fileName)) {
    // If the file exists, read the existing workbook
    const existingWorkbook = XLSX.readFile(fileName);

    // Assume the first sheet is the target sheet
    const ws = existingWorkbook.Sheets[existingWorkbook.SheetNames[0]];

    for (let segment of segments)
    {
      // Sample data to be appended
      const newRowData = [filePath, segment];

      console.log(newRowData);

      // Append the new row
      XLSX.utils.sheet_add_aoa(ws, [newRowData], { origin: -1 });
    }

    // Write the updated workbook back to the file
    XLSX.writeFile(existingWorkbook, fileName, { bookType: 'xlsx', type: 'file' });

    console.log('File updated in append mode.');
  } else {
    // If the file doesn't exist, create a new workbook with column names
    // Define column names
    const columns = ['origin', 'candidate text'];
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([columns]);
    XLSX.utils.book_append_sheet(wb, ws, 'Candidate Texts Info');

    for (let segment of segments)
    {
      // Sample data to be appended
      const newRowData = [filePath, segment];

      console.log(segment);

      // Append the new row
      XLSX.utils.sheet_add_aoa(ws, [newRowData], { origin: -1 });
    }

    XLSX.writeFile(wb, fileName, { bookType: 'xlsx', type: 'file' });

    console.log('File created with column names.');
  }
}
