const xlsx = require('xlsx');
const puppeteer = require('puppeteer-extra'); 
const fs = require('fs');
const path = require('path');
 
// Add stealth plugin and use defaults 
const pluginStealth = require('puppeteer-extra-plugin-stealth'); 
const {executablePath} = require('puppeteer'); 
 
// Use stealth 
puppeteer.use(pluginStealth()); 

// Load URLs and folder names from an Excel file
let missingPage = 0;
const fileName = 'first-level-urls-part-4.xlsx';
const workbook = xlsx.readFile('ecommerce.xlsx');
const sheetName = 'MissingURLRemoved'; // Update with your sheet name
const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], { header: ['id', 'company', 'url', 'experience'] });
data.shift();
//console.log(data);

async function scrapper (url, company) {
  await puppeteer.launch({ executablePath: executablePath() }).then(async browser => { 
    // Create a new page 
    const page = await browser.newPage(); 
   
    // Setting page view 
    await page.setViewport({ width: 1280, height: 720 }); 

    try {
      const formattedUrl = url.startsWith('http://') || url.startsWith('https://') ? url : `http://${url}`;
      const formattedUrlAlternateProtocol = formattedUrl.startsWith('http://') ? `https://${url}` : `http://${url}`;

      await page.goto(formattedUrl, { waitUntil: 'domcontentloaded' });

      await page.waitForTimeout(10000); 

      const finalUrl = page.url();

      console.log(formattedUrl);
      console.log(formattedUrlAlternateProtocol);
      console.log(finalUrl);

      const hrefs = await page.$$eval('a', as => as.map(a => a.href));

      let numberOfHrefs = 0;

      //console.log(hrefs);

      for(let href of hrefs)
      {
        if(href.startsWith(formattedUrl) || href.startsWith(formattedUrlAlternateProtocol) || href.startsWith(finalUrl))
        {
            console.log(href);
            writeInExcel(company, url, href);
            numberOfHrefs++;
        }
      }

      console.log(`Extracted ${numberOfHrefs} URLS from ${formattedUrl}`);
    } catch (error) {
      console.error(`Error scraping: ${error.message}`);
      missingPage++;
    } finally {
      await page.close();
    }
   
    await browser.close(); 
  });
}

async function scrapePages() {
  for (const [index, { id, company, url, experience }] of data.entries()) {

    console.log(index+1);
    console.log(url);
    await scrapper(url, company);
  }

  console.log(missingPage);
}

function writeInExcel (company, url, href)
{
  // Check if the file exists
  if (fs.existsSync(fileName)) {
    // If the file exists, read the existing workbook
    const existingWorkbook = xlsx.readFile(fileName);

    // Assume the first sheet is the target sheet
    const ws = existingWorkbook.Sheets[existingWorkbook.SheetNames[0]];

    // Sample data to be appended
    const newRowData = [company, url, href];

    // Append the new row
    xlsx.utils.sheet_add_aoa(ws, [newRowData], { origin: -1 });

    // Write the updated workbook back to the file
    xlsx.writeFile(existingWorkbook, fileName, { bookType: 'xlsx', type: 'file' });

    //console.log('File updated in append mode.');
  } else {
    // If the file doesn't exist, create a new workbook with column names
    // Define column names
    const columns = ['company', 'homepage', 'url'];
    const wb = xlsx.utils.book_new();
    const ws = xlsx.utils.aoa_to_sheet([columns]);
    xlsx.utils.book_append_sheet(wb, ws, 'First Level URLs');
    // Sample data to be appended
    const newRowData = [company, url, href];
    // Append the new row
    xlsx.utils.sheet_add_aoa(ws, [newRowData], { origin: -1 });
    xlsx.writeFile(wb, fileName, { bookType: 'xlsx', type: 'file' });

    console.log('File created with column names.');
  }
}

scrapePages();

// scrapper("https://eshikhon.com/");
