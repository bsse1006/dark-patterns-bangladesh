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
const workbook = xlsx.readFile('ecommerce.xlsx');
const sheetName = 'allURLs'; // Update with your sheet name
const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], { header: ['id', 'company', 'homepage_url', 'category_page_url', 'product_page_url'] });
data.shift();
//console.log(data);

async function scrapper (url, folderPath, fileName) {
  await puppeteer.launch({ executablePath: executablePath() }).then(async browser => { 
    // Create a new page 
    const page = await browser.newPage(); 
   
    // Setting page view 
    await page.setViewport({ width: 1280, height: 720 }); 

    try {
      const formattedUrl = url.startsWith('http://') || url.startsWith('https://') ? url : `http://${url}`;

      console.log(formattedUrl);

      await page.goto(formattedUrl, { waitUntil: 'domcontentloaded' });

      await page.waitForTimeout(10000); 

      // Save the HTML content to a file
      const htmlContent = await page.content();

      if (!fs.existsSync(folderPath)) {
        fs.mkdirSync(folderPath);
      }
      const filePath = path.join(folderPath, fileName);
      fs.writeFileSync(filePath, htmlContent);

      console.log(`Scrapped ${formattedUrl} and saved HTML to ${filePath}`);
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
  // const browser = await puppeteer.launch();

  for (const [index, { id, homepage_url, category_page_url, product_page_url, company }] of data.entries()) {
    console.log(index+1);

    // Create a folder for each URL
    let folderName = company; //|| formattedUrl.replace(/(^\w+:|^)\/\//, '').replace(/\//g, '_');
    folderName = `${id}.` + folderName;
    const folderPath = path.join('data', folderName);

    scrapper(homepage_url, folderPath, '1.index.html');
    scrapper(category_page_url, folderPath, '2.category.html');
    await scrapper(product_page_url, folderPath, '3.product.html');
  }

  // await browser.close();

  console.log(missingPage);
}

scrapePages();
