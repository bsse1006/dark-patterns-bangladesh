const puppeteer = require('puppeteer');
const fs = require('fs');
const XLSX = require('xlsx');

const baseURL = 'https://e-cab.net/company-profile/';

let companies = [];

const fileName = 'ecommerce.xlsx';

async function scrape ()
{
    const browser  = await puppeteer.launch();

    for (let i=1; i<=2238; i++)
    {
        console.log(i);

        try{
            const page = await browser.newPage();

            const id = convertToFourDigitString(i);
            await page.goto(baseURL + id);

            const [el] = await page.$x('/html/body/div[1]/div/div/div[1]/div[2]/div[1]/div[2]/div/div[1]/h4');
            const title = await el.getProperty('textContent');
            const companyName = await title.jsonValue();

            const [el2] = await page.$x('/html/body/div[1]/div/div/div[2]/div/div[2]/p[4]/a');
            const link = await el2.getProperty('textContent');
            const companyURL = await link.jsonValue();

            const [el3] = await page.$x('/html/body/div[1]/div/div/div[1]/div[2]/div[1]/div[1]/div[3]/div[2]/div/h4/b');
            const ex = await el3.getProperty('textContent');
            const experience = await ex.jsonValue();

            companies.push([id, companyName.trim(), companyURL, experience]);

            // Check if the file exists
            if (fs.existsSync(fileName)) {
                // If the file exists, read the existing workbook
                const existingWorkbook = XLSX.readFile(fileName);

                // Assume the first sheet is the target sheet
                const ws = existingWorkbook.Sheets[existingWorkbook.SheetNames[0]];

                // Sample data to be appended
                const newRowData = [id, companyName.trim(), companyURL, experience];

                // Append the new row
                XLSX.utils.sheet_add_aoa(ws, [newRowData], { origin: -1 });

                // Write the updated workbook back to the file
                XLSX.writeFile(existingWorkbook, fileName, { bookType: 'xlsx', type: 'file' });

                console.log('File updated in append mode.');
            } else {
                // If the file doesn't exist, create a new workbook with column names
                // Define column names
                const columns = ['id', 'company', 'url', 'experience'];
                const wb = XLSX.utils.book_new();
                const ws = XLSX.utils.aoa_to_sheet([columns]);
                XLSX.utils.book_append_sheet(wb, ws, 'ECAB Company Info');
                // Sample data to be appended
                const newRowData = [id, companyName.trim(), companyURL, experience];
                // Append the new row
                XLSX.utils.sheet_add_aoa(ws, [newRowData], { origin: -1 });
                XLSX.writeFile(wb, fileName, { bookType: 'xlsx', type: 'file' });

                console.log('File created with column names.');
            }
        }
        catch (error)
        {
            console.log('Company ID ' + i + ' not found');
            continue;
        }
        finally
        {
            await page.close();
        }
    }

    browser.close();
    
    console.log('Excel file created successfully.');

    //console.log(companies);
}

function convertToFourDigitString(number) {
    // Ensure the input is a valid number
    if (isNaN(number)) {
        return "Invalid input";
    }

    // Convert the number to a string
    let numberString = number.toString();

    // Check the length of the string
    if (numberString.length >= 4) {
        // If the number has 4 or more digits, return the original string
        return numberString;
    } else {
        // If the number has fewer than 4 digits, add preceding zeros
        const numberOfZeros = 4 - numberString.length;
        const zeros = '0'.repeat(numberOfZeros);
        return zeros + numberString;
    }
}

scrape();