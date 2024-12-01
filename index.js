import puppeteer from 'puppeteer';
import * as XLSX from 'xlsx';
import fs from 'fs';

(async () => {
  const delay = (time) => new Promise(resolve => setTimeout(resolve, time));

  const browser = await puppeteer.launch({
    headless: false,
    defaultViewport: null
  });
  const page = await browser.newPage();

  await page.goto('https://www.claimservices.com.ar/login.php');

  await page.setViewport({ width: 1080, height: 1024 });

  await page.type('#username', 'CASADPARABRISAS');
  await page.type('#password', 'cdp1005');

  await page.click('#btnLogin');

  await page.waitForNavigation({ waitUntil: 'networkidle2' });

  await page.goto('https://www.claimservices.com.ar/main.php?screen=claims&stateDashboard=I', { waitUntil: 'networkidle2' });

  await page.waitForSelector('#dt_claimsQuote_processing', { hidden: true });

  await delay(4000);

  const outsideTableData = [];

  const headerRow = [
    'Id', 'Fecha', 'Nombre y Apellido', 'Nro de Reclamo', 'Telefono', 'Vehiculo', 'Patente', 'Tipo de Siniestro', 'Compañia', 'Estado',
    'Marca', 'Modelo', 'Chasis', 'Año del Vehiculo', 'Patente', 'Provincia', 'Localidad', 'Cristales/Cerraduras'
  ];

  outsideTableData.push(headerRow);

  const fetchTableData = async (processRows = true) => {
    await delay(5000);
    const tableData = await page.$$eval('table tr', rows => {
      return rows.slice(1).map(row => {
        const cells = Array.from(row.querySelectorAll('td'));
        return cells.map(cell => cell.innerText);
      });
    });

    if (!processRows) {
      return tableData.length;
    }

    const rows = await page.$$('table tr');
    for (let i = 1; i < 11; i++) {
      const row = tableData[i - 1];

      try {
        await delay(2000);
        await rows[i].click();
        await delay(2000);
        await page.click("#btnQuoteAll");
        await page.waitForSelector("#claimsDialog", { visible: true });
        await delay(2000);

        const crystalTableData = await page.$$eval('.table.table-bordered.bootstrap-datatable.dataTable', tables => {
          const rows = tables[3].querySelectorAll('tr');
          const firstRow = Array.from(rows).slice(1, 2);
          return firstRow.map(row => {
            const cells = Array.from(row.querySelectorAll('td'));
            return cells[0].innerText;
          });
        });

        const additionalData = await page.evaluate(() => {
          const detailField1 = document.querySelector('#carBrand')?.value;
          const detailField2 = document.querySelector('#carDescription')?.value;
          const detailField3 = document.querySelector('#carChassis')?.value;
          const detailField4 = document.querySelector('#carYears')?.value;
          const detailField5 = document.querySelector('#patent')?.value;
          const detailField6 = document.querySelector('#province')?.value;
          const detailField7 = document.querySelector('#location')?.value;

          return [detailField1, detailField2, detailField3, detailField4, detailField5, detailField6, detailField7];
        });
        additionalData.push(crystalTableData[0]);

        row.push(...additionalData);
        outsideTableData.push(row);
        await page.evaluate(() => {
          document.querySelector("#claimsDialog .modal-header .close").click();
        });
        await page.waitForSelector("#claimsDialog", { hidden: true });
        await delay(1000);
      } catch (error) {
        console.error(`Error processing row ${i}:`, error);
      }
    }
    return tableData.length;
  };

  let initialTableDataLength = await fetchTableData(false);
  const maxPages = Math.ceil(initialTableDataLength / 10);
  for (let i = 0; i < maxPages; i++) {
    await fetchTableData();
    await page.click('.next');
    await delay(4000);
  }

  console.log(outsideTableData);

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(outsideTableData);


  XLSX.utils.book_append_sheet(wb, ws, 'Claims Data');

  XLSX.writeFile(wb, 'claims_data.xlsx');

  await browser.close();
})();