// server.js
require('dotenv').config();
const express = require('express');
const axios = require('axios');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const app = express();
const PORT = process.env.PORT || 3002;;
const { exec } = require('child_process');
const XLSX = require('xlsx');




app.use(express.static(path.join(__dirname, 'public')));

app.get('/api/token', async (req, res) => {
  try {
    const params = new URLSearchParams();
    params.append('client_id', process.env.FORGE_CLIENT_ID);
    params.append('client_secret', process.env.FORGE_CLIENT_SECRET);
    params.append('grant_type', 'client_credentials');
    params.append('scope', 'bucket:create bucket:read data:read data:write viewables:read bucket:delete');

    const response = await axios.post(
      'https://developer.api.autodesk.com/authentication/v2/token',
      params.toString(),
      { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    );

    res.json({
      
      access_token: response.data.access_token,
      expires_in: response.data.expires_in
    });
  } catch (err) {
    console.error('❌ Token fetch error:', err.response?.data || err.message);
    res.status(500).send('Failed to get access token');
  }
});

const bodyParser = require('body-parser');
const nodemailer = require('nodemailer');

// Middleware to parse JSON body
app.use(bodyParser.json());

function generateQuoteNumber() {
  const now = new Date();
  const timestamp = now.getTime();
  const last3 = timestamp % 1000;
  return last3.toString().padStart(3, '0'); 
}
function parseNumber(value) {
  if (!value) return 0;
  return parseInt(String(value).replace(/[^0-9]/g, ''), 10);
}
async function updateExcelWithNamedCells({ customer, productType, configuration, quoteNo }) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(path.join(__dirname, 'QuoteTemplate.xlsx'));
  const sheet = workbook.getWorksheet(1);
    
  sheet.getCell('B10').value = customer.name;
  sheet.getCell('B11').value = customer.email;
  sheet.getCell('B12').value = customer.phone;
  sheet.getCell('B13').value = customer.dealer;
  sheet.getCell('G4').value = { formula: 'TODAY()', result: new Date() };
  sheet.getCell('G5').value = quoteNo;
  sheet.getCell('E10').value = `${productType}`;
  sheet.getCell('B38').value = 'This Is a quotation on the goods named, subject to the condition noted below: (Describe any conditions pertaining to these prices and any additional terms of the agreement. You may want to include contingencies that will affect the quotation.';
  sheet.getCell('B43').value = 'THANK YOU FOR YOUR BUSINESS';

  
  

  if (productType === 'Trailer Hitches') {

    
  // ===== LABELS & VALUES =====
  sheet.getCell('B18').value = 'Hitch Type';
  sheet.getCell('C18').value = configuration.Hitchtype;

  sheet.getCell('B19').value = 'Shape';
  sheet.getCell('C19').value = configuration.Shape;

  sheet.getCell('B20').value = 'Hitch Location';
  sheet.getCell('C20').value = configuration.Hitchlocation;

  sheet.getCell('B21').value = 'Hitch Class';
  sheet.getCell('C21').value = configuration.Hitchclass;

  sheet.getCell('B22').value = 'OEM';
  sheet.getCell('C22').value = configuration.OEM;

  sheet.getCell('B23').value = 'Payload Range';
  sheet.getCell('C23').value = configuration.PayloadRange;

  sheet.getCell('B24').value = 'Hardware Needs';
  sheet.getCell('C24').value = configuration.Hardwareneeds;

  // ===== QUANTITIES =====
  sheet.getCell('E18').value = '1';
  

  // ===== PRICING LOGIC =====
  let hitchtypePrice = 0;
  let ShapePrice = 0;
  let hitchlocationPrice = 0;
  let hitchclassprice = 0;
  let OEMprice=0;
  let payloadprice=0;
  let Hardwareneeds=0;

  // Hitchtype
  if (configuration.Hitchtype === 'Rear Mount') hitchtypePrice = 475;
  else if (configuration.Hitchtype === 'Front Mount') hitchtypePrice = 265;
  else if (configuration.Hitchtype === 'Multifit') hitchtypePrice = 127.50;
  else if (configuration.Hitchtype === 'Bumper Mount') hitchtypePrice = 127.50;
  

  // Shape
  if (configuration.Shape === 'Tube') ShapePrice = 0;
  else if (configuration.Shape === 'Ball') ShapePrice = 35;

  // Hitch location
  if (configuration.Hitchlocation === 'Front') hitchlocationPrice = 28;
  else if (configuration.Hitchlocation === 'Rear') hitchlocationPrice = 0;

  // Hitch class
  if (configuration.Hitchclass === 'Class 1') hitchclassprice = 350;
  else if (configuration.Hitchclass === 'Class 2') hitchclassprice = 375;
  else if (configuration.Hitchclass === 'Class 3') hitchclassprice = 420;
  else if (configuration.Hitchclass === 'Class 4') hitchclassprice = 530;
  else if (configuration.Hitchclass === 'Class 5') hitchclassprice = 574;

  // OEM
  if (configuration.OEM === 'RAM') OEMprice = 6;
  else if (configuration.OEM === 'FORD') OEMprice = 8;
  else if (configuration.OEM === 'CHEVY') OEMprice = 6;

   // payload range
  if (configuration.PayloadRange === '2K') payloadprice = 15;
  else if (configuration.PayloadRange === '3K') payloadprice = 22;
  else if (configuration.PayloadRange === '5K') payloadprice = 28;

   // Hardwareneeds
  if (configuration.Hardwareneeds === 'Yes') Hardwareneeds = 24;
  else if (configuration.Hardwareneeds === 'No') Hardwareneeds = 0;

  // ===== WRITE PRICES =====
  sheet.getCell('F18').value = hitchtypePrice;
  sheet.getCell('F19').value = ShapePrice;
  sheet.getCell('F20').value = hitchlocationPrice;
  sheet.getCell('F21').value = hitchclassprice;
  sheet.getCell('F22').value = OEMprice;
  sheet.getCell('F23').value = payloadprice;
  sheet.getCell('F24').value = Hardwareneeds;
}


  else if (productType === '5th Wheel Hitches') {

  // Labels
  sheet.getCell('B18').value = 'Load Rating';
  sheet.getCell('C18').value = configuration.Loadrating;

  sheet.getCell('B19').value = 'Type';
  sheet.getCell('C19').value = configuration.Type;

  sheet.getCell('B20').value = 'Profile';
  sheet.getCell('C20').value = configuration.Profile;

  sheet.getCell('B21').value = 'OEM';
  sheet.getCell('C21').value = configuration.OEM;

  sheet.getCell('B22').value = 'Plug Type';
  sheet.getCell('C22').value = configuration.Plugtype;

  // Quantities (fixed to 1)
  sheet.getCell('E18').value = '1';
  

  // Pricing logic (EDIT PRICES IF NEEDED)
  let ratingPrice = 0;
  let typePrice = 0;
  let ProfilePrice = 0;
  let OEMPrice = 0;
  let PlugtypePrice = 0;

  if (configuration.Loadrating === '16K') ratingPrice = 1600;
  else if (configuration.Loadrating === '20K') ratingPrice = 1640;
  else if (configuration.Loadrating === '24K') ratingPrice = 1740;
  else if (configuration.Loadrating === '25K') ratingPrice = 1810;
  else if (configuration.Loadrating === '30K') ratingPrice = 1932;

  if (configuration.Type === 'Sliding') typePrice = 120; 
 else if (configuration.Type === 'Stationary') typePrice = 0; 

 if (configuration.Profile === 'Mid') ProfilePrice = 0; 
 else if (configuration.Profile === 'Full') ProfilePrice = 235; 

if (configuration.OEM === 'RAM') OEMPrice = 8;
  else if (configuration.OEM === 'FORD') OEMPrice = 12;
  else if (configuration.OEM === 'CHEVY') OEMPrice = 14;

  if (configuration.Plugtype === '5 Way') PlugtypePrice = 0;
  else if (configuration.Plugtype === '7 Way') PlugtypePrice = 48;


  sheet.getCell('F18').value = ratingPrice;
  sheet.getCell('F19').value = typePrice;
  sheet.getCell('F20').value = ProfilePrice;
  sheet.getCell('F21').value = OEMPrice;
  sheet.getCell('F22').value = PlugtypePrice;
}
else if (productType === 'Gooseneck Hitches') {

  // ===== LABELS & VALUES =====
  sheet.getCell('B18').value = 'Gooseneck Type';
  sheet.getCell('C18').value = configuration.Goosenecktype;

  sheet.getCell('B19').value = 'Size';
  sheet.getCell('C19').value = configuration.Size;

  sheet.getCell('B20').value = 'OEM';
  sheet.getCell('C20').value = configuration.OEM;

  sheet.getCell('B21').value = 'Plug Type';
  sheet.getCell('C21').value = configuration.Plugtype;

  sheet.getCell('B22').value = 'Rating Range';
  sheet.getCell('C22').value = configuration.Ratingrange;

  sheet.getCell('B23').value = 'Cover Needs';
  sheet.getCell('C23').value = configuration.Coverneeds;

  // ===== QUANTITIES =====
  sheet.getCell('E18').value = '1';


  // ===== PRICING =====
  let GoosenecktypePrice = 0;
  let SizePrice = 0;
  let OEMPrice = 0;
  let PlugtypePrice = 0;
  let RatingrangePrice = 0;
  let CoverneedsPrice = 0;

  // Gooseneck type pricing
  if (configuration.Goosenecktype === 'Folding Ball') GoosenecktypePrice = 220;
  else if (configuration.Goosenecktype === 'Fixed Ball') GoosenecktypePrice = 0;


  // Size pricing
  if (configuration.Size === '2-5/16 in') SizePrice = 29;
  else if (configuration.Size === '3 in') SizePrice = 0;

  // OEM pricing
  if (configuration.OEM === 'FORD') OEMPrice = 25;
  else if (configuration.OEM === 'RAM') OEMPrice = 25;
  else if (configuration.OEM === 'CHEVY') OEMPrice = 25;

  // Plugtype pricing
  if (configuration.Plugtype === '5 Way') PlugtypePrice = 0;
  else if (configuration.Plugtype === '7 Way') PlugtypePrice = 42;

  //Rating range pricing
  if (configuration.Ratingrange === '7K') RatingrangePrice = 1150;
  else if (configuration.Ratingrange === '10K') RatingrangePrice = 1250;
  else if (configuration.Ratingrange === '12K') RatingrangePrice = 1380;

   // Coverneeds pricing
  if (configuration.Coverneeds === 'Yes') CoverneedsPrice = 40;
  else if (configuration.Coverneeds === 'No') CoverneedsPrice = 0;


  // ===== WRITE PRICES =====
  sheet.getCell('F18').value = GoosenecktypePrice;
  sheet.getCell('F19').value = SizePrice;
  sheet.getCell('F20').value = OEMPrice;
  sheet.getCell('F21').value = PlugtypePrice;
  sheet.getCell('F22').value = RatingrangePrice;
  sheet.getCell('F23').value = CoverneedsPrice;
}
else if (productType === 'Towing Accessories') {

  // ===== LABELS & VALUES =====
  sheet.getCell('B18').value = 'Hitch Type';
  sheet.getCell('C18').value = configuration.Hitchtype;

  sheet.getCell('B19').value = 'Hardware';
  sheet.getCell('C19').value = configuration.Hardware;

  sheet.getCell('B20').value = 'Electrical Systems';
  sheet.getCell('C20').value = configuration.Electrical;

  sheet.getCell('B21').value = 'Safety Kits';
  sheet.getCell('C21').value = configuration.Safetykits;

  sheet.getCell('B22').value = 'Covers';
  sheet.getCell('C22').value = configuration.Covers;

  sheet.getCell('B23').value = 'Wheel Chocks';
  sheet.getCell('C23').value = configuration.Wheelchocks;

  sheet.getCell('B24').value = 'Extended Mirrors';
  sheet.getCell('C24').value = configuration.Extendedmirrors;

  // ===== QUANTITIES =====
  sheet.getCell('E18').value = '1';
 

  // ===== PRICING =====
  let HitchtypePrice = 0;
  let HardwarePrice = 0;
  let ElectricalPrice = 0;
  let SafetykitsPrice = 0;
  let CoversPrice = 0;
  let WheelchocksPrice = 0;
  let ExtendedmirrorsPrice = 0;

  // Gooseneck type pricing
  if (configuration.Hitchtype === 'Trailer Hitches') HitchtypePrice = 210;
  else if (configuration.Hitchtype === '5th Wheel Hitches') HitchtypePrice = 275;
  else if (configuration.Hitchtype === 'Gooseneck Hitches') HitchtypePrice = 320;


  // Hardware pricing
  if (configuration.Hardware === 'Pin') HardwarePrice = 45;
  else if (configuration.Hardware === 'Fasteners') HardwarePrice = 45;

  // Electrical pricing
  if (configuration.Electrical === 'Yes') ElectricalPrice = 120;
  else if (configuration.Electrical === 'No') ElectricalPrice = 0;


  // Safety kit pricing
  if (configuration.Safetykits === 'Yes') SafetykitsPrice = 90;
  else if (configuration.Safetykits === 'No') SafetykitsPrice = 0;

  // Covers pricing
  if (configuration.Covers === 'Yes') CoversPrice = 85;
  else if (configuration.Covers === 'No') CoversPrice = 0;

   // Wheelchocks pricing
  if (configuration.Wheelchocks === 'Yes') WheelchocksPrice = 70;
  else if (configuration.Wheelchocks === 'No') WheelchocksPrice = 0;

  // Extendedmirrors pricing
  if (configuration.Extendedmirrors === 'Yes') ExtendedmirrorsPrice = 115;
  else if (configuration.Extendedmirrors === 'No') ExtendedmirrorsPrice = 0;


  // ===== WRITE PRICES =====
  sheet.getCell('F18').value = HitchtypePrice;
  sheet.getCell('F19').value = HardwarePrice;
  sheet.getCell('F20').value = ElectricalPrice;
  sheet.getCell('F21').value = SafetykitsPrice;
  sheet.getCell('F22').value = CoversPrice;
  sheet.getCell('F23').value = WheelchocksPrice;
  sheet.getCell('F24').value = ExtendedmirrorsPrice;
}

  
  return await workbook.xlsx.writeBuffer();
}
const logFilePath = path.join(
  'C:',
  'Users',
  'DESKTOP-25',
  'Desktop',
  'Towing & Hitches CustomerLog.xlsx'
);


function appendCustomerLog(customer,quoteNo,productType) {
  try {
    let workbook;
    const sheetName = 'Quotes';

    // 1) Load or create workbook
    if (fs.existsSync(logFilePath)) {
      workbook = XLSX.readFile(logFilePath);
      console.log('📂 Opened existing log file');
    } else {
      workbook = XLSX.utils.book_new();
      console.log('🆕 Creating new log file');
    }

    // 2) Ensure Quotes sheet exists (create with headers if not)
    let worksheet = workbook.Sheets[sheetName];
    if (!worksheet) {
      const headers = [
        ['Name', 'Email', 'Dealer', 'Quote No', 'Product Type', 'Date']
      ];
      worksheet = XLSX.utils.aoa_to_sheet(headers);
      XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    }

    // 3) Remove default Sheet1 if present (keeps workbook clean)
    if (workbook.SheetNames.includes('Sheet1') && workbook.SheetNames.length > 1) {
      delete workbook.Sheets['Sheet1'];
      workbook.SheetNames = workbook.SheetNames.filter(n => n !== 'Sheet1');
    }

    // 4) Read current rows as array-of-arrays
    worksheet = workbook.Sheets[sheetName];
    let rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    console.log(`Rows before add: ${rows.length}`);

    // 5) Build values to append (match header order)
    const row = [
      customer?.name || '',
      customer?.email || '',
      customer?.dealer || '',
      quoteNo || '',
      productType || '',
      new Date().toLocaleString()
    ];

    // 6) Append and write back
    rows.push(row);
    worksheet = XLSX.utils.aoa_to_sheet(rows);
    workbook.Sheets[sheetName] = worksheet;

    XLSX.writeFile(workbook, logFilePath);
    console.log(`SAVING TO: ${logFilePath}`);
    console.log(`Rows after add: ${rows.length}`);
    console.log('✅ Log file updated successfully');
  } catch (err) {
    console.error('❌ appendCustomerLog error:', err);
    throw err;
  }
}

// Endpoint to receive customer data
app.post('/log-customer', (req, res) => {
  const customer = req.body;

  if (!customer.name || !customer.email || !customer.quote) {
    return res.status(400).json({ error: 'Missing required fields' });
  }

  appendCustomerLog(customer);
  res.json({ status: 'success' });
});
app.post('/api/send-quote', async (req, res) => {
  const { customer, configuration, productType } = req.body;

  if (!customer?.email || !customer?.name) {
    return res.status(400).json({ error: 'Missing customer details' });
  }


  try {
    const quoteNo = generateQuoteNumber();

    // ✅ Save to log after values are ready
    // ✅ Save to log after values are ready
await appendCustomerLog(customer, quoteNo, productType);


    const excelBuffer = await updateExcelWithNamedCells({
      customer,
      configuration,
      productType,
      quoteNo,
    });

    const transporter = nodemailer.createTransport({
      service: 'gmail',
      auth: {
        user: process.env.EMAIL_USER,
        pass: process.env.EMAIL_PASS
      }
    });

   await transporter.sendMail({
  from: `"Towing & Hitches Configurator" <${process.env.EMAIL_USER}>`,
  to: customer.email,
  subject: 'Your Equipment Quote',
  html: `
    <div style="font-family: Arial, sans-serif; color: #333; padding: 20px; line-height: 1.5;">
      <h2 style="color:#2c3e50;">Hello ${customer.name},</h2>

      <p>Thank you for configuring your <strong>${productType}</strong> with us.</p>

      <p><strong>Quote No:</strong> ${quoteNo}</p>
      <p>Please find your detailed quote attached to this email.</p>

      <p>Our team will get in touch with you shortly to discuss the next steps.</p>

      <hr style="margin: 30px 0; border: none; border-top: 1px solid #ccc;">

      <div style="text-align: center;">
        <p style="font-size: 12px; color: #777;">
          Jaydu, Inc.<br/>
          5th Floor, 504 A, PSR Prime Towers, DLF Rd, Gachibowli, Hyderabad, Telangana 500032, India<br/>
          Contact Us:  +91 -40-48577800
        </p>
        <img src="cid:companylogo" alt="Company Logo" style="max-height: 60px; margin-top: 10px;"/>
      </div>
    </div>
  `,
  attachments: [
    {
      filename: `Quote(${quoteNo}).xlsx`,
      content: excelBuffer,
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    },
    {
      filename: 'logo.png',
      path: path.join(__dirname, 'public/logo.png'),
      cid: 'companylogo'
    }
  ]
});

res.json({ message: 'Email sent successfully with Excel attachment' });

  } catch (error) {
    console.error('❌ Error:', error);
    res
      .status(500)
      .json({ error: 'Failed to send email or log customer' });
  }
});

app.listen(PORT, () => {
  console.log(`✅ Server started at http://localhost:${PORT}`);
});
