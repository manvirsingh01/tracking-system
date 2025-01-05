const express = require('express');
const bodyParser = require('body-parser');
const xlsx = require('xlsx');
const bcrypt = require('bcrypt');
const fs = require('fs');
const ExcelJS = require('exceljs');
const QRCode = require('qrcode');
const path = require('path');

const app = express();
const port = 3000;

// Constants
const EXCEL_FILE = './users.xlsx';
const OUTPUT_FILE = './output.xlsx';
const LOGS_DIR = path.join(__dirname, 'logs');
const QR_CODES_DIR = path.join(__dirname, 'public/qrcodes');
const VALID_DEPARTMENTS = ['admin', 'forensic', 'account', 'academics'];

// Middleware
app.set('view engine', 'ejs');
app.use(express.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static('public'));

// Utility Functions
const ensureFileExists = (filePath, worksheetColumns = null) => {
  if (!fs.existsSync(filePath)) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Sheet1');
    if (worksheetColumns) worksheet.columns = worksheetColumns;
    workbook.xlsx.writeFile(filePath).catch(err => {
      console.error(`Error creating file at ${filePath}:`, err);
    });
  }
};

const readExcelData = async (filePath, sheetName = 'Sheet1') => {
  if (!fs.existsSync(filePath)) {
    console.error(`File not found: ${filePath}`);
    return [];
  }

  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.readFile(filePath);
  } catch (err) {
    console.error(`Error reading file ${filePath}:`, err);
    return [];
  }

  const worksheet = workbook.getWorksheet(sheetName);
  if (!worksheet) {
    console.error(`Sheet "${sheetName}" not found in file ${filePath}`);
    return [];
  }

  const headers = worksheet.getRow(1).values.slice(1);
  const rows = worksheet.getSheetValues().slice(2);

  return rows.map(row => {
    const rowData = {};
    headers.forEach((header, index) => {
      rowData[header] = row[index + 1] || '';
    });
    return rowData;
  });
};

const writeExcelData = async (filePath, data, sheetName = 'Sheet1') => {
  const workbook = new ExcelJS.Workbook();

  // If the file exists, load it; otherwise, create a new workbook
  if (fs.existsSync(filePath)) {
    try {
      await workbook.xlsx.readFile(filePath);
    } catch (err) {
      console.error(`Error reading file ${filePath} for writing:`, err);
      return;
    }
  }

  // Get the worksheet or create a new one
  let worksheet = workbook.getWorksheet(sheetName);
  if (!worksheet) {
    worksheet = workbook.addWorksheet(sheetName);
  } else {
    // Clear all rows from the worksheet if it exists
    const rowCount = worksheet.rowCount;
    for (let i = rowCount; i > 0; i--) {
      worksheet.spliceRows(i, 1);
    }
  }

  // Add data to the worksheet
  if (data.length > 0) {
    worksheet.columns = Object.keys(data[0]).map(key => ({ header: key, key }));
    worksheet.addRows(data);
  }

  // Write the updated workbook to the file
  try {
    await workbook.xlsx.writeFile(filePath);
  } catch (err) {
    console.error(`Error writing to file ${filePath}:`, err);
  }
};


const updateLogFile = async (logFilePath, logEntry) => {
  const sheetName = 'Sheet1';
  let workbook = new ExcelJS.Workbook();

  // Ensure the file exists or create a new one
  if (fs.existsSync(logFilePath)) {
    try {
      await workbook.xlsx.readFile(logFilePath);
    } catch (err) {
      console.error(`Error reading log file ${logFilePath}:`, err);
    }
  }

  let worksheet = workbook.getWorksheet(sheetName);
  if (!worksheet) {
    worksheet = workbook.addWorksheet(sheetName);
    worksheet.columns = [
      { header: 'Action', key: 'Action' },
      { header: 'Date', key: 'Date' },
      { header: 'InTime', key: 'InTime' },
      { header: 'Place', key: 'Place' },
      { header: 'OutTime', key: 'OutTime' },
      { header: 'PreviousPlace', key: 'PreviousPlace' },
      { header: 'PreviousTime', key: 'PreviousTime' },
      { header: 'PreviousDate', key: 'PreviousDate' },
    ];
  }

  // Add the log entry
  worksheet.addRow({
    Action: logEntry.Action || '',
    Date: logEntry.Date || '',
    InTime: logEntry.InTime || '',
    Place: logEntry.Place || '',
    OutTime: logEntry.OutTime || '',
    PreviousPlace: logEntry.PreviousPlace || '',
    PreviousTime: logEntry.PreviousTime || '',
    PreviousDate: logEntry.PreviousDate || ''
  });

  // Unhide all rows and columns
  worksheet.eachRow((row, rowNumber) => {
    row.hidden = false; // Ensure rows are not hidden
  });

  worksheet.columns.forEach(column => {
    column.hidden = false; // Ensure columns are not hidden
  });

  try {
    // Write to the Excel file
    await workbook.xlsx.writeFile(logFilePath);
    console.log('Log entry added successfully to the file.');
  } catch (err) {
    console.error(`Error writing log file ${logFilePath}:`, err);
  }
};



async function generateQRCode(data, path) {
  try {
      await QRCode.toFile(path, data, {
          color: {
              dark: '#000000', // QR code color
              light: '#FFFFFF', // Background color
          },
      });
      console.log('QR Code generated successfully:', path);
  } catch (err) {
      console.error('Error generating QR Code:', err);
      throw err;
  }
}

// Routes
app.get('/', async (req, res) => {
  ensureFileExists(OUTPUT_FILE);
  const data = await readExcelData(OUTPUT_FILE);
  res.render('index', { data });
});

app.get('/signup', (req, res) => res.render('NewUser'));
app.get('/NewDocument.html', (req, res) => res.render('NewDocument'));

app.post('/signup', async (req, res) => {
  try {
    const { name, email, password, department } = req.body;

    // Log received data for debugging
    console.log('Received signup data:', req.body);

    // Validate the input fields
    if (!name || !email || !password || !VALID_DEPARTMENTS.includes(department)) {
      console.log('Validation failed: Missing or invalid data.');
      return res.status(400).send('Invalid input or department.');
    }

    // Ensure the Excel file exists with the required headers
    ensureFileExists(EXCEL_FILE, [
      { header: 'Name', key: 'name', width: 20 },
      { header: 'Email', key: 'email', width: 30 },
      { header: 'Password', key: 'password', width: 40 },
      { header: 'Department', key: 'department', width: 20 },
    ]);

    // Read existing data from the Excel file
    const users = await readExcelData(EXCEL_FILE);
    console.log('Current users in file:', users);

    // Check for existing email
    if (users.some(user => user.email === email)) {
      console.log('Duplicate email detected:', email);
      return res.status(400).send('Email already exists.');
    }

    // Hash the password for secure storage
    const hashedPassword = await bcrypt.hash(password, 10);
    console.log('Hashed password generated.');

    // Append new user data to the list
    const newUser = { name, email, password: hashedPassword, department };
    users.push(newUser);

    // Write updated data back to the Excel file
    await writeExcelData(EXCEL_FILE, users);
    console.log('User added successfully:', newUser);

    // Redirect to home page
    res.redirect('/');
  } catch (err) {
    console.error('Error during signup process:', err);
    res.status(500).send('An error occurred during signup.');
  }
});



app.get('/department/:department', async (req, res) => {
  const { department } = req.params;

  // Check if the department is valid
  if (!VALID_DEPARTMENTS.includes(department)) {
    return res.status(404).send('Department not found.');
  }

  // Ensure the file exists and read data
  ensureFileExists(OUTPUT_FILE);
  const data = await readExcelData(OUTPUT_FILE);

  // Filter data for the requested department based on 'recentPlace'
  const filteredData = data.filter(
    item => item.recentPlace && item.recentPlace.toLowerCase() === department.toLowerCase()
  );

  // Render the department page with filtered data
  res.render('Department', {
    department,
    title: `Welcome to the ${department.charAt(0).toUpperCase() + department.slice(1)} Department`,
    description: `This is the page for the ${department.charAt(0).toUpperCase() + department.slice(1)} department.`,
    data: filteredData, // Pass only relevant department data
  });
});



app.post('/submit', async (req, res) => {
  const formData = req.body;
  formData.id = `DOC-${Date.now()}`;

  ensureFileExists(OUTPUT_FILE);
  const data = await readExcelData(OUTPUT_FILE);
  data.push(formData);
  await writeExcelData(OUTPUT_FILE, data);

  const qrCodePath = path.join(QR_CODES_DIR, `${formData.id}.png`);
  await generateQRCode(`http://localhost:3000/EditDetail/${formData.id}`, qrCodePath);

  const logFilePath = path.join(LOGS_DIR, `${formData.id}.xlsx`);
  await updateLogFile(logFilePath, {
    Date: new Date().toLocaleDateString(),
    InTime: new Date().toLocaleTimeString(),
    Place: formData.place || 'Unknown',
    OutTime: '',
  });

  res.redirect('/');
});

app.get('/EditDetail/:id', async (req, res) => {
  const { id } = req.params;

  ensureFileExists(OUTPUT_FILE);
  const data = await readExcelData(OUTPUT_FILE);
  const documentDetails = data.find(doc => doc.id === id);

  if (!documentDetails) {
    return res.status(404).send('Document not found.');
  }

  res.render('EditDetail', { documentDetails });
});
app.post('/EditDetail/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const { action, place, date, time } = req.body;

    // Ensure the main data file exists
    ensureFileExists(OUTPUT_FILE);
    const data = await readExcelData(OUTPUT_FILE);
    const documentIndex = data.findIndex(doc => doc.id === id);

    if (documentIndex === -1) {
      return res.status(404).send('Document not found.');
    }

    const logFilePath = path.join(LOGS_DIR, `${id}.xlsx`);
    const document = data[documentIndex];

    // Save previous state for logging
    const previousState = {
      Place: document.recentPlace || 'Unknown',
      Time: document.recentTime || 'Unknown',
      Date: document.recentDate || 'Unknown',
    };

    // Update the document with new values
    document.recentPlace = place || previousState.Place;
    document.recentTime = time || new Date().toLocaleTimeString();
    document.recentDate = date || new Date().toLocaleDateString();

    // Define log entry based on action
    const logEntry = {
      Action: action || 'Update',
      Date: document.recentDate,
      InTime: document.recentTime,
      Place: document.recentPlace,
      OutTime: '',
      PreviousPlace: previousState.Place,
      PreviousTime: previousState.Time,
      PreviousDate: previousState.Date,
    };

    // Update the log file
    try {
      await updateLogFile(logFilePath, logEntry);
      console.log('Log updated successfully:', logEntry);
    } catch (err) {
      console.error('Error updating log file:', err);
    }

    // Save the updated document back to the main file
    await writeExcelData(OUTPUT_FILE, data);
    console.log('Document updated successfully:', document);

    res.redirect(`/EditDetail/${id}`);
  } catch (err) {
    console.error('Error processing EditDetail route:', err);
    res.status(500).send('Internal server error.');
  }
});



app.get('/view-qr/:id', async (req, res) => {
  try {
    const { id } = req.params;

    ensureFileExists(OUTPUT_FILE);
    const data = await readExcelData(OUTPUT_FILE);
    const documentDetails = data.find(doc => doc.id === id);

    if (!documentDetails) {
      return res.status(404).send('Document not found.');
    }

    const qrCodePath = path.join(QR_CODES_DIR, `${id}.png`);
    const qrCodeExists = fs.existsSync(qrCodePath);

    if (!qrCodeExists) {
      return res.status(404).send('QR Code not found.');
    }

    const logFilePath = path.join(LOGS_DIR, `${id}.xlsx`);
    const logData = fs.existsSync(logFilePath) ? await readExcelData(logFilePath) : [];

    res.render('viewQr', { documentDetails, qrCodePath, logData });
  } catch (error) {
    console.error('Error in /view-qr route:', error.message);
    res.status(500).send('An error occurred while processing your request.');
  }
});


app.get('/login', (req, res) => res.render('login'));

app.post('/login', async (req, res) => {
  const { email, password, department } = req.body;

  if (!email || !password || !VALID_DEPARTMENTS.includes(department)) {
    return res.status(400).send('Invalid login credentials.');
  }

  ensureFileExists(EXCEL_FILE);
  const users = await readExcelData(EXCEL_FILE);
  const user = users.find(user => user.Email === email && user.Department === department);

  if (!user || !(await bcrypt.compare(password, user.Password))) {
    return res.status(401).send('Invalid email, department, or password.');
  }
  res.redirect(`/department/${department}`);
});

// Start Server
app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
});
