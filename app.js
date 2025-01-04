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
    workbook.xlsx.writeFile(filePath);
  }
};

const readExcelData = async (filePath, sheetName = 'Sheet1') => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const worksheet = workbook.getWorksheet(sheetName);
  const rows = worksheet.getSheetValues().slice(2);
  const headers = worksheet.getRow(1).values.slice(1);

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
  const worksheet = workbook.addWorksheet(sheetName);
  if (data.length > 0) worksheet.columns = Object.keys(data[0]).map(key => ({ header: key, key }));
  worksheet.addRows(data);
  await workbook.xlsx.writeFile(filePath);
};

const generateQRCode = async (data, filePath) => {
  await QRCode.toFile(filePath, data, {
    color: { dark: '#000000', light: '#ffffff' },
  });
};

const updateLogFile = async (logFilePath, logEntry) => {
  const workbook = fs.existsSync(logFilePath) ? xlsx.readFile(logFilePath) : new ExcelJS.Workbook();
  const sheetName = workbook.SheetNames ? workbook.SheetNames[0] : 'Sheet1';
  const worksheet = workbook.Sheets[sheetName] || workbook.addWorksheet(sheetName);
  const existingData = xlsx.utils.sheet_to_json(worksheet);

  existingData.push(logEntry);
  const updatedSheet = xlsx.utils.json_to_sheet(existingData);
  workbook.Sheets[sheetName] = updatedSheet;
  xlsx.writeFile(workbook, logFilePath);
};

// Routes
app.get('/', async (req, res) => {
  ensureFileExists(OUTPUT_FILE);
  const data = await readExcelData(OUTPUT_FILE);
  res.render('index', { data });
});

app.get('/signup', (req, res) => res.render('NewUser'));

app.post('/signup', async (req, res) => {
  const { name, email, password, department } = req.body;

  if (!name || !email || !password || !VALID_DEPARTMENTS.includes(department)) {
    return res.status(400).send('Invalid input or department.');
  }

  ensureFileExists(EXCEL_FILE, [
    { header: 'Name', key: 'name', width: 20 },
    { header: 'Email', key: 'email', width: 30 },
    { header: 'Password', key: 'password', width: 40 },
    { header: 'Department', key: 'department', width: 20 },
  ]);

  const users = await readExcelData(EXCEL_FILE);
  if (users.find(user => user.Email === email)) {
    return res.status(400).send('Email already exists.');
  }

  const hashedPassword = await bcrypt.hash(password, 10);
  users.push({ Name: name, Email: email, Password: hashedPassword, Department: department });
  await writeExcelData(EXCEL_FILE, users);

  res.redirect('/');
});

app.get('/department/:department', async (req, res) => {
  const { department } = req.params;

  if (!VALID_DEPARTMENTS.includes(department)) {
    return res.status(404).send('Department not found.');
  }

  ensureFileExists(OUTPUT_FILE);
  const data = await readExcelData(OUTPUT_FILE);
  res.render('Department', {
    department,
    title: `Welcome to the ${department.charAt(0).toUpperCase() + department.slice(1)} Department`,
    description: `This is the page for the ${department.charAt(0).toUpperCase() + department.slice(1)} department.`,
    data,
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
  const { id } = req.params;
  const { action, newPlace, newDate, newTime, recentPlace, recentTime, recentDate } = req.body;

  ensureFileExists(OUTPUT_FILE);
  const data = await readExcelData(OUTPUT_FILE);
  const documentIndex = data.findIndex(doc => doc.id === id);

  if (documentIndex === -1) {
    return res.status(404).send('Document not found.');
  }

  const logFilePath = path.join(LOGS_DIR, `${id}.xlsx`);

  if (action === 'receive') {
    data[documentIndex] = { ...data[documentIndex], recentPlace, recentTime, recentDate };
    await updateLogFile(logFilePath, {
      Action: 'Receive',
      Date: recentDate || new Date().toLocaleDateString(),
      InTime: recentTime || new Date().toLocaleTimeString(),
      Place: recentPlace || 'Unknown',
      OutTime: '',
    });
  } else if (action === 'forward') {
    data[documentIndex] = { ...data[documentIndex], newPlace, newTime, newDate };
    await updateLogFile(logFilePath, {
      Action: 'Forward',
      Date: newDate || new Date().toLocaleDateString(),
      InTime: '',
      Place: newPlace || 'Unknown',
      OutTime: newTime || new Date().toLocaleTimeString(),
    });
  }

  await writeExcelData(OUTPUT_FILE, data);
  res.redirect(`/EditDetail/${id}`);
});

app.get('/view-qr/:id', async (req, res) => {
  const { id } = req.params;

  ensureFileExists(OUTPUT_FILE);
  const data = await readExcelData(OUTPUT_FILE);
  const documentDetails = data.find(doc => doc.id === id);

  if (!documentDetails) {
    return res.status(404).send('Document not found.');
  }

  const qrCodePath = path.join(QR_CODES_DIR, `${id}.png`);
  const logFilePath = path.join(LOGS_DIR, `${id}.xlsx`);
  const logData = fs.existsSync(logFilePath) ? await readExcelData(logFilePath) : [];

  res.render('viewQr', { documentDetails, qrCodePath, logData });
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
