const express = require('express');
const bodyParser = require('body-parser');
const xlsx = require('xlsx');
const bcrypt = require('bcrypt');
const fs = require('fs');
const ExcelJS = require('exceljs');
const QRCode = require('qrcode'); // Import the QR code library
const app = express();
const port = 3000;

app.set('view engine', 'ejs');
app.use(express.json());
app.use(bodyParser.urlencoded({ extended: true }));



// Serve static files like QR codes
app.use(express.static('public'));

app.get('/', (req, res) => {
    if (!fs.existsSync('output.xlsx')) {
        const wb = xlsx.utils.book_new();
        const ws = xlsx.utils.json_to_sheet([]);
        xlsx.utils.book_append_sheet(wb, ws, 'Sheet1');
        xlsx.writeFile(wb, 'output.xlsx');
    }

    const workbook = xlsx.readFile('output.xlsx');
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(worksheet);

    res.render('index', { data: data });
});

app.get('/NewDocument.html', (req, res) => {
    res.render('NewDocument');
});

app.get('/signup', (req, res) => {
    res.render('NewUser');
});

// Path to the Excel file
const EXCEL_FILE = './users.xlsx';

// Signup route
app.post('/signup', async (req, res) => {
    const { name, email, password, department } = req.body;
  
    if (!name || !email || !password || !department) {
      return res.status(400).send('All fields are required.');
    }
  
    const validDepartments = ['admin', 'forensic', 'account', 'academics'];
    if (!validDepartments.includes(department)) {
      return res.status(400).send('Invalid department selected.');
    }
  
    try {
      const hashedPassword = await bcrypt.hash(password, 10);
  
      let workbook = new ExcelJS.Workbook();
      const worksheetColumns = [
        { header: 'Name', key: 'name', width: 20 },
        { header: 'Email', key: 'email', width: 30 },
        { header: 'Password', key: 'password', width: 40 },
        { header: 'Department', key: 'department', width: 20 },
      ];
  
      try {
        if (fs.existsSync(EXCEL_FILE)) {
          await workbook.xlsx.readFile(EXCEL_FILE);
        } else {
          const worksheet = workbook.addWorksheet('Users');
          worksheet.columns = worksheetColumns;
          await workbook.xlsx.writeFile(EXCEL_FILE);
        }
      } catch (error) {
        console.error('Error reading Excel file. Creating a new one.', error);
        workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Users');
        worksheet.columns = worksheetColumns;
        await workbook.xlsx.writeFile(EXCEL_FILE);
      }
  
      const worksheet = workbook.getWorksheet('Users');
  
      // Check if email already exists
      let emailExists = false;
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber > 1 && row.getCell(2).value === email) {
          emailExists = true;
        }
      });
  
      if (emailExists) {
        return res.status(400).send('Email already exists.');
      }
  
      // Add the new user data
      worksheet.addRow({
        name,
        email,
        password: hashedPassword,
        department,
      });
  
      await workbook.xlsx.writeFile(EXCEL_FILE);
  
      res.status(201).send('User registered successfully');
    } catch (error) {
      console.error('Error registering user:', error);
      res.status(500).send('Error registering user');
    }
    res.redirect('/');
  });
  
  

app.post('/submit', async (req, res) => {
    const formData = req.body;

    // Generate a unique document ID
    const documentId = `DOC-${Date.now()}`;
    formData.id = documentId;

    if (fs.existsSync('output.xlsx')) {
        const workbook = xlsx.readFile('output.xlsx');
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        const existingData = xlsx.utils.sheet_to_json(worksheet);
        existingData.push(formData);

        const updatedWorksheet = xlsx.utils.json_to_sheet(existingData);
        workbook.Sheets[sheetName] = updatedWorksheet;

        xlsx.writeFile(workbook, 'output.xlsx');
    } else {
        const workbook = xlsx.utils.book_new();
        const worksheet = xlsx.utils.json_to_sheet([formData]);
        xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
        xlsx.writeFile(workbook, 'output.xlsx');
    }

    // Generate QR Code
    const qrCodeData = `http://localhost:3000/EditDetail/${documentId}`;
    const qrCodePath = `public/qrcodes/${documentId}.png`;
    try {
        await QRCode.toFile(qrCodePath, documentId, {
            color: {
                dark: '#000000', // Black dots
                light: '#ffffff', // White background
            },
        });

        console.log(`QR Code generated: ${qrCodePath}`);
    } catch (err) {
        console.error('Error generating QR Code:', err);
    }

    // Redirect to the homepage
    res.redirect('/');
});

// Route to view the QR code for a specific document
app.get('/view-qr/:id', (req, res) => {
    const documentId = req.params.id;
    const workbook = xlsx.readFile('output.xlsx');
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(worksheet);

    // Find the document data by ID
    const documentDetails = data.find(doc => doc.id === documentId);

    if (documentDetails) {
        const qrCodePath = `qrcodes/${documentId}.png`;
        res.render('viewQr', { documentDetails, qrCodePath });
    } else {
        res.status(404).send('Document not found');
    }
});

app.get('/EditDetail/:id', (req, res) => {
    const documentId = req.params.id;

    // Read the Excel file
    const workbook = xlsx.readFile('output.xlsx');
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(worksheet);

    // Find the document by ID
    const documentDetails = data.find(doc => doc.id === documentId);

    if (documentDetails) {
        res.render('EditDetail', { documentDetails });
    } else {
        res.status(404).send('Document not found');
    }
});

app.post('/EditDetail/:id', (req, res) => {
    const documentId = req.params.id;
    const { newPlace, newTime, newDate } = req.body;

    // Read the Excel file
    const workbook = xlsx.readFile('output.xlsx');
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(worksheet);

    // Find and update the document
    const documentIndex = data.findIndex(doc => doc.id === documentId);

    if (documentIndex !== -1) {
        // Update the recent details with the new details
        data[documentIndex].recentPlace = newPlace || data[documentIndex].recentPlace;
        data[documentIndex].recentTime = newTime || data[documentIndex].recentTime;
        data[documentIndex].recentDate = newDate || data[documentIndex].recentDate;

        // Replace the old worksheet with the updated data
        const updatedWorksheet = xlsx.utils.json_to_sheet(data);
        workbook.Sheets[sheetName] = updatedWorksheet;

        // Write the updated workbook back to the file
        xlsx.writeFile(workbook, 'output.xlsx');

        res.end('all data is updated\n');
    } else {
        res.status(404).send('Document not found');
    }
});
app.get('/login', (req, res) => {
    res.render('login');
});

app.post('/login', async (req, res) => {
    const { email, password, department } = req.body;
  
    // Validate input
    if (!email || !password || !department) {
      return res.status(400).send('All fields are required.');
    }
  
    // Departments allowed
    const validDepartments = ['admin', 'forensic', 'account', 'academics'];
    if (!validDepartments.includes(department)) {
      return res.status(400).send('Invalid department selected.');
    }
  
    try {
      const workbook = new ExcelJS.Workbook();
  
      if (!fs.existsSync(EXCEL_FILE)) {
        return res.status(404).send('No users found. Please sign up first.');
      }
  
      await workbook.xlsx.readFile(EXCEL_FILE);
      const worksheet = workbook.getWorksheet('Users');
  
      // Search for the user in the Excel file
      let userFound = null;
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber > 1) {
          const rowEmail = row.getCell(2).value;
          const rowPassword = row.getCell(3).value;
          const rowDepartment = row.getCell(4).value;
  
          if (rowEmail === email && rowDepartment === department) {
            userFound = { email: rowEmail, password: rowPassword, department: rowDepartment };
          }
        }
      });
  
      if (!userFound) {
        return res.status(401).send('Invalid email or department.');
      }
  
      // Compare the hashed password
      const isPasswordValid = await bcrypt.compare(password, userFound.password);
      if (!isPasswordValid) {
        return res.status(401).send('Invalid password.');
      }
  
      // Redirect to the department page
      switch (department) {
        case 'admin':
          return res.status(200).send('Redirecting to Admin Page');
        case 'forensic':
          return res.status(200).send('Redirecting to Forensic Page');
        case 'account':
          return res.status(200).send('Redirecting to Account Page');
        case 'academics':
          return res.status(200).send('Redirecting to Academics Page');
        default:
          return res.status(400).send('Invalid department.');
      }
    } catch (error) {
      console.error('Error logging in:', error);
      res.status(500).send('Error logging in.');
    }
  });
  

app.listen(port, () => {
    console.log(`Server is running on http://localhost:${port}`);
});
