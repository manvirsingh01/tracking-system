const express = require('express');
const bodyParser = require('body-parser');
const xlsx = require('xlsx');
const fs = require('fs');
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


app.listen(port, () => {
    console.log(`Server is running on http://localhost:${port}`);
});
