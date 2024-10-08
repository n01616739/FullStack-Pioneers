require('dotenv').config(); // Load environment variables from .env
const express = require('express');
const nodemailer = require('nodemailer');
const fs = require('fs');
const xlsx = require('xlsx');
const path = require('path');
const multer = require('multer'); // Import multer

const app = express();
const PORT = process.env.PORT || 3000;

// Set up Multer for handling multipart/form-data
const upload = multer(); // No file uploads, just text fields in the form

// Middleware
app.use(express.static("./public/")); // Serve static files

// Set up Nodemailer transport
const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
        user: process.env.EMAIL_USER,  // Load from environment
        pass: process.env.EMAIL_PASS   // Load from environment
    }
});

// Serve the main HTML file
app.get("/", (req, res) => {
    res.sendFile(path.join(__dirname, "/public/index.html"));
});

// Form submission route (using multer to handle form data)
app.post('/submit-form', upload.none(), (req, res) => {
    const { companyName, contactName, email, phone } = req.body;

    // Log received form data for debugging
    console.log('Received data:', { companyName, contactName, email, phone });

    // Validate email field
    if (!email) {
        return res.status(400).send('Email is required.');
    }

    // Send resumes via email
    const mailOptions = {
        from: process.env.EMAIL_USER,
        to: email,
        subject: 'Resumes from Event',
        text: `Thank you for attending our event! Here are the resumes.`,
        attachments: [
            { filename: 'resume1.pdf', path: './resumes/resume1.pdf' },
            { filename: 'resume2.pdf', path: './resumes/resume2.pdf' },
            { filename: 'resume3.pdf', path: './resumes/resume3.pdf' },
            { filename: 'resume4.pdf', path: './resumes/resume4.pdf' }
        ]
    };

    transporter.sendMail(mailOptions, (error, info) => {
        if (error) {
            console.error('Error sending email:', error);
            return res.status(500).send('Error sending email: ' + error.toString());
        }

        // Store sponsor data in Excel file
        const workbook = xlsx.utils.book_new();
        const filePath = './sponsor_data.xlsx';

        let existingData = [];
        if (fs.existsSync(filePath)) {
            const existingWorkbook = xlsx.readFile(filePath);
            existingData = xlsx.utils.sheet_to_json(existingWorkbook.Sheets[existingWorkbook.SheetNames[0]]);
        }

        const newData = { 
            'Company Name': companyName, 
            'Contact Name': contactName, 
            'Email': email, 
            'Phone': phone 
        };

        existingData.push(newData);
        const worksheet = xlsx.utils.json_to_sheet(existingData);
        xlsx.utils.book_append_sheet(workbook, worksheet, 'Sponsors');
        xlsx.writeFile(workbook, filePath);

        res.send('Form submitted successfully! Emails sent to ' + email);
    });
});

// Start server
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
