require('dotenv').config(); // Load environment variables from .env
const express = require('express');
const { google } = require('googleapis');
const nodemailer = require('nodemailer');
const fs = require('fs');
const xlsx = require('xlsx');
const path = require('path');
const WebSocket = require('ws');
const multer = require('multer'); // Import multer to handle form data
const { OAuth2 } = google.auth;

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware to serve static files
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// WebSocket server setup
const wss = new WebSocket.Server({ port: 8080 });
let clients = [];
let processedEmailIds = []; // Store the processed email message IDs to avoid duplicates

// When a client connects via WebSocket
wss.on('connection', (ws) => {
    clients.push(ws);
    console.log('New client connected');

    ws.on('close', () => {
        clients = clients.filter(client => client !== ws);
        console.log('Client disconnected');
    });
});

// Function to notify clients via WebSocket
function notifyClients(message) {
    clients.forEach(client => {
        client.send(JSON.stringify(message));
    });
}

// OAuth2 setup for Gmail API
const oAuth2Client = new OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    process.env.REDIRECT_URI
);

// Set OAuth2 credentials
oAuth2Client.setCredentials({
    refresh_token: process.env.GOOGLE_REFRESH_TOKEN
});

// Gmail API setup
const gmail = google.gmail({ version: 'v1', auth: oAuth2Client });

// Function to check for new emails
async function checkForEmails() {
    try {
        const res = await gmail.users.messages.list({
            userId: 'me',
            q: 'is:unread', // Search for unread emails
            maxResults: 10,
        });

        const messages = res.data.messages || [];
        if (messages.length) {
            for (const message of messages) {
                // Check if this email has already been processed
                if (processedEmailIds.includes(message.id)) {
                    continue; // Skip if the email has already been processed
                }

                const msg = await gmail.users.messages.get({
                    userId: 'me',
                    id: message.id,
                });

                const emailData = msg.data;
                const headers = emailData.payload.headers;
                const fromHeader = headers.find(h => h.name === 'From') || {};
                const subjectHeader = headers.find(h => h.name === 'Subject') || {};
                const from = fromHeader.value;
                const subject = subjectHeader.value;

                // Get sponsor emails from the Excel sheet
                const sponsorEmails = getSponsorEmails();

                // Check if the email is from a sponsor
                if (sponsorEmails.some(email => from.includes(email))) {
                    const snippet = emailData.snippet;

                    // Notify WebSocket clients with email details
                    notifyClients({
                        sender: from,
                        subject: subject,
                        snippet: snippet
                    });

                    // Mark the email as processed to avoid duplicates
                    processedEmailIds.push(message.id);
                }
            }
        }
    } catch (error) {
        console.error('Error checking for emails:', error);
    }
}

// Function to read sponsor emails from the Excel file
function getSponsorEmails() {
    const filePath = './sponsor_data.xlsx'; // Path to your Excel file
    if (!fs.existsSync(filePath)) {
        return [];
    }
    const workbook = xlsx.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const sponsors = xlsx.utils.sheet_to_json(sheet);
    return sponsors.map(sponsor => sponsor.Email); // Extract emails
}

// Check for new sponsor emails every 1 second
setInterval(checkForEmails, 1000); // Check for new emails every second

// ---------------- Form Submission and Email Sending -----------------

// Setup Multer for handling multipart/form-data
const upload = multer();

// Setup Nodemailer for sending emails with resumes attached
const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
        user: process.env.EMAIL_USER,  // Your email
        pass: process.env.EMAIL_PASS   // Your email password or app-specific password
    }
});

// Route to handle form submission and send email with attachments
app.post('/submit-form', upload.none(), (req, res) => {
    const { companyName, contactName, email, phone } = req.body;

    // Email options to send resumes
    const mailOptions = {
        from: process.env.EMAIL_USER,
        to: email, // Email to the sponsor
        subject: `Collaboration with ${companyName}`,
        text: `Dear ${contactName},\n\nThank you for your interest in collaboration with our project team. Attached are the resumes of our team members for your consideration.\n\nBest regards,\nFullstack Pioneers Team`,
        attachments: [
            { filename: 'resume1.pdf', path: './resumes/resume1.pdf' },
            { filename: 'resume2.pdf', path: './resumes/resume2.pdf' },
            { filename: 'resume3.pdf', path: './resumes/resume3.pdf' },
            { filename: 'resume4.pdf', path: './resumes/resume4.pdf' }
        ]
    };

    // Send the email with attached resumes
    transporter.sendMail(mailOptions, (error, info) => {
        if (error) {
            console.error('Error sending email:', error);
            return res.status(500).send('Error sending email: ' + error.toString());
        }
        console.log('Email sent: ' + info.response);
        res.send('Form submitted successfully! Email sent to ' + email);
    });
});

// Start the server
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
