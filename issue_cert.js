const express = require('express');
const nodemailer = require('nodemailer');
const admin = require('firebase-admin');
const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');
const docxtemplater = require('docxtemplater');
const PizZip = require('pizzip');
const cors = require('cors');
const serviceAccount = require('./firebaseConfig.json');

const libreOfficePath = 'C:\\Program Files\\LibreOffice\\program\\soffice';

const convertToPdf = (docPath, pdfPath) => {
  return new Promise((resolve, reject) => {
    exec(`"${libreOfficePath}" --headless --convert-to pdf "${docPath}" --outdir "${path.dirname(pdfPath)}"`, (error, stdout, stderr) => {
      if (error) {
        reject(`Error converting DOCX to PDF: ${stderr}`);
      } else {
        resolve(stdout);
      }
    });
  });
};

admin.initializeApp({
  credential: admin.credential.cert(serviceAccount),
});

const db = admin.firestore();
const app = express();
app.use(express.json());
app.use(cors());

const sendEmail = async (toEmail, subject, body, attachmentPath) => {
  const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
      user: 'sardar.hussyn@gmail.com',
      pass: 'axmm aayj djsd cqql',
    },
  });

  const mailOptions = {
    from: 'sardar.hussyn@gmail.com',
    to: toEmail,
    subject,
    text: body,
    attachments: [{ path: attachmentPath }],
  };

  try {
    await transporter.sendMail(mailOptions);
    console.log(`Email sent successfully to ${toEmail}`);
  } catch (error) {
    console.error('Error sending email:', error);
    throw new Error('Failed to send email.');
  }
};

const generateAndSendCertificate = async (userId) => {
  const userDoc = await db.collection('users').doc(userId).get();
  if (!userDoc.exists) {
    throw new Error(`User with ID ${userId} not found.`);
  }

  const userData = userDoc.data();
  const name = userData.name || 'Unknown Name';
  console.log(`Fetched user name: ${name}`);
  
  const email = userData.email;

  const templatePath = path.join(__dirname, 'Certificate.docx');
  const docxData = fs.readFileSync(templatePath);

  const zip = new PizZip(docxData);
  const doc = new docxtemplater(zip);

  doc.setData({ name });

  console.log(`Setting 'name' in template: ${name}`);

  try {
    doc.render();
    console.log('Template rendered successfully.');
  } catch (error) {
    console.error('Error rendering template:', error);
    throw new Error('Failed to render the template.');
  }

  const buffer = doc.getZip().generate({ type: 'nodebuffer' });
  const docPath = path.join(__dirname, 'Output', 'Doc', `${name}_Certificate.docx`);
  fs.mkdirSync(path.dirname(docPath), { recursive: true });
  fs.writeFileSync(docPath, buffer);

  const pdfPath = path.join(__dirname, 'Output', 'PDF', `${name}_Certificate.pdf`);
  fs.mkdirSync(path.dirname(pdfPath), { recursive: true });

  try {
    await convertToPdf(docPath, pdfPath);
    console.log('Converted to PDF successfully.');
    await sendEmail(email, 'Your Certificate', 'Attached is your certificate.', pdfPath);
    console.log(`Certificate sent to ${email}`);
  } catch (error) {
    console.error('Error during certificate generation or email sending:', error);
    throw error;
  }

  return { status: 'success', message: `Certificate sent to ${email}.` };
};

app.post('/issue_cert', async (req, res) => {
  try {
    const { user_id } = req.body;
    if (!user_id) {
      return res.status(400).json({ status: 'error', message: 'User ID is required.' });
    }

    const result = await generateAndSendCertificate(user_id);
    res.status(200).json(result);
  } catch (error) {
    console.error('Error during certificate issuance:', error);
    res.status(500).json({ status: 'error', message: error.message });
  }
});

app.listen(5000, () => {
  console.log('Server running on http://localhost:5000');
});
