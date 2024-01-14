const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
// const pdf = require('html-pdf');
const nodemailer = require('nodemailer');
// const path = require('path');
const puppeteer = require('puppeteer');
const { MongoClient } = require('mongodb');
const pdf2img=require('pdf2img');

const app = express();
const port = 3005; 

app.set('view engine', 'ejs');

const storage = multer.memoryStorage();
const upload = multer({ storage: storage });


const mongoURI = 'mongodb+srv://karuna0741be21:kannu@cluster0.paefxi1.mongodb.net/?retryWrites=true&w=majority';

const client = new MongoClient(mongoURI);

app.get('/', (req, res) => {
  res.render('upload');
});

app.post('/upload', upload.single('excelFile'), async (req, res) => {
  try {
    
    await client.connect();
    console.log('Connected to MongoDB');

    const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const jsonData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

    const transporter = nodemailer.createTransport({
      service: 'gmail',
      auth: {
        user: 'karuna0741.be21@chitkara.edu.in',
        pass: 'K@nnu22Y@shu24@',
      },
    });

    function isValidEmail(email) {
      return true;
    }

    const browser = await puppeteer.launch({ headless: 'new' });


    for (const data of jsonData) {
      const { id, name, email, mobile_number, amount, number_of_trees } = data;

      console.log(`Checking record with ID ${id}: ${email}`);

      const certificateHTML = await new Promise((resolve, reject) => {
        res.render('certificate', { name, id, mobile_number, email, amount, number_of_trees }, (err, html) => {
          if (err) {
            console.error('Error rendering certificate:', err);
            reject(err);
          } else {
            resolve(html);
          }
        });
      });

      const page = await browser.newPage();
      await page.setContent(certificateHTML);

      const pdfBuffer = await page.pdf({
         format: 'Letter', 
         landscape: true, 
         printBackground: true, 
         margin: { top: 0, bottom: 0, left: 0, right: 0 },
          preferCSSPageSize: true 
        });

      if (!email || !isValidEmail(email)) {
        console.error(`Invalid or missing email for record with ID ${id}: ${email}`);
        continue;
      }

      try {
        const mailOptions = {
          from: 'karuna0741.be21@chitkara.edu.in',
          to: email,
          subject: '🌟 Thank You for Your Generous Donation 🌳',
          text: `
Dear ${name}🙏,
We're deeply grateful for your Rs${amount} donation towards our tree-planting initiative. Your commitment to sustainability inspires us♥.
Attached is your Certificate of Contribution, recognizing your significant impact. Your generosity shapes a greener future🏆.
For questions or updates, feel free to reach out. Thank you, ${name}, for being a vital part of our mission🌍.
Warm regards,
Karuna's Farm🌱
          `,
          attachments: [
            {
              filename: 'certificate.pdf',
              content: pdfBuffer,
              encoding: 'base64',
            },
          ],
        };

        const info = await transporter.sendMail(mailOptions);
        console.log(`Email sent for record with ID ${id}:`, info.response);
      } catch (mailErr) {
        console.error(`Error sending email for record with ID ${id}:`, mailErr);
      }

      await page.close();
    }

    await browser.close();

    res.status(200).send('Processing completed');
  } catch (error) {
    console.error('Error processing and storing data in MongoDB:', error);
    res.status(500).send('Internal Server Error');
  } finally {
    
    await client.close();
  }
});


app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});