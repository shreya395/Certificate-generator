const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const mongoose = require('mongoose');
const nodemailer = require('nodemailer');
const pdf = require('html-pdf');
const ejs = require('ejs');
const app = express();


mongoose.connect('MONGODBURL').then(() => {
  console.log("Database connected")
})
const db = mongoose.connection;

const dataSchema = new mongoose.Schema({
  name: String,
  email_id: String,
  phone_no: Number,
  trees_donated: Number,
  amount: Number,
});

const Data = mongoose.model('Data', dataSchema);

const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

app.set("view engine", "ejs");

app.get('/', (req, res) => {
  res.render('upload');
})

app.post('/upload', upload.single('file'), async (req, res) => {
  const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const jsonData = xlsx.utils.sheet_to_json(sheet);

  
  const processedData = jsonData.map(item => {
    const treesDonated = Math.floor(item.amount / 100); // Calculate trees donated
    return {
      name: item.name,
      email_id: item.email_id,
      phone_no: item.phone_no,
      amount: item.amount,
      trees_donated: treesDonated
    };
  });

  try {
    await Data.insertMany(processedData);

   
    processedData.forEach(async (item) => {
      
      ejs.renderFile('certificate.ejs', { item }, async (err, htmlTemplate) => {
        if (err) {
          console.error(err);
          return res.status(500).send('Error rendering EJS template');
        }

       
        pdf.create(htmlTemplate).toFile(`certificate_${ item.name }.pdf`, async function (pdfErr, pdfPath) {
          if (pdfErr) {
            console.error(pdfErr);
            return res.status(500).send('Error generating PDF');
          }

          
          let transporter = nodemailer.createTransport({
           
            service: 'Gmail',
            auth: {
              user: 'AVC@gmail.com',
              pass: 'Password'
            }
          });

          // Email options
          let mailOptions = {
            from: 'shreya1330.be21@chitkara.edu.in',
            to: item.email_id,
            subject: 'Certificate of Donation',
            text: 'Certificate attached.',
            attachments: [{
              filename: `certificate_${ item.name }.pdf`,
              path: pdfPath.filename

            }]
        };

        // Send email with attachment
        try {
          await transporter.sendMail(mailOptions);
          console.log(`Certificate sent to ${item.email_id}`);
        } catch (emailErr) {
          console.error(emailErr);
        }
      });
    });
  });

res.status(200).send('Data uploaded successfully');
  } catch (err) {
  console.error(err);
  res.status(500).send('Error uploading data');
}
});



app.listen(3003, function () {
  console.log("server connected");
})



