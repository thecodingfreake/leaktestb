const express = require('express');
const cors = require('cors');
const xlsx = require('xlsx');
  const multer = require('multer');
  const path = require('path');
  const fs = require('fs');

const app = express();
const port = 5000;

app.use(cors());

// Ensure the uploads folder exists
const uploadDirectory = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadDirectory)) {
  fs.mkdirSync(uploadDirectory);
}

// Configure multer storage options
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, uploadDirectory); // Save to 'uploads' folder
  },
  filename: (req, file, cb) => {
    const allowedFilenames = ['image1.jpg', 'image2.jpg'];
    const requestedFileName = `${file.fieldname}.jpg`;

    // Ensure the uploaded file is one of the two allowed filenames
    if (allowedFilenames.includes(requestedFileName)) {
      cb(null, requestedFileName); // Replace existing file
    } else {
      cb(new Error('Invalid file field name. Use "image1" or "image2".'));
    }
  }
});

const upload = multer({ storage: storage });

// Route to handle image uploads (accepts two files: image1 and image2)
app.post('/upload', upload.fields([{ name: 'image1' }, { name: 'image2' }]), (req, res) => {
  if (!req.files || (!req.files.image1 && !req.files.image2)) {
    return res.status(400).json({ error: 'No files uploaded or invalid file names' });
  }

  res.status(200).json({
    message: 'Files uploaded successfully',
    uploadedFiles: Object.keys(req.files).map((key) => `/uploads/${req.files[key][0].filename}`)
  });
});


app.get('/uploaded-images', (req, res) => {
  const image1Path = '/uploads/image1.jpg';
  const image2Path = '/uploads/image2.jpg';

  res.json({
      image1: fs.existsSync(path.join(__dirname, image1Path)) ? image1Path : null,
      image2: fs.existsSync(path.join(__dirname, image2Path)) ? image2Path : null,
  });
});


app.get('/data', (req, res) => {
  const filePath = path.join(__dirname, 'Site Refrigerant Inventory - R2.xlsx');
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  // Get the range of the sheet
  const range = xlsx.utils.decode_range(sheet['!ref']);

  // Get the headers from the third row
  const headers = [];
  for (let C = range.s.c; C <= range.e.c; ++C) {
    const cell = sheet[xlsx.utils.encode_cell({ r: 2, c: C })]; // Third row (index 2)
    headers.push(cell ? xlsx.utils.format_cell(cell) : `Column${C}`);
  }

  // Get the data starting from the fourth row
  const data = [];
  for (let R = 3; R <= range.e.r; ++R) {
    const row = {};
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cell = sheet[xlsx.utils.encode_cell({ r: R, c: C })];
      const header = headers[C];
      if (cell && header) {
        row[header] = xlsx.utils.format_cell(cell);
      }
    }
    data.push(row);
  }

  res.json(data);
});

app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
});
