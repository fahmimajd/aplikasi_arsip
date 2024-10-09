const express = require('express');
const multer = require('multer');
const path = require('path');
const ExcelJS = require('exceljs');
const xlsx = require('xlsx');


const app = express();
// Set storage untuk multer
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, './uploads');
  },
  filename: (req, file, cb) => {
    cb(null, 'db.xlsx');  // File diupload dengan nama tetap
  }
});

const upload = multer({ storage: storage });

// Set view engine
app.set('view engine', 'ejs');
app.use(express.urlencoded({ extended: true }));
//app.use(express.static('public'));
app.use(express.static(path.join(__dirname, 'public')));

// Tambahkan middleware untuk mengurai data form
app.use(express.urlencoded({ extended: true }));

// Fungsi untuk membaca data dari Excel
const readExcelData = (filePath) => {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const jsonData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
  return jsonData;
};

// Route utama untuk menampilkan form pencarian
app.get('/', (req, res) => {
  res.render('index', { results: null });
});

// Route untuk halaman upload
app.get('/upload', (req, res) => {
    res.render('upload');
  });

  // Route untuk menangani upload
app.post('/upload', upload.single('excelFile'), (req, res) => {
    // File diupload dan disimpan di folder 'uploads' dengan nama 'db.xlsx'
    res.redirect('/');  // Redirect ke halaman utama setelah upload
  });

// Route untuk menangani pencarian
app.post('/search', (req, res) => {
  const { creator, year, title } = req.body;
  const filePath = path.join(__dirname, 'uploads', 'db.xlsx'); // Path ke file db.xlsx

  const data = readExcelData(filePath);

  // Filter data berdasarkan input pencarian
  const results = data.filter((row) => {
    const matchCreator = creator ? row[10] && row[10].toLowerCase().includes(creator.toLowerCase()) : true;
    const matchTitle = title ? row[3] && row[3].toLowerCase().includes(title.toLowerCase()) : true;
    const matchYear = year ? row[6] && row[6].toString() === year : true;
    return matchCreator && matchTitle && matchYear;
  });

  res.render('index', { results });
});

const PORT = process.env.PORT || 3005;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
