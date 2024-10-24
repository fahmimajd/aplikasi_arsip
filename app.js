const express = require('express');
const multer = require('multer');
const path = require('path');
const ExcelJS = require('exceljs');
const xlsx = require('xlsx');
const helmet = require('helmet');  // Tambahkan Helmet
const csrf = require('csurf');     // Tambahkan CSRF Protection
const cookieParser = require('cookie-parser'); // Dibutuhkan untuk CSRF
const fs = require('fs');  // Tambahkan modul fs (File System)

const app = express();

// Aktifkan helmet untuk keamanan header HTTP
app.use(helmet());

// Middleware untuk parsing cookie (diperlukan untuk CSRF protection)
app.use(cookieParser());

// Konfigurasi CSRF protection
const csrfProtection = csrf({ cookie: true });

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
app.get('/', csrfProtection, (req, res) => {
  res.render('index', { results: null, csrfToken: req.csrfToken() });
});

// Route untuk halaman upload
app.get('/upload', csrfProtection, (req, res) => {
  res.render('upload', { csrfToken: req.csrfToken() });
});

// Route untuk menangani upload
app.post('/upload', upload.single('excelFile'), csrfProtection, (req, res) => {
  // File diupload dan disimpan di folder 'uploads' dengan nama 'db.xlsx'
  res.redirect('/');  // Redirect ke halaman utama setelah upload
});

// Route untuk menangani pencarian
app.post('/search', csrfProtection, (req, res) => {
  const { creator, year, title } = req.body;
  const filePath = path.join(__dirname, 'uploads', 'db.xlsx'); // Path ke file db.xlsx

  if (!fs.existsSync(filePath)) {
    return res.render('index', { results: null, error: 'File Excel belum diupload.', csrfToken: req.csrfToken() });
  }

  const data = readExcelData(filePath);
  // Log data yang dibaca dari Excel
  console.log('Data dari Excel:', data);

  // Filter data berdasarkan input pencarian
  const results = data.filter((row) => {
    const matchCreator = creator ? row[8] && row[8].toLowerCase().includes(creator.toLowerCase()) : true; // Kolom I
    const matchTitle = title ? row[2] && row[2].toLowerCase().includes(title.toLowerCase()) : true; // Kolom C
    const matchYear = year ? row[3] && row[3].toString().includes(year.toString()) : true; // Kolom D
    return matchCreator && matchTitle && matchYear;
  });

  // Log hasil pencarian
  console.log('Hasil pencarian:', results);
  
  res.render('index', { results, csrfToken: req.csrfToken() });
});

const PORT = process.env.PORT || 3005;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
