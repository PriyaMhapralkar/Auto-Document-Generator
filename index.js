require('dotenv').config();

const express = require('express');
const passport = require('passport');
const session = require('express-session');
const GoogleStrategy = require('passport-google-oauth20').Strategy;
const path = require('path');
const ExcelJS = require('exceljs');
const fs = require('fs').promises;

const app = express();
const port = process.env.PORT || 3000;

const EXCEL_FILE_PATH = path.join(__dirname, 'saved_texts.xlsx');
const COMPANY_FILE_PATH = path.join(__dirname, 'defined.xlsx');

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

app.use(session({
  secret: process.env.SESSION_SECRET || 'your_secret_key',
  resave: false,
  saveUninitialized: true,
  cookie: { secure: false }
}));

app.use(passport.initialize());
app.use(passport.session());

passport.use(new GoogleStrategy({
  clientID: process.env.GOOGLE_CLIENT_ID,
  clientSecret: process.env.GOOGLE_CLIENT_SECRET,
  callbackURL: "http://localhost:3000/auth/google/callback"
},
function (accessToken, refreshToken, profile, done) {
  return done(null, profile);
}));

passport.serializeUser((user, done) => done(null, user));
passport.deserializeUser((obj, done) => done(null, obj));

app.use(express.static(path.join(__dirname, 'public')));

app.get('/', (req, res) => {
  res.redirect('/login.html');
});


app.get('/auth/google',
  passport.authenticate('google', { scope: ['profile', 'email'] })
);

app.get('/auth/google/callback',
  passport.authenticate('google', { failureRedirect: '/' }),
  (req, res) => {
    res.redirect('/profile.html');
  }
);

app.get('/profile-data', (req, res) => {
  if (!req.isAuthenticated()) {
    return res.status(401).json({ message: 'Unauthorized' });
  }

  res.json({
    user: {
      displayName: req.user.displayName,
      email: req.user.emails?.[0]?.value || "N/A"
    }
  });
});

app.get('/get-companies', async (req, res) => {
  const workbook = new ExcelJS.Workbook();
  let companies = [];

  try {
    await fs.access(COMPANY_FILE_PATH);
    await workbook.xlsx.readFile(COMPANY_FILE_PATH);

    let sheet = workbook.getWorksheet('Companies');

    if (!sheet) {
      return res.json([]);
    }

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) { // Skip header row
        companies.push({
          CompanyName: row.getCell(1).value,
          RegNo: row.getCell(2).value,
          Address: row.getCell(3).value,
          // ADD NEW CA FIELDS HERE
          CAFirmName: row.getCell(4).value,
          CAName: row.getCell(5).value,
          MemberRegNo: row.getCell(6).value,
          PartnerProprietor: row.getCell(7).value
        });
      }
    });

    res.json(companies);

  } catch {
    res.json([]);
  }
});

app.post('/add-company', async (req, res) => {
  // DESTRUCTURE NEW CA FIELDS
  const { CompanyName, RegNo, Address, CAFirmName, CAName, MemberRegNo, PartnerProprietor } = req.body;

  if (!CompanyName) {
    return res.status(400).json({ message: "Company name required" });
  }

  const workbook = new ExcelJS.Workbook();
  let sheet;

  try {
    await fs.access(COMPANY_FILE_PATH);
    await workbook.xlsx.readFile(COMPANY_FILE_PATH);
    sheet = workbook.getWorksheet('Companies');

    if (!sheet) {
      sheet = workbook.addWorksheet('Companies');
      // ADD NEW HEADERS
      sheet.addRow(['CompanyName', 'RegNo', 'Address', 'CAFirmName', 'CAName', 'MemberRegNo', 'PartnerProprietor']);
    }

  } catch {
    sheet = workbook.addWorksheet('Companies');
    // ADD NEW HEADERS
    sheet.addRow(['CompanyName', 'RegNo', 'Address', 'CAFirmName', 'CAName', 'MemberRegNo', 'PartnerProprietor']);
  }

  // ADD NEW VALUES
  sheet.addRow([CompanyName, RegNo, Address, CAFirmName, CAName, MemberRegNo, PartnerProprietor]);

  await workbook.xlsx.writeFile(COMPANY_FILE_PATH);

  res.json({ success: true, message: "Company Added" });
});

app.post('/api/save-text', async (req, res) => {
  if (!req.isAuthenticated()) {
    return res.status(401).json({ message: "Unauthorized" });
  }

  const {
    "User Email": userEmail,
    "Date": date,
    "Company Name A": companyA,
    "LLPIN": llpin,
    "Address A": addressA,
    "Company Name B": companyB,
    "CIN": cin,
    "Address B": addressB,
    // NEW CA FIELDS TO DESTRUCTURE FROM REQ.BODY
    "CA Firm Name": caFirmName,
    "CA Name": caName,
    "Member Reg No": memberRegNo,
    "Partner/Proprietor": partnerProprietor
  } = req.body;

  const workbook = new ExcelJS.Workbook();
  let sheet;

  try {
    await fs.access(EXCEL_FILE_PATH);
    await workbook.xlsx.readFile(EXCEL_FILE_PATH);
    sheet = workbook.getWorksheet('User Details');

    if (!sheet) {
      sheet = workbook.addWorksheet('User Details');
      sheet.addRow([
        "Timestamp",
        "User Email",
        "Date",
        "Company Name A",
        "LLPIN",
        "Address A",
        "Company Name B",
        "CIN",
        "Address B",
        // NEW HEADERS FOR SAVED TEXTS
        "CA Firm Name",
        "CA Name",
        "Member Reg No",
        "Partner/Proprietor"
      ]);
    }

  } catch {
    sheet = workbook.addWorksheet('User Details');
    sheet.addRow([
      "Timestamp",
      "User Email",
      "Date",
      "Company Name A",
      "LLPIN",
      "Address A",
      "Company Name B",
      "CIN",
      "Address B",
      // NEW HEADERS FOR SAVED TEXTS
      "CA Firm Name",
      "CA Name",
      "Member Reg No",
      "Partner/Proprietor"
    ]);
  }

  sheet.addRow([
    new Date().toISOString(),
    userEmail,
    date,
    companyA,
    llpin,
    addressA,
    companyB,
    cin,
    addressB,
    // NEW VALUES FOR SAVED TEXTS
    caFirmName,
    caName,
    memberRegNo,
    partnerProprietor
  ]);

  await workbook.xlsx.writeFile(EXCEL_FILE_PATH);

  res.json({ success: true, message: "Saved successfully" });
});

app.get('/api/get-last-excel-entry', async (req, res) => {
  const workbook = new ExcelJS.Workbook();

  try {
    await fs.access(EXCEL_FILE_PATH);
    await workbook.xlsx.readFile(EXCEL_FILE_PATH);

    const sheet = workbook.getWorksheet('User Details');

    if (!sheet || sheet.rowCount < 2) {
      return res.json({ success: false });
    }

    const headers = [];
    sheet.getRow(1).eachCell((cell) => headers.push(cell.value));

    const last = sheet.getRow(sheet.rowCount);

    const row = {};
    headers.forEach((h, i) => {
      row[h] = last.getCell(i + 1).value;
    });

    res.json({ success: true, row });

  } catch {
    res.json({ success: false });
  }
});

app.get('/logout', (req, res) => {
  req.logout(() => {
    res.redirect('/');
  });
});


app.listen(port, '0.0.0.0', () => {
    console.log(`Server running on port ${port}`);
});
