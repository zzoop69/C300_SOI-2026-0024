const express = require("express");
const multer = require("multer");
const path = require("path");

const {
  addDocumentSet,
  getDashboardStats,
  getDocumentSets,
  getReportRows,
} = require("../models/supplierpayment");

const router = express.Router();

// Uploaded files are kept locally so the project is easy to demo in class.
const upload = multer({
  dest: path.join(__dirname, "..", "uploads"),
});

router.get("/", (req, res) => {
  res.redirect("/dashboard");
});

router.get("/dashboard", (req, res) => {
  res.render("dashboard", {
    pageTitle: "Dashboard",
    activePage: "dashboard",
    stats: getDashboardStats(),
    recentSets: getDocumentSets().slice(0, 5),
  });
});

router.get("/upload", (req, res) => {
  res.render("upload", {
    pageTitle: "Upload Documents",
    activePage: "upload",
    uploaded: req.query.uploaded === "true",
  });
});

router.post(
  "/upload",
  upload.fields([
    { name: "poFile", maxCount: 1 },
    { name: "soFile", maxCount: 1 },
    { name: "invoiceFile", maxCount: 1 },
  ]),
  (req, res) => {
    // This simulates document extraction until OCR/PDF reading is added later.
    addDocumentSet({
      supplierName: req.body.supplierName,
      poNumber: req.body.poNumber,
      soNumber: req.body.soNumber,
      invoiceNumber: req.body.invoiceNumber,
      itemName: req.body.itemName,
      quantity: req.body.quantity,
      unitPrice: req.body.unitPrice,
      totalAmount: req.body.totalAmount,
      date: req.body.date,
      uploadedFiles: {
        po: req.files.poFile?.[0]?.originalname || "No PO file uploaded",
        so: req.files.soFile?.[0]?.originalname || "No SO file uploaded",
        invoice: req.files.invoiceFile?.[0]?.originalname || "No invoice file uploaded",
      },
    });

    res.redirect("/upload?uploaded=true");
  }
);

router.get("/records", (req, res) => {
  res.render("records", {
    pageTitle: "Document Records",
    activePage: "records",
    documentSets: getDocumentSets(),
  });
});

router.get("/matching-results", (req, res) => {
  res.render("matching-results", {
    pageTitle: "Matching Results",
    activePage: "matching",
    documentSets: getDocumentSets(),
  });
});

router.get("/reports", (req, res) => {
  res.render("reports", {
    pageTitle: "Reports",
    activePage: "reports",
    stats: getDashboardStats(),
    rows: getReportRows(),
  });
});

module.exports = router;
