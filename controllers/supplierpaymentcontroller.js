const express = require("express");
const multer = require("multer");
const path = require("path");

const {
  approvePayment,
  createUploadRecord,
  fieldsToCompare,
  getAllDiscrepancies,
  getPaymentList,
  getRecord,
  getRecords,
  getReportRows,
  getStats,
  markPaymentPaid,
  paymentStatuses,
  rejectPayment,
  saveCorrectedData,
  setPaymentStatus,
  simulatePayment,
} = require("../models/supplierpayment");

const router = express.Router();

// multer saves uploaded Excel files inside the uploads folder.
const upload = multer({
  dest: path.join(__dirname, "..", "uploads"),
});

function findRecordOrRedirect(req, res) {
  const record = getRecord(req.params.id);

  if (!record) {
    res.redirect("/records");
    return null;
  }

  return record;
}

router.get("/", (req, res) => {
  res.redirect("/dashboard");
});

router.get("/dashboard", (req, res) => {
  res.render("dashboard", {
    pageTitle: "Dashboard",
    activePage: "dashboard",
    stats: getStats(),
    records: getRecords().slice(0, 6),
  });
});

router.get("/upload", (req, res) => {
  res.render("upload", {
    pageTitle: "Upload Documents",
    activePage: "upload",
  });
});

router.post(
  "/upload",
  upload.fields([
    { name: "poFile", maxCount: 1 },
    { name: "doGrnFile", maxCount: 1 },
    { name: "invoiceFile", maxCount: 1 },
  ]),
  (req, res) => {
    const recordId = createUploadRecord(req.files || {});
    res.redirect(`/extract/${recordId}`);
  }
);

router.get("/extracted-data", (req, res) => {
  res.render("records", {
    pageTitle: "Extracted Data Review",
    activePage: "review",
    records: getRecords(),
  });
});

router.get("/extract/:id", (req, res) => {
  const record = findRecordOrRedirect(req, res);

  if (!record) {
    return;
  }

  res.render("extract-review", {
    pageTitle: "Extracted Data Review",
    activePage: "review",
    record,
    saved: req.query.saved === "true",
  });
});

router.post("/extract/:id/save", (req, res) => {
  saveCorrectedData(req.params.id, req.body);
  res.redirect(`/validate-data?record=${req.params.id}`);
});

router.get("/records", (req, res) => {
  res.redirect("/extracted-data");
});

router.get("/validate-data", (req, res) => {
  res.render("validate-data", {
    pageTitle: "Validate Data",
    activePage: "validate",
    records: getRecords(),
    selectedRecordId: req.query.record,
  });
});

router.get("/matching-results", (req, res) => {
  res.render("matching-results", {
    pageTitle: "Matching Results",
    activePage: "matching",
    records: getRecords(),
    fieldsToCompare,
  });
});

router.get("/payment-approval", (req, res) => {
  const records = getRecords();
  const statusGroups = {
    pending: records.filter((record) => record.paymentStatus === paymentStatuses.pending),
    approved: records.filter((record) => record.paymentStatus === paymentStatuses.approved),
    processing: records.filter((record) => record.paymentStatus === paymentStatuses.processing),
    held: records.filter((record) => record.paymentStatus === paymentStatuses.held),
    rejected: records.filter((record) => record.paymentStatus === paymentStatuses.rejected),
    paid: records.filter((record) => record.paymentStatus === paymentStatuses.paid),
    actionNeeded: records.filter((record) =>
      [paymentStatuses.pending, paymentStatuses.held].includes(record.paymentStatus)
    ),
    approvedProcessing: records.filter((record) =>
      [paymentStatuses.approved, paymentStatuses.processing].includes(record.paymentStatus)
    ),
  };

  res.render("payment-approval", {
    pageTitle: "Payment Approval",
    activePage: "approval",
    records,
    paymentList: getPaymentList(),
    statusGroups,
    currentFilter: req.query.filter || "all",
    success: req.query.updated === "true",
  });
});

router.post("/payment-approval/:id/approve", (req, res) => {
  const record = findRecordOrRedirect(req, res);

  if (!record) {
    return;
  }

  approvePayment(req.params.id);

  res.redirect("/payment-approval?updated=true");
});

router.post("/payment-approval/:id/reject", (req, res) => {
  rejectPayment(req.params.id);
  res.redirect("/payment-approval?updated=true&filter=rejected");
});

router.post("/payment-approval/:id/process", (req, res) => {
  const record = findRecordOrRedirect(req, res);

  if (!record) {
    return;
  }

  if (record.paymentStatus === paymentStatuses.approved) {
    setPaymentStatus(req.params.id, paymentStatuses.processing);
  }

  res.redirect("/payment-approval?updated=true&filter=processing");
});

router.post("/payment-approval/:id/paid", (req, res) => {
  markPaymentPaid(req.params.id, req.body.paymentMethod);
  res.redirect("/payment-approval?updated=true&filter=paid");
});

router.get("/payment-simulation", (req, res) => {
  res.render("payment-simulation", {
    pageTitle: "Payment Simulation",
    activePage: "simulation",
    records: getRecords(),
  });
});

router.post("/payment-simulation/:id/pay", (req, res) => {
  const record = simulatePayment(req.params.id, req.body.paymentMethod);

  if (!record || !record.payment) {
    res.redirect("/payment-simulation");
    return;
  }

  res.redirect(`/receipt/${record.id}`);
});

router.get("/receipt/:id", (req, res) => {
  const record = findRecordOrRedirect(req, res);

  if (!record) {
    return;
  }

  res.render("receipt", {
    pageTitle: "Receipt",
    activePage: "simulation",
    record,
  });
});

router.get("/reports", (req, res) => {
  res.render("reports", {
    pageTitle: "Reports",
    activePage: "reports",
    stats: getStats(),
    rows: getReportRows(),
    discrepancies: getAllDiscrepancies(),
  });
});

module.exports = router;
