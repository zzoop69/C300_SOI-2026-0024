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

function asyncRoute(handler) {
  return (req, res, next) => {
    Promise.resolve(handler(req, res, next)).catch(next);
  };
}

async function findRecordOrRedirect(req, res) {
  const record = await getRecord(req.params.id);

  if (!record) {
    res.redirect("/records");
    return null;
  }

  return record;
}

router.get("/", (req, res) => {
  res.redirect("/dashboard");
});

router.get("/dashboard", asyncRoute(async (req, res) => {
  const [stats, records] = await Promise.all([getStats(), getRecords()]);

  res.render("dashboard", {
    pageTitle: "Dashboard",
    activePage: "dashboard",
    stats,
    records: records.slice(0, 6),
  });
}));

router.get("/upload", (req, res) => {
  res.render("upload", {
    pageTitle: "Upload Documents",
    activePage: "upload",
    successMessage: req.query.uploaded === "true" ? "Upload processed successfully." : "",
    errorMessage: "",
  });
});

router.post(
  "/upload",
  upload.fields([
    { name: "poFile", maxCount: 1 },
    { name: "doGrnFile", maxCount: 1 },
    { name: "invoiceFile", maxCount: 1 },
  ]),
  asyncRoute(async (req, res) => {
    try {
      const result = await createUploadRecord(req.files || {});
      const supplierCreatedQuery = result.supplierAutoCreated ? "?supplierCreated=true" : "";
      res.redirect(`/extract/${result.recordId}${supplierCreatedQuery}`);
    } catch (error) {
      res.status(400).render("upload", {
        pageTitle: "Upload Documents",
        activePage: "upload",
        successMessage: "",
        errorMessage: error.message || "Upload could not be processed. Please check the Excel files and try again.",
      });
    }
  })
);

router.get("/extracted-data", asyncRoute(async (req, res) => {
  res.render("records", {
    pageTitle: "Extracted Data Review",
    activePage: "review",
    records: await getRecords(),
  });
}));

router.get("/extract/:id", asyncRoute(async (req, res) => {
  const record = await findRecordOrRedirect(req, res);

  if (!record) {
    return;
  }

  res.render("extract-review", {
    pageTitle: "Extracted Data Review",
    activePage: "review",
    record,
    saved: req.query.saved === "true",
    supplierCreated: req.query.supplierCreated === "true",
  });
}));

router.post("/extract/:id/save", asyncRoute(async (req, res) => {
  const result = await saveCorrectedData(req.params.id, req.body);
  const supplierCreatedQuery = result.supplierAutoCreated ? "&supplierCreated=true" : "";
  res.redirect(`/validate-data?record=${result.recordId}${supplierCreatedQuery}`);
}));

router.get("/records", (req, res) => {
  res.redirect("/extracted-data");
});

router.get("/validate-data", asyncRoute(async (req, res) => {
  res.render("validate-data", {
    pageTitle: "Validate Data",
    activePage: "validate",
    records: await getRecords(),
    selectedRecordId: req.query.record,
    supplierCreated: req.query.supplierCreated === "true",
  });
}));

router.get("/matching-results", asyncRoute(async (req, res) => {
  res.render("matching-results", {
    pageTitle: "Matching Results",
    activePage: "matching",
    records: await getRecords(),
    fieldsToCompare,
  });
}));

router.get("/payment-approval", asyncRoute(async (req, res) => {
  const records = await getRecords();
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
    paymentList: await getPaymentList(),
    statusGroups,
    currentFilter: req.query.filter || "all",
    success: req.query.updated === "true",
  });
}));

router.post("/payment-approval/:id/approve", asyncRoute(async (req, res) => {
  const record = await findRecordOrRedirect(req, res);

  if (!record) {
    return;
  }

  await approvePayment(req.params.id);

  res.redirect("/payment-approval?updated=true");
}));

router.post("/payment-approval/:id/reject", asyncRoute(async (req, res) => {
  await rejectPayment(req.params.id);
  res.redirect("/payment-approval?updated=true&filter=rejected");
}));

router.post("/payment-approval/:id/process", asyncRoute(async (req, res) => {
  const record = await findRecordOrRedirect(req, res);

  if (!record) {
    return;
  }

  if (record.paymentStatus === paymentStatuses.approved) {
    await setPaymentStatus(req.params.id, paymentStatuses.processing);
  }

  res.redirect("/payment-approval?updated=true&filter=processing");
}));

router.post("/payment-approval/:id/paid", asyncRoute(async (req, res) => {
  await markPaymentPaid(req.params.id, req.body.paymentMethod);
  res.redirect("/payment-approval?updated=true&filter=paid");
}));

router.get("/payment-simulation", asyncRoute(async (req, res) => {
  res.render("payment-simulation", {
    pageTitle: "Payment Simulation",
    activePage: "simulation",
    records: await getRecords(),
  });
}));

router.post("/payment-simulation/:id/pay", asyncRoute(async (req, res) => {
  const record = await simulatePayment(req.params.id, req.body.paymentMethod);

  if (!record || !record.payment) {
    res.redirect("/payment-simulation");
    return;
  }

  res.redirect(`/receipt/${record.id}`);
}));

router.get("/receipt/:id", asyncRoute(async (req, res) => {
  const record = await findRecordOrRedirect(req, res);

  if (!record) {
    return;
  }

  res.render("receipt", {
    pageTitle: "Receipt",
    activePage: "simulation",
    record,
  });
}));

router.get("/reports", asyncRoute(async (req, res) => {
  const [stats, rows, discrepancies] = await Promise.all([getStats(), getReportRows(), getAllDiscrepancies()]);

  res.render("reports", {
    pageTitle: "Reports",
    activePage: "reports",
    stats,
    rows,
    discrepancies,
  });
}));

module.exports = router;
