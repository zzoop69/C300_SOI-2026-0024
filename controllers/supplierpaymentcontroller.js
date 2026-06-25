const express = require("express");
const multer = require("multer");
const path = require("path");

const {
  approvePayment,
  createUploadRecord,
  deletePendingDocumentFromSet,
  deletePendingDocumentSet,
  getAllDiscrepancies,
  getDiscrepanciesByInvoiceId,
  getMatchingDetailsByInvoiceId,
  getMatchingSummary,
  getPendingDocumentSet,
  getPendingDocumentSets,
  getPaymentList,
  getRecord,
  getRecords,
  getReportRows,
  getStats,
  markPaymentPaid,
  paymentStatuses,
  rejectPayment,
  saveCorrectedData,
  savePendingDocumentUpload,
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

async function savePartialDocumentFiles(files, formData) {
  const uploadMap = [
    ["poFile", "purchaseOrder"],
    ["doGrnFile", "deliveryOrder"],
    ["invoiceFile", "invoice"],
  ];
  let result = null;
  let pendingSetId = formData.pendingSetId || "";

  for (const [fileField, documentType] of uploadMap) {
    if (!files[fileField]) {
      continue;
    }

    result = await savePendingDocumentUpload(
      { documentFile: files[fileField] },
      {
        ...formData,
        pendingSetId,
        documentType,
      }
    );
    pendingSetId = result.pendingSet?.id || pendingSetId;
  }

  return result;
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
    records,
  });
}));

router.get("/upload", asyncRoute(async (req, res) => {
  const [pendingDocumentSets, selectedPendingSet] = await Promise.all([
    getPendingDocumentSets(),
    req.query.set ? getPendingDocumentSet(req.query.set) : Promise.resolve(null),
  ]);

  res.render("upload", {
    pageTitle: "Upload Documents",
    activePage: "upload",
    successMessage: req.query.uploaded === "true"
      ? "Upload processed successfully."
      : req.query.deleted === "true"
        ? "Pending document set deleted."
        : "",
    holdMessage: req.query.held === "true" ? `Document saved on hold. Missing: ${req.query.missing || "required documents"}.` : "",
    completedMessage: req.query.completed === "true" ? "Document set is complete and extraction has started." : "",
    errorMessage: "",
    pendingDocumentSets,
    selectedPendingSet,
  });
}));

router.post(
  "/upload",
  upload.fields([
    { name: "documentFile", maxCount: 1 },
    { name: "poFile", maxCount: 1 },
    { name: "doGrnFile", maxCount: 1 },
    { name: "invoiceFile", maxCount: 1 },
  ]),
  asyncRoute(async (req, res) => {
    try {
      const legacyFiles = req.files || {};
      const hasLegacyUpload = legacyFiles.poFile || legacyFiles.doGrnFile || legacyFiles.invoiceFile;
      const hasCompleteLegacyUpload = legacyFiles.poFile && legacyFiles.doGrnFile && legacyFiles.invoiceFile;

      if (hasCompleteLegacyUpload && !req.body.pendingSetId) {
        await createUploadRecord(req.files || {});
        res.redirect(303, "/upload");
        return;
      }

      if (hasLegacyUpload) {
        if (req.body.holdIncomplete !== "true" && !req.body.pendingSetId) {
          throw new Error("Please upload all three files, or choose Put On Hold for an incomplete set.");
        }

        const result = await savePartialDocumentFiles(legacyFiles, req.body);

        if (!result) {
          throw new Error("Please select at least one Excel file.");
        }

        if (result.extractedRecordId) {
          res.redirect(303, "/upload");
          return;
        }

        res.redirect(303, "/upload");
        return;
      }

      const result = await savePendingDocumentUpload(req.files || {}, req.body);

      if (result.extractedRecordId) {
        res.redirect(303, "/upload");
        return;
      }

      res.redirect(303, "/upload");
    } catch (error) {
      const [pendingDocumentSets, selectedPendingSet] = await Promise.all([
        getPendingDocumentSets(),
        req.body?.pendingSetId ? getPendingDocumentSet(req.body.pendingSetId) : Promise.resolve(null),
      ]);

      res.status(400).render("upload", {
        pageTitle: "Upload Documents",
        activePage: "upload",
        successMessage: "",
        holdMessage: "",
        completedMessage: "",
        errorMessage: error.message || "Upload could not be processed. Please check the Excel files and try again.",
        pendingDocumentSets,
        selectedPendingSet,
      });
    }
  })
);

router.post("/pending-document-sets/:id/delete", asyncRoute(async (req, res) => {
  await deletePendingDocumentSet(req.params.id);
  res.redirect("/upload?deleted=true");
}));

router.post("/pending-document-sets/:id/documents/:documentType/delete", asyncRoute(async (req, res) => {
  await deletePendingDocumentFromSet(req.params.id, req.params.documentType);
  res.redirect(`/upload?set=${encodeURIComponent(req.params.id)}`);
}));

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
  res.redirect(`/matching-results/${encodeURIComponent(result.recordId)}?updated=true`);
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
    summaries: await getMatchingSummary(),
  });
}));

router.get("/matching-results/:invoiceId", asyncRoute(async (req, res) => {
  const [details, discrepancies] = await Promise.all([
    getMatchingDetailsByInvoiceId(req.params.invoiceId),
    getDiscrepanciesByInvoiceId(req.params.invoiceId),
  ]);

  if (!details) {
    res.redirect("/matching-results");
    return;
  }

  res.render("matching-details", {
    pageTitle: "Matching Details",
    activePage: "matching",
    details,
    discrepancies,
    updated: req.query.updated === "true",
  });
}));

router.get("/payment-approval", asyncRoute(async (req, res) => {
  const records = await getRecords();
  const statusGroups = {
    pending: records.filter((record) =>
      [paymentStatuses.pending, paymentStatuses.held].includes(record.paymentStatus)
    ),
    approved: records.filter((record) => record.paymentStatus === paymentStatuses.approved),
    processing: records.filter((record) => record.paymentStatus === paymentStatuses.processing),
    rejected: records.filter((record) => record.paymentStatus === paymentStatuses.rejected),
    paid: records.filter((record) => record.paymentStatus === paymentStatuses.paid),
    actionNeeded: records.filter((record) =>
      [paymentStatuses.pending, paymentStatuses.held].includes(record.paymentStatus)
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

  res.redirect("/payment-approval?updated=true&filter=approved");
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

  await setPaymentStatus(req.params.id, paymentStatuses.processing);

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