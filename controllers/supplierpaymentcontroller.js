const express = require("express");
const multer = require("multer");
const path = require("path");
const XLSX = require("xlsx-js-style");
const paypal = require("../services/paypal");
const {
  applyPayPalStatusByBatch,
  getPayoutByInvoice,
  itemStatusFromBatch,
  preparePayout,
  recordFailure,
  releaseWebhookEvent,
  recordSubmission,
  recordWebhookOnce,
} = require("../models/paypalpayment");

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
  paymentStatuses,
  rejectPayment,
  saveCorrectedData,
  savePendingDocumentUpload,
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

const PAYPAL_SUCCESS_STATUSES = new Set(["SUCCESS", "SUCCEEDED"]);
const PAYPAL_TERMINAL_STATUSES = new Set([
  ...PAYPAL_SUCCESS_STATUSES,
  "FAILED",
  "DENIED",
  "BLOCKED",
  "RETURNED",
  "REFUNDED",
  "CANCELED",
  "UNCLAIMED",
]);

function wait(milliseconds) {
  return new Promise((resolve) => setTimeout(resolve, milliseconds));
}

async function waitForPayPalPayoutResult(batchId, initialResponse) {
  let response = initialResponse;
  let status = String(itemStatusFromBatch(response)).toUpperCase();

  for (let attempt = 0; attempt < 8 && !PAYPAL_TERMINAL_STATUSES.has(status); attempt += 1) {
    if (attempt > 0) await wait(750);
    response = await paypal.getPayoutBatch(batchId);
    status = String(itemStatusFromBatch(response)).toUpperCase();
  }

  return { response, status };
}

async function submitPayPalPayout(invoiceId) {
  let prepared;
  try {
    const record = await getRecord(invoiceId);
    if (!record) throw new Error("Payment record was not found.");
    if (
      !record.canApprovePayment ||
      record.matching?.status !== "Matched" ||
      record.validation?.status !== "Valid" ||
      record.unresolvedDiscrepancies?.length
    ) {
      throw new Error(
        "Accept is disabled until matching is Matched, validation is Valid, and all discrepancies are resolved."
      );
    }
    prepared = await preparePayout(invoiceId);
    const response = await paypal.createPayout({
      senderBatchId: prepared.senderBatchId,
      senderItemId: prepared.senderItemId,
      recipientEmail: prepared.recipientEmail,
      amount: prepared.amount_due,
      currency: prepared.currency,
      invoiceId: prepared.invoiceId,
    });
    await recordSubmission(prepared.payoutId, response);
    const batchId = response.batch_header?.payout_batch_id;
    if (!batchId) {
      throw new Error("PayPal accepted the request without returning a payout batch ID.");
    }
    return {
      ...prepared,
      payoutStatus: String(response.batch_header?.batch_status || "PENDING").toUpperCase(),
      completed: false,
    };
  } catch (error) {
    if (prepared?.payoutId) await recordFailure(prepared.payoutId, error);
    throw error;
  }
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

function matchingHistoryExportRow(row) {
  return {
    Record: `#${row.id}`,
    Supplier: row.supplierName,
    PO: row.poNumber,
    "DO/GRN": row.doGrnNumber,
    Invoice: row.invoiceNumber,
    Amount: Number(row.amount || 0),
    "Match Status": row.matchStatus,
    "Validation Status": row.validationStatus,
    "Approval Status": row.approvalStatus,
    "Mismatch Reasons": row.mismatchReasons || "",
  };
}

function appendMatchingHistorySheet(workbook, sheetName, rows) {
  const exportRows = rows.map(matchingHistoryExportRow);
  const worksheet = XLSX.utils.json_to_sheet(exportRows, {
    header: [
      "Record",
      "Supplier",
      "PO",
      "DO/GRN",
      "Invoice",
      "Amount",
      "Match Status",
      "Validation Status",
      "Approval Status",
      "Mismatch Reasons",
    ],
  });

  worksheet["!cols"] = [
    { wch: 12 },
    { wch: 28 },
    { wch: 18 },
    { wch: 18 },
    { wch: 18 },
    { wch: 14 },
    { wch: 18 },
    { wch: 20 },
    { wch: 20 },
    { wch: 55 },
  ];

  worksheet["!rows"] = [
    { hpt: 24 },
    ...exportRows.map((row) => {
      const reasonLines = String(row["Mismatch Reasons"] || "")
        .split("\n")
        .reduce((lineCount, reason) => lineCount + Math.max(1, Math.ceil(reason.length / 55)), 0);
      return { hpt: Math.max(22, Math.min(reasonLines * 18, 240)) };
    }),
  ];

  const range = XLSX.utils.decode_range(worksheet["!ref"]);
  for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex += 1) {
    for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex += 1) {
      const cell = worksheet[XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex })];
      if (!cell) continue;
      cell.s = {
        ...(cell.s || {}),
        alignment: {
          vertical: "top",
          wrapText: columnIndex === range.e.c,
        },
      };
    }
  }

  XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
}

function createMatchingHistoryWorkbook(rows) {
  const workbook = XLSX.utils.book_new();
  appendMatchingHistorySheet(workbook, "All Matching History", rows);
  appendMatchingHistorySheet(
    workbook,
    "Matched Records",
    rows.filter((row) => row.matchStatus === "Matched")
  );
  appendMatchingHistorySheet(
    workbook,
    "Mismatched Records",
    rows.filter((row) => row.matchStatus === "Mismatch")
  );

  return workbook;
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
      ? "Documents uploaded successfully!"
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
        res.redirect(303, "/upload?uploaded=true");
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
          res.redirect(303, "/upload?uploaded=true");
          return;
        }

        res.redirect(303, "/upload");
        return;
      }

      const result = await savePendingDocumentUpload(req.files || {}, req.body);

      if (result.extractedRecordId) {
        res.redirect(303, "/upload?uploaded=true");
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
    viewOnly: req.query.mode === "view",
    fromValidation: req.query.from === "validate-data",
    backUrl: req.query.from === "extracted-data" ? "/extracted-data" : "/validate-data",
    saved: req.query.saved === "true",
    revalidated: req.query.revalidated === "true",
    supplierCreated: req.query.supplierCreated === "true",
  });
}));

router.post("/extract/:id/save", asyncRoute(async (req, res) => {
  const result = await saveCorrectedData(req.params.id, req.body);
  if (!result.saved) {
    const fromExtractedData = req.body.returnContext === "extracted-data";
    res.status(422).render("extract-review", {
      pageTitle: "Extracted Data Review",
      activePage: "review",
      record: result.record,
      viewOnly: false,
      fromValidation: !fromExtractedData,
      backUrl: fromExtractedData ? "/extracted-data" : "/validate-data",
      saved: false,
      revalidated: true,
      supplierCreated: false,
    });
    return;
  }
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
    searchQuery: req.query.q || "",
    success: req.query.updated === "true",
    workflowError: req.query.error || "",
  });
}));

router.post("/payment-approval/:id/approve", asyncRoute(async (req, res) => {
  const record = await findRecordOrRedirect(req, res);
  if (!record) return;

  if (
    !record.canApprovePayment ||
    record.matching?.status !== "Matched" ||
    record.validation?.status !== "Valid" ||
    record.unresolvedDiscrepancies?.length
  ) {
    res.redirect(
      `/payment-approval?filter=pending&error=${encodeURIComponent(
        "Accept is disabled until matching is Matched, validation is Valid, and all discrepancies are resolved."
      )}`
    );
    return;
  }

  await approvePayment(req.params.id);
  res.redirect("/payment-approval?updated=true&filter=approved");
}));

router.post("/payment-approval/:id/paypal", asyncRoute(async (req, res) => {
  try {
    const result = await submitPayPalPayout(req.params.id);
    res.redirect(
      `/receipt/${encodeURIComponent(req.params.id)}?submitted=${result.completed ? "true" : "false"}`
    );
  } catch (error) {
    res.redirect(
      `/payment-approval?filter=all&error=${encodeURIComponent(error.message || "PayPal payout could not be submitted.")}`
    );
  }
}));

router.post("/payment-approval/:id/reject", asyncRoute(async (req, res) => {
  await rejectPayment(req.params.id);
  res.redirect("/payment-approval?updated=true&filter=rejected");
}));

router.get("/payment-simulation", (req, res) => {
  res.redirect(301, "/payment-approval");
});

// Compatibility endpoint for old bookmarks/forms. It uses the same duplicate-safe
// server-side submission path and is intentionally not linked from the UI.
router.post("/payment-simulation/:id/pay", asyncRoute(async (req, res) => {
  try {
    const result = await submitPayPalPayout(req.params.id);
    res.redirect(
      `/receipt/${encodeURIComponent(req.params.id)}?submitted=${result.completed ? "true" : "false"}`
    );
  } catch (error) {
    res.redirect(`/payment-approval?error=${encodeURIComponent(error.message || "PayPal payout could not be submitted.")}`);
  }
}));

router.get("/receipt/:id/status", asyncRoute(async (req, res) => {
  const invoiceId = String(req.params.id || "").trim();
  if (!/^[A-Za-z0-9_-]{1,20}$/.test(invoiceId)) {
    res.status(400).json({
      ok: false,
      retryable: false,
      error: "Invalid invoice ID.",
    });
    return;
  }

  let [record, payout] = await Promise.all([
    getRecord(invoiceId),
    getPayoutByInvoice(invoiceId),
  ]);

  if (!record || !payout) {
    res.status(404).json({
      ok: false,
      retryable: false,
      error: "Payment status was not found.",
    });
    return;
  }

  const terminalStatuses = PAYPAL_TERMINAL_STATUSES;
  let gatewayStatus = String(payout.status || "PENDING").toUpperCase();
  const paymentAlreadyPaid = record.dbPaymentStatus === "PAID";
  const alreadyFinal = paymentAlreadyPaid || terminalStatuses.has(gatewayStatus);

  if (!alreadyFinal && record.paymentStatus === paymentStatuses.processing) {
    if (!payout.paypal_batch_id) {
      res.status(409).json({
        ok: false,
        retryable: true,
        invoiceId,
        paymentStatus: record.dbPaymentStatus,
        gatewayStatus,
        final: false,
        successful: false,
      });
      return;
    }

    let response;
    try {
      response = await paypal.getPayoutBatch(payout.paypal_batch_id);
    } catch {
      res.status(503).json({
        ok: false,
        retryable: true,
        invoiceId,
        paymentStatus: record.dbPaymentStatus,
        gatewayStatus,
        final: false,
        successful: false,
      });
      return;
    }

    const item = response.items?.[0];
    const itemStatus = item?.transaction_status
      ? String(item.transaction_status).toUpperCase()
      : "";

    // Only an individual payout-item status is allowed to drive finalisation.
    // A batch-level SUCCESS without an item result remains non-final.
    if (itemStatus) {
      await applyPayPalStatusByBatch(
        payout.paypal_batch_id,
        itemStatus,
        response,
        item.payout_item_id || null
      );
      [record, payout] = await Promise.all([
        getRecord(invoiceId),
        getPayoutByInvoice(invoiceId),
      ]);
      gatewayStatus = String(payout?.status || itemStatus).toUpperCase();
    }
  }

  const successful =
    record.dbPaymentStatus === "PAID" &&
    PAYPAL_SUCCESS_STATUSES.has(gatewayStatus);
  const final =
    successful ||
    (PAYPAL_TERMINAL_STATUSES.has(gatewayStatus) &&
      !PAYPAL_SUCCESS_STATUSES.has(gatewayStatus));

  res.json({
    ok: true,
    invoiceId,
    paymentStatus: record.dbPaymentStatus,
    gatewayStatus,
    final,
    successful,
  });
}));

router.post("/receipt/:id/refresh", asyncRoute(async (req, res) => {
  const payout = await getPayoutByInvoice(req.params.id);
  if (!payout?.paypal_batch_id) throw new Error("No submitted PayPal payout was found for this invoice.");
  const response = await paypal.getPayoutBatch(payout.paypal_batch_id);
  const item = response.items?.[0];
  await applyPayPalStatusByBatch(payout.paypal_batch_id, itemStatusFromBatch(response), response, item?.payout_item_id || null);
  res.redirect(`/receipt/${encodeURIComponent(req.params.id)}?refreshed=true`);
}));

router.post("/webhooks/paypal", asyncRoute(async (req, res) => {
  const verification = await paypal.verifyWebhook(req.headers, req.body);
  if (verification.verification_status !== "SUCCESS") {
    res.status(400).json({ error: "Invalid PayPal webhook signature." });
    return;
  }
  if (!(await recordWebhookOnce(req.body))) {
    res.sendStatus(200);
    return;
  }
  const resource = req.body.resource || {};
  const batchId = resource.payout_batch_id || resource.batch_header?.payout_batch_id;
  const status = resource.transaction_status || resource.batch_status;
  try {
    if (batchId && status) {
      await applyPayPalStatusByBatch(batchId, status, req.body, resource.payout_item_id || null);
    }
  } catch (error) {
    await releaseWebhookEvent(req.body.id);
    throw error;
  }
  res.sendStatus(200);
}));

router.get("/receipt/:id", asyncRoute(async (req, res) => {
  const record = await findRecordOrRedirect(req, res);

  if (!record) {
    return;
  }

  res.render("receipt", {
    pageTitle: "Receipt",
    activePage: "approval",
    record,
    paypalPayout: await getPayoutByInvoice(req.params.id),
    payoutSubmitted: req.query.submitted === "true",
    payoutRefreshed: req.query.refreshed === "true",
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

router.get("/reports/matching-history.xlsx", asyncRoute(async (req, res) => {
  const rows = await getReportRows();
  const workbook = createMatchingHistoryWorkbook(rows);
  const buffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });

  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", "attachment; filename=\"matching-history-report.xlsx\"");
  res.send(buffer);
}));

module.exports = router;
