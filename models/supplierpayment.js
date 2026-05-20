const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

const dataDirectory = path.join(__dirname, "..", "data");
const uploadsDirectory = path.join(__dirname, "..", "uploads");
const dataFile = path.join(dataDirectory, "supplier-payments.json");

const paymentStatuses = {
  pending: "Pending Approval",
  approved: "Approved for Payment",
  processing: "Payment Processing",
  paid: "Paid",
  rejected: "Rejected",
  held: "Payment Held",
};

const fieldsToCompare = [
  "supplierName",
  "poNumber",
  "itemDescription",
  "quantityOrdered",
  "quantityReceived",
  "quantityBilled",
  "unitPrice",
  "totalAmount",
];

const comparisonLabels = {
  supplierName: "Supplier Name",
  poNumber: "PO Reference Number",
  itemDescription: "Item Description",
  quantityOrdered: "Quantity Ordered vs Received",
  quantityReceived: "Quantity Received vs Billed",
  quantityBilled: "Quantity Ordered vs Billed",
  unitPrice: "Unit Price",
  totalAmount: "Total Amount",
};

function ensureDataFile() {
  fs.mkdirSync(dataDirectory, { recursive: true });
  fs.mkdirSync(uploadsDirectory, { recursive: true });

  if (!fs.existsSync(dataFile)) {
    fs.writeFileSync(dataFile, JSON.stringify([], null, 2));
  }
}

function readRecords() {
  ensureDataFile();

  try {
    const records = JSON.parse(fs.readFileSync(dataFile, "utf8"));
    return Array.isArray(records) ? records : [];
  } catch (error) {
    return [];
  }
}

function saveRecords(records) {
  ensureDataFile();
  fs.writeFileSync(dataFile, JSON.stringify(records, null, 2));
}

function createUploadedFile(file) {
  if (!file) {
    return null;
  }

  return {
    originalName: file.originalname,
    storedName: file.filename,
    path: file.path,
    mimetype: file.mimetype,
  };
}

function emptyPurchaseOrder() {
  return {
    supplierName: "",
    poNumber: "",
    itemDescription: "",
    quantityOrdered: "",
    unitPrice: "",
    totalAmount: "",
    documentDate: "",
  };
}

function emptyDeliveryOrder() {
  return {
    supplierName: "",
    poNumber: "",
    doGrnNumber: "",
    itemDescription: "",
    quantityReceived: "",
    documentDate: "",
  };
}

function emptyInvoice() {
  return {
    supplierName: "",
    poNumber: "",
    invoiceNumber: "",
    itemDescription: "",
    quantityBilled: "",
    unitPrice: "",
    totalAmount: "",
    documentDate: "",
  };
}

function getNextId(records) {
  return records.length ? Math.max(...records.map((record) => Number(record.id) || 0)) + 1 : 1;
}

function normaliseHeader(header) {
  return String(header || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function normaliseValue(value) {
  if (value === null || value === undefined) {
    return "";
  }

  if (value instanceof Date) {
    return value.toISOString().slice(0, 10);
  }

  return String(value).trim();
}

function toNumber(value) {
  if (value === null || value === undefined || String(value).trim() === "") {
    return Number.NaN;
  }

  return Number(String(value).replace(/[$,]/g, ""));
}

function money(value) {
  const number = toNumber(value);
  return Number.isNaN(number) ? 0 : number;
}

function formatDate(value) {
  if (!value) {
    return "";
  }

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }

  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
    }
  }

  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? String(value).trim() : date.toISOString().slice(0, 10);
}

function isValidDate(value) {
  return Boolean(value) && !Number.isNaN(Date.parse(value));
}

function readFirstExcelRow(file) {
  if (!file) {
    return null;
  }

  const extension = path.extname(file.originalName).toLowerCase();
  if (![".xlsx", ".xls", ".csv"].includes(extension)) {
    return null;
  }

  // Excel extraction: uploaded PO, DO/GRN, and invoice files are read with xlsx.
  // The first worksheet and first data row are used as the extracted document set.
  const workbook = XLSX.readFile(file.path, { cellDates: true });
  const firstSheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheetName];
  const rows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

  return rows[0] || null;
}

function getColumn(row, possibleNames) {
  if (!row) {
    return "";
  }

  const normalisedRow = {};
  Object.keys(row).forEach((key) => {
    normalisedRow[normaliseHeader(key)] = row[key];
  });

  for (const name of possibleNames) {
    const value = normalisedRow[normaliseHeader(name)];
    if (value !== undefined && value !== null && String(value).trim() !== "") {
      return value;
    }
  }

  return "";
}

function extractPurchaseOrder(file) {
  const row = readFirstExcelRow(file);

  if (!row) {
    return emptyPurchaseOrder();
  }

  return {
    supplierName: normaliseValue(getColumn(row, ["Supplier Name", "Vendor Name"])),
    poNumber: normaliseValue(getColumn(row, ["PO Reference Number", "PO Number", "PO No", "Order Number", "Order No"])),
    itemDescription: normaliseValue(getColumn(row, ["Item Description", "Description", "Item Name", "Product"])),
    quantityOrdered: normaliseValue(getColumn(row, ["Quantity Ordered", "Ordered Quantity", "PO Quantity", "Quantity", "Qty"])),
    unitPrice: normaliseValue(getColumn(row, ["Unit Price", "Price"])),
    totalAmount: normaliseValue(getColumn(row, ["Total Amount", "Total Price", "Amount"])),
    documentDate: formatDate(getColumn(row, ["Document Date", "PO Date", "Order Date", "Date"])),
  };
}

function extractDeliveryOrder(file) {
  const row = readFirstExcelRow(file);

  if (!row) {
    return emptyDeliveryOrder();
  }

  return {
    supplierName: normaliseValue(getColumn(row, ["Supplier Name", "Vendor Name"])),
    poNumber: normaliseValue(getColumn(row, ["PO Reference Number", "PO Number", "PO No", "Order Number", "Order No"])),
    doGrnNumber: normaliseValue(getColumn(row, ["DO/GRN Number", "GRN Number", "GRN No", "DO Number", "DO No", "Delivery Order Number"])),
    itemDescription: normaliseValue(getColumn(row, ["Item Description", "Description", "Item Name", "Product"])),
    quantityReceived: normaliseValue(getColumn(row, ["Quantity Received", "Received Quantity", "GRN Quantity", "DO Quantity", "Quantity", "Qty"])),
    documentDate: formatDate(getColumn(row, ["Document Date", "GRN Date", "DO Date", "Delivery Date", "Date"])),
  };
}

function extractInvoice(file) {
  const row = readFirstExcelRow(file);

  if (!row) {
    return emptyInvoice();
  }

  return {
    supplierName: normaliseValue(getColumn(row, ["Supplier Name", "Vendor Name"])),
    poNumber: normaliseValue(getColumn(row, ["PO Reference Number", "PO Number", "PO No", "Order Number", "Order No"])),
    invoiceNumber: normaliseValue(getColumn(row, ["Invoice Number", "Invoice No", "Supplier Invoice Number"])),
    itemDescription: normaliseValue(getColumn(row, ["Item Description", "Description", "Item Name", "Product"])),
    quantityBilled: normaliseValue(getColumn(row, ["Quantity Billed", "Billed Quantity", "Invoice Quantity", "Quantity", "Qty"])),
    unitPrice: normaliseValue(getColumn(row, ["Unit Price", "Price"])),
    totalAmount: normaliseValue(getColumn(row, ["Total Amount", "Item Total", "Subtotal", "Total Price", "Amount"])),
    documentDate: formatDate(getColumn(row, ["Document Date", "Invoice Date", "Order Date", "Date"])),
  };
}

function migrateLegacyRecord(record) {
  if (record.purchaseOrder?.quantityOrdered || record.deliveryOrder || record.invoice?.quantityBilled) {
    return {
      ...record,
      deliveryOrderFile: record.deliveryOrderFile || record.doGrnFile || null,
      deliveryOrder: record.deliveryOrder || emptyDeliveryOrder(),
    };
  }

  const purchaseOrder = record.purchaseOrder || {};
  const invoice = record.invoice || {};

  return {
    ...record,
    deliveryOrderFile: record.deliveryOrderFile || null,
    purchaseOrder: {
      supplierName: purchaseOrder.supplierName || "",
      poNumber: purchaseOrder.poNumber || "",
      itemDescription: purchaseOrder.itemDescription || purchaseOrder.description || purchaseOrder.itemName || "",
      quantityOrdered: purchaseOrder.quantityOrdered || purchaseOrder.quantity || "",
      unitPrice: purchaseOrder.unitPrice || "",
      totalAmount: purchaseOrder.totalAmount || purchaseOrder.totalPrice || "",
      documentDate: purchaseOrder.documentDate || purchaseOrder.orderDate || "",
    },
    deliveryOrder: record.deliveryOrder || emptyDeliveryOrder(),
    invoice: {
      supplierName: invoice.supplierName || "",
      poNumber: invoice.poNumber || invoice.orderNumber || "",
      invoiceNumber: invoice.invoiceNumber || "",
      itemDescription: invoice.itemDescription || invoice.description || invoice.itemName || "",
      quantityBilled: invoice.quantityBilled || invoice.quantity || "",
      unitPrice: invoice.unitPrice || "",
      totalAmount: invoice.totalAmount || invoice.itemTotal || invoice.subtotal || "",
      documentDate: invoice.documentDate || invoice.orderDate || "",
    },
  };
}

function createUploadRecord(files) {
  const records = readRecords();
  const poFile = createUploadedFile(files.poFile?.[0]);
  const deliveryOrderFile = createUploadedFile(files.doGrnFile?.[0]);
  const invoiceFile = createUploadedFile(files.invoiceFile?.[0]);

  const record = {
    id: getNextId(records),
    createdAt: new Date().toISOString(),
    purchaseOrderFile: poFile,
    deliveryOrderFile,
    invoiceFile,
    purchaseOrder: extractPurchaseOrder(poFile),
    deliveryOrder: extractDeliveryOrder(deliveryOrderFile),
    invoice: extractInvoice(invoiceFile),
    extractionStatus: poFile && deliveryOrderFile && invoiceFile ? "Extracted" : "Needs Review",
    paymentStatus: paymentStatuses.pending,
    approvalStatus: "Not Approved",
    payment: null,
  };

  records.push(record);
  saveRecords(records);
  return record.id;
}

function getRawRecord(id) {
  const record = readRecords().find((savedRecord) => Number(savedRecord.id) === Number(id));
  return record ? migrateLegacyRecord(record) : null;
}

function updateRecord(id, updater) {
  const records = readRecords();
  const index = records.findIndex((record) => Number(record.id) === Number(id));

  if (index === -1) {
    return null;
  }

  records[index] = updater(migrateLegacyRecord(records[index]));
  saveRecords(records);
  return records[index];
}

function saveCorrectedData(id, formData) {
  return updateRecord(id, (record) => ({
    ...record,
    purchaseOrder: {
      supplierName: formData.poSupplierName || "",
      poNumber: formData.poNumber || "",
      itemDescription: formData.poItemDescription || "",
      quantityOrdered: formData.poQuantityOrdered || "",
      unitPrice: formData.poUnitPrice || "",
      totalAmount: formData.poTotalAmount || "",
      documentDate: formData.poDocumentDate || "",
    },
    deliveryOrder: {
      supplierName: formData.doSupplierName || "",
      poNumber: formData.doPoNumber || "",
      doGrnNumber: formData.doGrnNumber || "",
      itemDescription: formData.doItemDescription || "",
      quantityReceived: formData.doQuantityReceived || "",
      documentDate: formData.doDocumentDate || "",
    },
    invoice: {
      supplierName: formData.invoiceSupplierName || "",
      poNumber: formData.invoicePoNumber || "",
      invoiceNumber: formData.invoiceNumber || "",
      itemDescription: formData.invoiceItemDescription || "",
      quantityBilled: formData.invoiceQuantityBilled || "",
      unitPrice: formData.invoiceUnitPrice || "",
      totalAmount: formData.invoiceTotalAmount || "",
      documentDate: formData.invoiceDocumentDate || "",
    },
    extractionStatus: "Reviewed",
  }));
}

function pushIssue(issues, invalidFields, field, message) {
  issues.push(message);
  invalidFields.push(field);
}

function validateTotal(issues, invalidFields, field, quantity, unitPrice, total, label) {
  const expected = Number((money(quantity) * money(unitPrice)).toFixed(2));
  const actual = money(total);

  if (!(actual > 0) || Math.abs(actual - expected) > 0.01) {
    pushIssue(issues, invalidFields, field, `${label} total amount must equal quantity x unit price.`);
  }
}

function calculateValidation(record) {
  const invalidFields = [];
  const messages = [];
  const po = record.purchaseOrder || emptyPurchaseOrder();
  const deliveryOrder = record.deliveryOrder || emptyDeliveryOrder();
  const invoice = record.invoice || emptyInvoice();

  if (!record.purchaseOrderFile || !record.deliveryOrderFile || !record.invoiceFile) {
    return {
      status: "Needs Review",
      invalidFields,
      messages: ["PO, DO/GRN, and Supplier Invoice files are all required."],
    };
  }

  [
    ["po.supplierName", po.supplierName, "PO supplier name is required."],
    ["po.poNumber", po.poNumber, "PO reference number is required."],
    ["po.itemDescription", po.itemDescription, "PO item description is required."],
    ["po.documentDate", po.documentDate, "PO document date is required."],
    ["do.supplierName", deliveryOrder.supplierName, "DO/GRN supplier name is required."],
    ["do.doGrnNumber", deliveryOrder.doGrnNumber, "DO/GRN number is required."],
    ["do.poNumber", deliveryOrder.poNumber, "DO/GRN PO reference number is required."],
    ["do.itemDescription", deliveryOrder.itemDescription, "DO/GRN item description is required."],
    ["do.documentDate", deliveryOrder.documentDate, "DO/GRN document date is required."],
    ["invoice.invoiceNumber", invoice.invoiceNumber, "Invoice number is required."],
    ["invoice.poNumber", invoice.poNumber, "Invoice PO reference number is required."],
    ["invoice.supplierName", invoice.supplierName, "Invoice supplier name is required."],
    ["invoice.itemDescription", invoice.itemDescription, "Invoice item description is required."],
    ["invoice.documentDate", invoice.documentDate, "Invoice document date is required."],
  ].forEach(([field, value, message]) => {
    if (String(value || "").trim() === "") {
      pushIssue(messages, invalidFields, field, message);
    }
  });

  [
    ["po.quantityOrdered", po.quantityOrdered, "Quantity ordered must be positive."],
    ["do.quantityReceived", deliveryOrder.quantityReceived, "Quantity received must be positive."],
    ["invoice.quantityBilled", invoice.quantityBilled, "Quantity billed must be positive."],
    ["po.unitPrice", po.unitPrice, "PO unit price must be positive."],
    ["invoice.unitPrice", invoice.unitPrice, "Invoice unit price must be positive."],
  ].forEach(([field, value, message]) => {
    if (!(toNumber(value) > 0)) {
      pushIssue(messages, invalidFields, field, message);
    }
  });

  validateTotal(messages, invalidFields, "po.totalAmount", po.quantityOrdered, po.unitPrice, po.totalAmount, "PO");
  validateTotal(messages, invalidFields, "invoice.totalAmount", invoice.quantityBilled, invoice.unitPrice, invoice.totalAmount, "Invoice");

  [
    ["po.documentDate", po.documentDate, "PO document date must be valid."],
    ["do.documentDate", deliveryOrder.documentDate, "DO/GRN document date must be valid."],
    ["invoice.documentDate", invoice.documentDate, "Invoice document date must be valid."],
  ].forEach(([field, value, message]) => {
    if (!isValidDate(value)) {
      pushIssue(messages, invalidFields, field, message);
    }
  });

  return {
    status: invalidFields.length ? "Invalid" : "Valid",
    invalidFields: [...new Set(invalidFields)],
    messages: invalidFields.length ? [...new Set(messages)] : ["All validation checks passed."],
  };
}

function valuesMatch(firstValue, secondValue) {
  const firstNumber = toNumber(firstValue);
  const secondNumber = toNumber(secondValue);
  const firstText = String(firstValue || "").trim();
  const secondText = String(secondValue || "").trim();

  if (firstText === "" || secondText === "") {
    return false;
  }

  if (!Number.isNaN(firstNumber) && !Number.isNaN(secondNumber)) {
    return Math.abs(firstNumber - secondNumber) <= 0.01;
  }

  return firstText.toLowerCase() === secondText.toLowerCase();
}

function discrepancyAction(type) {
  if (type === "Missing document") {
    return "Request missing document";
  }

  if (type === "Quantity mismatch") {
    return "Hold payment and review delivery record";
  }

  if (type === "Price mismatch") {
    return "Review supplier invoice against PO";
  }

  return "Review extracted document data";
}

function makeDiscrepancy(record, field, type, poValue, doValue, invoiceValue) {
  const po = record.purchaseOrder || emptyPurchaseOrder();
  const deliveryOrder = record.deliveryOrder || emptyDeliveryOrder();
  const invoice = record.invoice || emptyInvoice();

  return {
    supplierName: invoice.supplierName || po.supplierName || deliveryOrder.supplierName || "Missing",
    poNumber: po.poNumber || deliveryOrder.poNumber || invoice.poNumber || "Missing",
    doGrnNumber: deliveryOrder.doGrnNumber || "Missing",
    invoiceNumber: invoice.invoiceNumber || "Missing",
    field,
    poValue: normaliseValue(poValue) || "Missing",
    doGrnValue: normaliseValue(doValue) || "Missing",
    invoiceValue: normaliseValue(invoiceValue) || "Missing",
    type,
    action: discrepancyAction(type),
    resolved: false,
  };
}

function calculateDiscrepancies(record) {
  const po = record.purchaseOrder || emptyPurchaseOrder();
  const deliveryOrder = record.deliveryOrder || emptyDeliveryOrder();
  const invoice = record.invoice || emptyInvoice();
  const discrepancies = [];

  // Discrepancy logic: every row is created from uploaded and saved document data.
  if (!record.purchaseOrderFile) {
    discrepancies.push(makeDiscrepancy(record, "Purchase Order", "Missing document", "", "", ""));
  }

  if (!record.deliveryOrderFile) {
    discrepancies.push(makeDiscrepancy(record, "DO/GRN", "Missing document", "", "", ""));
  }

  if (!record.invoiceFile) {
    discrepancies.push(makeDiscrepancy(record, "Supplier Invoice", "Missing document", "", "", ""));
  }

  if (!valuesMatch(po.supplierName, invoice.supplierName) || !valuesMatch(po.supplierName, deliveryOrder.supplierName)) {
    discrepancies.push(makeDiscrepancy(record, "Supplier Name", "Supplier mismatch", po.supplierName, deliveryOrder.supplierName, invoice.supplierName));
  }

  if (!valuesMatch(po.poNumber, deliveryOrder.poNumber) || !valuesMatch(po.poNumber, invoice.poNumber)) {
    discrepancies.push(makeDiscrepancy(record, "PO Reference Number", "Reference mismatch", po.poNumber, deliveryOrder.poNumber, invoice.poNumber));
  }

  if (!valuesMatch(po.itemDescription, deliveryOrder.itemDescription) || !valuesMatch(po.itemDescription, invoice.itemDescription)) {
    discrepancies.push(makeDiscrepancy(record, "Item Description", "Item mismatch", po.itemDescription, deliveryOrder.itemDescription, invoice.itemDescription));
  }

  if (!valuesMatch(po.quantityOrdered, deliveryOrder.quantityReceived)) {
    discrepancies.push(makeDiscrepancy(record, "Quantity Ordered vs Received", "Quantity mismatch", po.quantityOrdered, deliveryOrder.quantityReceived, ""));
  }

  if (!valuesMatch(po.quantityOrdered, invoice.quantityBilled)) {
    discrepancies.push(makeDiscrepancy(record, "Quantity Ordered vs Billed", "Quantity mismatch", po.quantityOrdered, "", invoice.quantityBilled));
  }

  if (!valuesMatch(deliveryOrder.quantityReceived, invoice.quantityBilled)) {
    discrepancies.push(makeDiscrepancy(record, "Quantity Received vs Billed", "Quantity mismatch", "", deliveryOrder.quantityReceived, invoice.quantityBilled));
  }

  if (!valuesMatch(po.unitPrice, invoice.unitPrice)) {
    discrepancies.push(makeDiscrepancy(record, "Unit Price", "Price mismatch", po.unitPrice, "", invoice.unitPrice));
  }

  if (!valuesMatch(po.totalAmount, invoice.totalAmount)) {
    discrepancies.push(makeDiscrepancy(record, "Total Amount", "Amount mismatch", po.totalAmount, "", invoice.totalAmount));
  }

  return discrepancies;
}

function calculateMatching(record) {
  if (!record.purchaseOrderFile || !record.deliveryOrderFile || !record.invoiceFile) {
    return {
      status: "Pending Review",
      rows: [],
      mismatchFields: [],
      message: "A PO, DO/GRN, or Supplier Invoice document is missing.",
    };
  }

  const po = record.purchaseOrder || emptyPurchaseOrder();
  const deliveryOrder = record.deliveryOrder || emptyDeliveryOrder();
  const invoice = record.invoice || emptyInvoice();

  const rows = [
    {
      field: "supplierName",
      label: comparisonLabels.supplierName,
      poValue: po.supplierName,
      doGrnValue: deliveryOrder.supplierName,
      invoiceValue: invoice.supplierName,
      result: valuesMatch(po.supplierName, invoice.supplierName) && valuesMatch(po.supplierName, deliveryOrder.supplierName) ? "Match" : "Mismatch",
    },
    {
      field: "poNumber",
      label: comparisonLabels.poNumber,
      poValue: po.poNumber,
      doGrnValue: deliveryOrder.poNumber,
      invoiceValue: invoice.poNumber,
      result: valuesMatch(po.poNumber, deliveryOrder.poNumber) && valuesMatch(po.poNumber, invoice.poNumber) ? "Match" : "Mismatch",
    },
    {
      field: "itemDescription",
      label: comparisonLabels.itemDescription,
      poValue: po.itemDescription,
      doGrnValue: deliveryOrder.itemDescription,
      invoiceValue: invoice.itemDescription,
      result: valuesMatch(po.itemDescription, deliveryOrder.itemDescription) && valuesMatch(po.itemDescription, invoice.itemDescription) ? "Match" : "Mismatch",
    },
    {
      field: "quantityOrdered",
      label: comparisonLabels.quantityOrdered,
      poValue: po.quantityOrdered,
      doGrnValue: deliveryOrder.quantityReceived,
      invoiceValue: "-",
      result: valuesMatch(po.quantityOrdered, deliveryOrder.quantityReceived) ? "Match" : "Mismatch",
    },
    {
      field: "quantityBilled",
      label: comparisonLabels.quantityBilled,
      poValue: po.quantityOrdered,
      doGrnValue: "-",
      invoiceValue: invoice.quantityBilled,
      result: valuesMatch(po.quantityOrdered, invoice.quantityBilled) ? "Match" : "Mismatch",
    },
    {
      field: "quantityReceived",
      label: comparisonLabels.quantityReceived,
      poValue: "-",
      doGrnValue: deliveryOrder.quantityReceived,
      invoiceValue: invoice.quantityBilled,
      result: valuesMatch(deliveryOrder.quantityReceived, invoice.quantityBilled) ? "Match" : "Mismatch",
    },
    {
      field: "unitPrice",
      label: comparisonLabels.unitPrice,
      poValue: po.unitPrice,
      doGrnValue: "-",
      invoiceValue: invoice.unitPrice,
      result: valuesMatch(po.unitPrice, invoice.unitPrice) ? "Match" : "Mismatch",
    },
    {
      field: "totalAmount",
      label: comparisonLabels.totalAmount,
      poValue: po.totalAmount,
      doGrnValue: "-",
      invoiceValue: invoice.totalAmount,
      result: valuesMatch(po.totalAmount, invoice.totalAmount) ? "Match" : "Mismatch",
    },
  ].map((row) => ({
    ...row,
    poValue: normaliseValue(row.poValue) || "Missing",
    doGrnValue: normaliseValue(row.doGrnValue) || "Missing",
    invoiceValue: normaliseValue(row.invoiceValue) || "Missing",
  }));

  const mismatchFields = rows.filter((row) => row.result === "Mismatch").map((row) => row.field);

  return {
    status: mismatchFields.length ? "Mismatch" : "Matched",
    rows,
    mismatchFields,
    message: mismatchFields.length ? "One or more 3-way matching checks failed." : "PO, DO/GRN, and Supplier Invoice match.",
  };
}

function decorateRecord(record) {
  const migratedRecord = migrateLegacyRecord(record);
  const validation = calculateValidation(migratedRecord);
  const matching = calculateMatching(migratedRecord);
  const discrepancies = calculateDiscrepancies(migratedRecord);
  const unresolvedDiscrepancies = discrepancies.filter((discrepancy) => !discrepancy.resolved);
  const canApprovePayment = matching.status === "Matched" && validation.status === "Valid" && unresolvedDiscrepancies.length === 0;

  return {
    ...migratedRecord,
    validation,
    matching,
    discrepancies,
    unresolvedDiscrepancies,
    canApprovePayment,
    approvalStatus: migratedRecord.approvalStatus || (migratedRecord.paymentStatus === paymentStatuses.approved ? "Approved" : "Not Approved"),
    amountPayable: migratedRecord.invoice?.totalAmount || migratedRecord.purchaseOrder?.totalAmount || "0",
  };
}

function getRecords() {
  return readRecords()
    .map(decorateRecord)
    .sort((first, second) => Number(second.id) - Number(first.id));
}

function getRecord(id) {
  const record = getRawRecord(id);
  return record ? decorateRecord(record) : null;
}

function setPaymentStatus(id, paymentStatus) {
  return updateRecord(id, (record) => ({
    ...record,
    paymentStatus,
    approvalStatus: paymentStatus === paymentStatuses.approved ? "Approved" : record.approvalStatus || "Not Approved",
    paymentUpdatedAt: new Date().toISOString(),
  }));
}

function approvePayment(id) {
  const record = getRecord(id);

  if (!record || !record.canApprovePayment) {
    return record ? setPaymentStatus(id, paymentStatuses.held) : null;
  }

  return updateRecord(id, (savedRecord) => ({
    ...savedRecord,
    paymentStatus: paymentStatuses.approved,
    approvalStatus: "Approved",
    approvedAt: new Date().toISOString(),
  }));
}

function rejectPayment(id) {
  return updateRecord(id, (record) => ({
    ...record,
    paymentStatus: paymentStatuses.rejected,
    approvalStatus: "Rejected",
    rejectedAt: new Date().toISOString(),
  }));
}

function getNextTransactionId(records) {
  const year = new Date().getFullYear();
  const paidCount = records.filter((record) => record.payment?.transactionId).length + 1;
  return `PAY-${year}-${String(paidCount).padStart(4, "0")}`;
}

function simulatePayment(id, paymentMethod = "Demo Bank Transfer") {
  const records = readRecords();
  const index = records.findIndex((record) => Number(record.id) === Number(id));

  if (index === -1) {
    return null;
  }

  const decorated = decorateRecord(records[index]);
  if (decorated.paymentStatus !== paymentStatuses.approved && decorated.paymentStatus !== paymentStatuses.processing) {
    return decorated;
  }

  // Payment logic: this is only a demo state transition, not a payment gateway.
  records[index] = {
    ...migrateLegacyRecord(records[index]),
    paymentStatus: paymentStatuses.paid,
    approvalStatus: "Approved",
    payment: {
      transactionId: getNextTransactionId(records),
      method: paymentMethod,
      amountPaid: decorated.amountPayable,
      paidAt: new Date().toISOString(),
      status: paymentStatuses.paid,
    },
  };

  saveRecords(records);
  return decorateRecord(records[index]);
}

function getPaymentList() {
  return getRecords().filter((record) => record.canApprovePayment);
}

function getAllDiscrepancies() {
  return getRecords().flatMap((record) =>
    record.discrepancies.map((discrepancy) => ({
      recordId: record.id,
      ...discrepancy,
    }))
  );
}

function getStats() {
  const records = getRecords();
  const paidRecords = records.filter((record) => record.paymentStatus === paymentStatuses.paid);

  return {
    totalRecords: records.length,
    totalUploadedDocumentSets: records.length,
    matched: records.filter((record) => record.matching.status === "Matched").length,
    mismatched: records.filter((record) => record.matching.status === "Mismatch").length,
    pendingReview: records.filter((record) => record.matching.status === "Pending Review").length,
    discrepancies: getAllDiscrepancies().length,
    valid: records.filter((record) => record.validation.status === "Valid").length,
    invalid: records.filter((record) => record.validation.status === "Invalid").length,
    pendingApproval: records.filter((record) => record.paymentStatus === paymentStatuses.pending).length,
    pendingPayments: records.filter((record) => record.paymentStatus === paymentStatuses.pending).length,
    approvedPayments: records.filter((record) => record.paymentStatus === paymentStatuses.approved).length,
    rejectedPayments: records.filter((record) => record.paymentStatus === paymentStatuses.rejected).length,
    paidInvoices: paidRecords.length,
    totalAmountPaid: paidRecords.reduce((total, record) => total + money(record.payment?.amountPaid), 0),
  };
}

function getReportRows() {
  return getRecords().map((record) => ({
    id: record.id,
    supplierName: record.invoice?.supplierName || record.purchaseOrder?.supplierName || "Needs Review",
    poNumber: record.purchaseOrder?.poNumber || "Missing",
    doGrnNumber: record.deliveryOrder?.doGrnNumber || "Missing",
    invoiceNumber: record.invoice?.invoiceNumber || "Missing",
    amount: record.amountPayable,
    matchStatus: record.matching.status,
    validationStatus: record.validation.status,
    approvalStatus: record.approvalStatus,
    paymentStatus: record.paymentStatus,
    transactionId: record.payment?.transactionId || "-",
    paymentMethod: record.payment?.method || "-",
    paidAt: record.payment?.paidAt || "",
  }));
}

module.exports = {
  approvePayment,
  createUploadRecord,
  fieldsToCompare,
  getAllDiscrepancies,
  getPaymentList,
  getRecord,
  getRecords,
  getReportRows,
  getStats,
  paymentStatuses,
  rejectPayment,
  saveCorrectedData,
  setPaymentStatus,
  simulatePayment,
  ensureDataFile,
};
