const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const db = require("../db");

const uploadsDirectory = path.join(__dirname, "..", "uploads");

const paymentStatuses = {
  pending: "Pending Approval",
  approved: "Approved for Payment",
  processing: "Payment Processing",
  paid: "Paid",
  rejected: "Rejected",
  held: "Payment Held",
};

const dbPaymentStatuses = {
  pending: "PENDING_APPROVAL",
  approved: "APPROVED_FOR_PAYMENT",
  processing: "PAYMENT_PROCESSING",
  paid: "PAID",
  rejected: "REJECTED",
  held: "PAYMENT_HELD",
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
  fs.mkdirSync(uploadsDirectory, { recursive: true });
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

  const workbook = XLSX.readFile(file.path, { cellDates: true });
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
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

function addDays(dateValue, days) {
  const date = new Date(`${dateValue}T00:00:00`);
  date.setDate(date.getDate() + Number(days || 0));
  return date.toISOString().slice(0, 10);
}

function eomPlusDays(dateValue, days) {
  const date = new Date(`${dateValue}T00:00:00`);
  const endOfMonth = new Date(date.getFullYear(), date.getMonth() + 1, 0);
  endOfMonth.setDate(endOfMonth.getDate() + Number(days || 0));
  return endOfMonth.toISOString().slice(0, 10);
}

function sqlDate(value) {
  return formatDate(value) || new Date().toISOString().slice(0, 10);
}

function makeItemId(value) {
  const cleaned = String(value || "ITEM")
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, "")
    .slice(0, 16);
  return cleaned || "ITEM";
}

function dbPaymentStatusToUi(status, exceptionFlag) {
  if (status === dbPaymentStatuses.paid || status === paymentStatuses.paid) {
    return paymentStatuses.paid;
  }

  if (status === dbPaymentStatuses.rejected || status === paymentStatuses.rejected) {
    return paymentStatuses.rejected;
  }

  if (status === dbPaymentStatuses.processing || status === paymentStatuses.processing) {
    return paymentStatuses.processing;
  }

  if (status === dbPaymentStatuses.approved || status === paymentStatuses.approved) {
    return paymentStatuses.approved;
  }

  if (status === dbPaymentStatuses.pending || status === paymentStatuses.pending) {
    return paymentStatuses.pending;
  }

  if (status === "READY" && exceptionFlag !== "Y") {
    return paymentStatuses.approved;
  }

  return paymentStatuses.held;
}

function uiPaymentStatusToDb(status) {
  const statusMap = {
    [paymentStatuses.pending]: dbPaymentStatuses.pending,
    [paymentStatuses.approved]: dbPaymentStatuses.approved,
    [paymentStatuses.processing]: dbPaymentStatuses.processing,
    [paymentStatuses.paid]: dbPaymentStatuses.paid,
    [paymentStatuses.rejected]: dbPaymentStatuses.rejected,
    [paymentStatuses.held]: dbPaymentStatuses.held,
  };

  return statusMap[status] || dbPaymentStatuses.held;
}

function approvalStatusLabel(paymentStatus) {
  if (paymentStatus === paymentStatuses.rejected) {
    return "Rejected";
  }

  if (paymentStatus === paymentStatuses.paid) {
    return "Paid";
  }

  if ([paymentStatuses.approved, paymentStatuses.processing].includes(paymentStatus)) {
    return "Approved";
  }

  return "Not Approved";
}

let paymentStatusColumnReady;

async function ensurePaymentStatusColumnSupportsWorkflow() {
  if (!paymentStatusColumnReady) {
    paymentStatusColumnReady = (async () => {
      const [rows] = await db.execute(
        `
          SELECT COLUMN_TYPE
          FROM INFORMATION_SCHEMA.COLUMNS
          WHERE TABLE_SCHEMA = DATABASE()
            AND TABLE_NAME = 'payment_due_list'
            AND COLUMN_NAME = 'payment_status'
          LIMIT 1
        `
      );
      const columnType = rows[0]?.COLUMN_TYPE || "";
      const requiredValues = Object.values(dbPaymentStatuses);

      if (requiredValues.every((status) => columnType.includes(`'${status}'`))) {
        return;
      }

      await db.execute(
        `
          ALTER TABLE payment_due_list
          MODIFY payment_status ENUM(
            'READY',
            'HOLD',
            'PAID',
            'PENDING_APPROVAL',
            'APPROVED_FOR_PAYMENT',
            'PAYMENT_PROCESSING',
            'REJECTED',
            'PAYMENT_HELD'
          ) NOT NULL DEFAULT 'PAYMENT_HELD'
        `
      );
    })();
  }

  return paymentStatusColumnReady;
}

async function queryRecords(whereClause = "", values = []) {
  const [rows] = await db.execute(
    `
      SELECT
        si.invoice_id,
        si.supplier_id,
        si.po_id,
        si.do_id,
        si.invoice_date,
        si.item_id AS invoice_item_id,
        si.qty_invoiced,
        si.unit_price AS invoice_unit_price,
        si.tax_amount,
        si.total_amount AS invoice_total_amount,
        si.currency AS invoice_currency,
        sm.supplier_name,
        sm.active_flag,
        pt.days AS payment_term_days,
        pt.type AS payment_term_type,
        po.po_date,
        po.description AS po_description,
        po.qty_ordered,
        po.unit_price AS po_unit_price,
        po.currency AS po_currency,
        po.total_amount AS po_total_amount,
        dox.delivery_date,
        dox.qty_delivered,
        pdl.payment_due_id,
        pdl.due_date,
        pdl.amount_due,
        pdl.payment_status,
        pdl.exception_flag
      FROM supplier_invoices si
      INNER JOIN supplier_master sm ON sm.supplier_id = si.supplier_id
      INNER JOIN payment_terms pt ON pt.term_code = sm.payment_term_code
      INNER JOIN purchase_orders po ON po.po_id = si.po_id
      LEFT JOIN delivery_orders dox ON dox.do_id = si.do_id
      LEFT JOIN payment_due_list pdl ON pdl.invoice_id = si.invoice_id
      ${whereClause}
      ORDER BY si.invoice_date DESC, si.invoice_id DESC
    `,
    values
  );

  const invoiceIds = rows.map((row) => row.invoice_id);
  let exceptionsByInvoice = new Map();

  if (invoiceIds.length) {
    const placeholders = invoiceIds.map(() => "?").join(", ");
    const [exceptions] = await db.execute(
      `
        SELECT exception_id, invoice_id, exception_type, description, created_at
        FROM matching_exceptions
        WHERE invoice_id IN (${placeholders})
        ORDER BY exception_id ASC
      `,
      invoiceIds
    );

    exceptionsByInvoice = exceptions.reduce((map, exception) => {
      const current = map.get(exception.invoice_id) || [];
      current.push(exception);
      map.set(exception.invoice_id, current);
      return map;
    }, new Map());
  }

  return rows.map((row) => decorateRecord(rowToRecord(row, exceptionsByInvoice.get(row.invoice_id) || [])));
}

function rowToRecord(row, exceptions) {
  const paymentStatus = dbPaymentStatusToUi(row.payment_status, row.exception_flag || (exceptions.length ? "Y" : "N"));

  return {
    id: row.invoice_id,
    createdAt: row.invoice_date,
    purchaseOrderFile: { originalName: `${row.po_id}.sql`, storedName: row.po_id },
    deliveryOrderFile: row.do_id ? { originalName: `${row.do_id}.sql`, storedName: row.do_id } : null,
    invoiceFile: { originalName: `${row.invoice_id}.sql`, storedName: row.invoice_id },
    supplierId: row.supplier_id,
    purchaseOrder: {
      supplierName: row.supplier_name,
      poNumber: row.po_id,
      itemDescription: row.po_description,
      quantityOrdered: normaliseValue(row.qty_ordered),
      unitPrice: normaliseValue(row.po_unit_price),
      totalAmount: normaliseValue(row.po_total_amount),
      documentDate: row.po_date,
    },
    deliveryOrder: {
      supplierName: row.supplier_name,
      poNumber: row.po_id,
      doGrnNumber: row.do_id || "",
      itemDescription: row.po_description,
      quantityReceived: normaliseValue(row.qty_delivered),
      documentDate: row.delivery_date || "",
    },
    invoice: {
      supplierName: row.supplier_name,
      poNumber: row.po_id,
      invoiceNumber: row.invoice_id,
      itemDescription: row.po_description,
      quantityBilled: normaliseValue(row.qty_invoiced),
      unitPrice: normaliseValue(row.invoice_unit_price),
      totalAmount: normaliseValue(row.invoice_total_amount),
      documentDate: row.invoice_date,
    },
    dbPaymentStatus: row.payment_status || dbPaymentStatuses.held,
    dbExceptionFlag: row.exception_flag || (exceptions.length ? "Y" : "N"),
    extractionStatus: row.do_id ? "Extracted" : "Needs Review",
    paymentStatus,
    approvalStatus: approvalStatusLabel(paymentStatus),
    payment: paymentStatus === paymentStatuses.paid
      ? {
          transactionId: `PAY-${row.payment_due_id || row.invoice_id}`,
          method: "Demo Bank Transfer",
          amountPaid: row.amount_due || row.invoice_total_amount,
          paidAt: row.due_date,
          status: paymentStatuses.paid,
        }
      : null,
    databaseExceptions: exceptions,
  };
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

  if (!record.deliveryOrderFile) {
    discrepancies.push(makeDiscrepancy(record, "DO/GRN", "Missing document", "", "", ""));
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

  (record.databaseExceptions || []).forEach((exception) => {
    discrepancies.push(makeDiscrepancy(record, exception.description, exception.exception_type, "", "", ""));
  });

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
  const validation = calculateValidation(record);
  const matching = calculateMatching(record);
  const discrepancies = calculateDiscrepancies(record);
  const unresolvedDiscrepancies = discrepancies.filter((discrepancy) => !discrepancy.resolved);
  const canApprovePayment = matching.status === "Matched" && validation.status === "Valid" && unresolvedDiscrepancies.length === 0;

  return {
    ...record,
    validation,
    matching,
    discrepancies,
    unresolvedDiscrepancies,
    canApprovePayment,
    amountPayable: record.invoice?.totalAmount || record.purchaseOrder?.totalAmount || "0",
  };
}

async function getRecords() {
  return queryRecords();
}

async function getRecord(id) {
  const records = await queryRecords("WHERE si.invoice_id = ?", [id]);
  return records[0] || null;
}

function makeSupplierId() {
  const timestamp = Date.now().toString(36).toUpperCase();
  const random = Math.random().toString(36).slice(2, 6).toUpperCase();
  return `SUP${timestamp}${random}`.slice(0, 20);
}

async function findSupplierByNameOrId(connection, supplierName) {
  const [rows] = await connection.execute(
    `
      SELECT sm.supplier_id, sm.supplier_name, sm.currency, pt.days, pt.type
      FROM supplier_master sm
      INNER JOIN payment_terms pt ON pt.term_code = sm.payment_term_code
      WHERE sm.supplier_name = ? OR sm.supplier_id = ?
      LIMIT 1
    `,
    [supplierName, supplierName]
  );
  return rows[0] || null;
}

async function findOrCreateSupplier(connection, supplierName) {
  const cleanedSupplierName = normaliseValue(supplierName);

  if (!cleanedSupplierName) {
    throw new Error("Supplier Name is required before saving uploaded documents.");
  }

  const existingSupplier = await findSupplierByNameOrId(connection, cleanedSupplierName);

  if (existingSupplier) {
    return {
      supplier: existingSupplier,
      created: false,
    };
  }

  const supplierId = makeSupplierId();

  await connection.execute(
    `
      INSERT INTO supplier_master
        (supplier_id, supplier_name, currency, payment_term_code, credit_limit, tax_id, bank_account, bank_name, active_flag)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    `,
    [supplierId, cleanedSupplierName, "SGD", "NET30", null, null, null, null, "Y"]
  );

  const supplier = await findSupplierByNameOrId(connection, cleanedSupplierName);

  return {
    supplier,
    created: true,
  };
}

async function saveDocumentSet(record) {
  const po = record.purchaseOrder || emptyPurchaseOrder();
  const deliveryOrder = record.deliveryOrder || emptyDeliveryOrder();
  const invoice = record.invoice || emptyInvoice();
  const poId = po.poNumber || `PO${Date.now()}`;
  const invoiceId = invoice.invoiceNumber || `INV${Date.now()}`;
  const doId = deliveryOrder.doGrnNumber || null;
  const itemId = makeItemId(po.itemDescription || invoice.itemDescription);
  const poDate = sqlDate(po.documentDate);
  const invoiceDate = sqlDate(invoice.documentDate);
  await ensurePaymentStatusColumnSupportsWorkflow();
  const connection = await db.getConnection();

  try {
    await connection.beginTransaction();
    const supplierResult = await findOrCreateSupplier(connection, invoice.supplierName || po.supplierName || deliveryOrder.supplierName);
    const supplier = supplierResult.supplier;
    const dueDate = supplier.type === "EOM" ? eomPlusDays(invoiceDate, supplier.days) : addDays(invoiceDate, supplier.days);

    await connection.execute(
      `
        INSERT INTO purchase_orders
          (po_id, supplier_id, po_date, item_id, description, qty_ordered, unit_price, currency, total_amount)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON DUPLICATE KEY UPDATE
          supplier_id = VALUES(supplier_id),
          po_date = VALUES(po_date),
          item_id = VALUES(item_id),
          description = VALUES(description),
          qty_ordered = VALUES(qty_ordered),
          unit_price = VALUES(unit_price),
          currency = VALUES(currency),
          total_amount = VALUES(total_amount)
      `,
      [poId, supplier.supplier_id, poDate, itemId, po.itemDescription, money(po.quantityOrdered), money(po.unitPrice), supplier.currency, money(po.totalAmount)]
    );

    if (doId) {
      await connection.execute(
        `
          INSERT INTO delivery_orders
            (do_id, po_id, delivery_date, item_id, qty_delivered)
          VALUES (?, ?, ?, ?, ?)
          ON DUPLICATE KEY UPDATE
            po_id = VALUES(po_id),
            delivery_date = VALUES(delivery_date),
            item_id = VALUES(item_id),
            qty_delivered = VALUES(qty_delivered)
        `,
        [doId, poId, sqlDate(deliveryOrder.documentDate), itemId, money(deliveryOrder.quantityReceived)]
      );
    }

    await connection.execute(
      `
        INSERT INTO supplier_invoices
          (invoice_id, supplier_id, po_id, do_id, invoice_date, item_id, qty_invoiced, unit_price, tax_amount, total_amount, currency)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON DUPLICATE KEY UPDATE
          supplier_id = VALUES(supplier_id),
          po_id = VALUES(po_id),
          do_id = VALUES(do_id),
          invoice_date = VALUES(invoice_date),
          item_id = VALUES(item_id),
          qty_invoiced = VALUES(qty_invoiced),
          unit_price = VALUES(unit_price),
          tax_amount = VALUES(tax_amount),
          total_amount = VALUES(total_amount),
          currency = VALUES(currency)
      `,
      [invoiceId, supplier.supplier_id, poId, doId, invoiceDate, itemId, money(invoice.quantityBilled), money(invoice.unitPrice), 0, money(invoice.totalAmount), supplier.currency]
    );

    const draftRecord = decorateRecord({
      ...record,
      id: invoiceId,
      purchaseOrderFile: record.purchaseOrderFile || { originalName: `${poId}.sql`, storedName: poId },
      deliveryOrderFile: doId ? record.deliveryOrderFile || { originalName: `${doId}.sql`, storedName: doId } : null,
      invoiceFile: record.invoiceFile || { originalName: `${invoiceId}.sql`, storedName: invoiceId },
      databaseExceptions: [],
    });
    const exceptionFlag = draftRecord.discrepancies.length ? "Y" : "N";
    const paymentStatus = draftRecord.canApprovePayment ? dbPaymentStatuses.approved : dbPaymentStatuses.held;

    await connection.execute(
      `
        INSERT INTO payment_due_list
          (supplier_id, invoice_id, invoice_date, due_date, amount_due, currency, payment_status, exception_flag)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ON DUPLICATE KEY UPDATE
          supplier_id = VALUES(supplier_id),
          invoice_date = VALUES(invoice_date),
          due_date = VALUES(due_date),
          amount_due = VALUES(amount_due),
          currency = VALUES(currency),
          payment_status = VALUES(payment_status),
          exception_flag = VALUES(exception_flag)
      `,
      [supplier.supplier_id, invoiceId, invoiceDate, dueDate, money(invoice.totalAmount), supplier.currency, paymentStatus, exceptionFlag]
    );

    await connection.execute("DELETE FROM matching_exceptions WHERE invoice_id = ?", [invoiceId]);
    for (const discrepancy of draftRecord.discrepancies) {
      await connection.execute(
        `
          INSERT INTO matching_exceptions
            (invoice_id, exception_type, description)
          VALUES (?, ?, ?)
        `,
        [invoiceId, discrepancy.type.toUpperCase().replace(/[^A-Z0-9]+/g, "_").slice(0, 50), discrepancy.field]
      );
    }

    await connection.commit();
    return {
      recordId: invoiceId,
      supplierAutoCreated: supplierResult.created,
    };
  } catch (error) {
    await connection.rollback();
    throw error;
  } finally {
    connection.release();
  }
}

async function createUploadRecord(files) {
  const poFile = createUploadedFile(files.poFile?.[0]);
  const deliveryOrderFile = createUploadedFile(files.doGrnFile?.[0]);
  const invoiceFile = createUploadedFile(files.invoiceFile?.[0]);

  return saveDocumentSet({
    purchaseOrderFile: poFile,
    deliveryOrderFile,
    invoiceFile,
    purchaseOrder: extractPurchaseOrder(poFile),
    deliveryOrder: extractDeliveryOrder(deliveryOrderFile),
    invoice: extractInvoice(invoiceFile),
  });
}

async function saveCorrectedData(id, formData) {
  const existingRecord = await getRecord(id);
  return saveDocumentSet({
    ...(existingRecord || {}),
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
      invoiceNumber: formData.invoiceNumber || id,
      itemDescription: formData.invoiceItemDescription || "",
      quantityBilled: formData.invoiceQuantityBilled || "",
      unitPrice: formData.invoiceUnitPrice || "",
      totalAmount: formData.invoiceTotalAmount || "",
      documentDate: formData.invoiceDocumentDate || "",
    },
  });
}

async function setPaymentStatus(id, paymentStatus) {
  await ensurePaymentStatusColumnSupportsWorkflow();
  await db.execute(
    `
      UPDATE payment_due_list
      SET payment_status = ?
      WHERE invoice_id = ?
    `,
    [uiPaymentStatusToDb(paymentStatus), id]
  );
  return getRecord(id);
}

async function approvePayment(id) {
  return setPaymentStatus(id, paymentStatuses.approved);
}

async function rejectPayment(id) {
  return setPaymentStatus(id, paymentStatuses.rejected);
}

async function simulatePayment(id) {
  return setPaymentStatus(id, paymentStatuses.paid);
}

async function markPaymentPaid(id) {
  return setPaymentStatus(id, paymentStatuses.paid);
}

async function getPaymentList() {
  const records = await getRecords();
  return records.filter((record) =>
    record.canApprovePayment &&
    [paymentStatuses.pending, paymentStatuses.approved, paymentStatuses.held].includes(record.paymentStatus)
  );
}

async function getAllDiscrepancies() {
  const records = await getRecords();
  return records.flatMap((record) =>
    record.discrepancies.map((discrepancy) => ({
      recordId: record.id,
      ...discrepancy,
    }))
  );
}

async function getStats() {
  const records = await getRecords();
  const discrepancies = records.flatMap((record) => record.discrepancies);
  const paidRecords = records.filter((record) => record.paymentStatus === paymentStatuses.paid);

  return {
    totalRecords: records.length,
    totalUploadedDocumentSets: records.length,
    matched: records.filter((record) => record.matching.status === "Matched").length,
    mismatched: records.filter((record) => record.matching.status === "Mismatch").length,
    pendingReview: records.filter((record) => record.matching.status === "Pending Review").length,
    discrepancies: discrepancies.length,
    valid: records.filter((record) => record.validation.status === "Valid").length,
    invalid: records.filter((record) => record.validation.status === "Invalid").length,
    pendingApproval: records.filter((record) => [paymentStatuses.pending, paymentStatuses.held].includes(record.paymentStatus)).length,
    pendingPayments: records.filter((record) => [paymentStatuses.pending, paymentStatuses.held].includes(record.paymentStatus)).length,
    approvedPayments: records.filter((record) => record.paymentStatus === paymentStatuses.approved).length,
    rejectedPayments: records.filter((record) => record.paymentStatus === paymentStatuses.rejected).length,
    paidInvoices: paidRecords.length,
    totalAmountPaid: paidRecords.reduce((total, record) => total + money(record.payment?.amountPaid), 0),
  };
}

async function getReportRows() {
  const records = await getRecords();
  return records.map((record) => ({
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
    actionDate: record.payment?.paidAt || record.createdAt || "",
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
  markPaymentPaid,
  paymentStatuses,
  rejectPayment,
  saveCorrectedData,
  setPaymentStatus,
  simulatePayment,
  ensureDataFile,
};
