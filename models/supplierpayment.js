const fs = require("fs");
const path = require("path");

const dataDirectory = path.join(__dirname, "..", "data");
const uploadsDirectory = path.join(__dirname, "..", "uploads");
const dataFile = path.join(dataDirectory, "supplier-payments.json");

const fieldsToCompare = [
  "supplierName",
  "itemName",
  "quantity",
  "unitPrice",
  "totalAmount",
  "date",
];

const demoDocumentSets = [
  {
    id: 1,
    createdAt: "2026-05-01",
    po: {
      supplierName: "Alpha Office Supplies",
      poNumber: "PO-1001",
      itemName: "A4 Paper Box",
      quantity: 20,
      unitPrice: 12.5,
      totalAmount: 250,
      date: "2026-04-22",
    },
    so: {
      supplierName: "Alpha Office Supplies",
      soNumber: "SO-9001",
      itemName: "A4 Paper Box",
      quantity: 20,
      unitPrice: 12.5,
      totalAmount: 250,
      date: "2026-04-22",
    },
    invoice: {
      supplierName: "Alpha Office Supplies",
      invoiceNumber: "INV-7001",
      itemName: "A4 Paper Box",
      quantity: 20,
      unitPrice: 12.5,
      totalAmount: 250,
      date: "2026-04-22",
    },
    uploadedFiles: {
      po: "alpha-po.pdf",
      so: "alpha-so.pdf",
      invoice: "alpha-invoice.pdf",
    },
  },
  {
    id: 2,
    createdAt: "2026-05-03",
    po: {
      supplierName: "Bright Tech Hardware",
      poNumber: "PO-1002",
      itemName: "Wireless Keyboard",
      quantity: 10,
      unitPrice: 45,
      totalAmount: 450,
      date: "2026-04-25",
    },
    so: {
      supplierName: "Bright Tech Hardware",
      soNumber: "SO-9002",
      itemName: "Wireless Keyboard",
      quantity: 10,
      unitPrice: 45,
      totalAmount: 450,
      date: "2026-04-25",
    },
    invoice: {
      supplierName: "Bright Tech Hardware",
      invoiceNumber: "INV-7002",
      itemName: "Wireless Keyboard",
      quantity: 12,
      unitPrice: 45,
      totalAmount: 540,
      date: "2026-04-25",
    },
    uploadedFiles: {
      po: "bright-po.pdf",
      so: "bright-so.pdf",
      invoice: "bright-invoice.pdf",
    },
  },
  {
    id: 3,
    createdAt: "2026-05-05",
    po: {
      supplierName: "City Cleaning Services",
      poNumber: "PO-1003",
      itemName: "Monthly Cleaning",
      quantity: 1,
      unitPrice: 800,
      totalAmount: 800,
      date: "2026-04-30",
    },
    so: null,
    invoice: {
      supplierName: "City Cleaning Services",
      invoiceNumber: "INV-7003",
      itemName: "Monthly Cleaning",
      quantity: 1,
      unitPrice: 800,
      totalAmount: 800,
      date: "2026-04-30",
    },
    uploadedFiles: {
      po: "city-po.pdf",
      so: "Pending upload",
      invoice: "city-invoice.pdf",
    },
  },
];

function ensureDataFile() {
  fs.mkdirSync(dataDirectory, { recursive: true });
  fs.mkdirSync(uploadsDirectory, { recursive: true });

  if (!fs.existsSync(dataFile)) {
    fs.writeFileSync(dataFile, JSON.stringify(demoDocumentSets, null, 2));
  }
}

function readDocumentSets() {
  ensureDataFile();
  return JSON.parse(fs.readFileSync(dataFile, "utf8"));
}

function saveDocumentSets(documentSets) {
  fs.writeFileSync(dataFile, JSON.stringify(documentSets, null, 2));
}

function normaliseValue(value) {
  if (typeof value === "number") {
    return Number(value.toFixed(2));
  }

  if (!Number.isNaN(Number(value)) && value !== "") {
    return Number(Number(value).toFixed(2));
  }

  return String(value || "").trim().toLowerCase();
}

function getStatus(documentSet) {
  if (!documentSet.po || !documentSet.so || !documentSet.invoice) {
    return "Pending review";
  }

  return getMismatchFields(documentSet).length === 0 ? "Matched" : "Mismatch";
}

function getMismatchFields(documentSet) {
  if (!documentSet.po || !documentSet.so || !documentSet.invoice) {
    return [];
  }

  return fieldsToCompare.filter((field) => {
    const poValue = normaliseValue(documentSet.po[field]);
    const soValue = normaliseValue(documentSet.so[field]);
    const invoiceValue = normaliseValue(documentSet.invoice[field]);

    return poValue !== soValue || poValue !== invoiceValue;
  });
}

function getDocumentSets() {
  return readDocumentSets()
    .map((documentSet) => ({
      ...documentSet,
      status: getStatus(documentSet),
      mismatchFields: getMismatchFields(documentSet),
    }))
    .sort((first, second) => second.id - first.id);
}

function addDocumentSet(formData) {
  const documentSets = readDocumentSets();
  const nextId = documentSets.length ? Math.max(...documentSets.map((set) => set.id)) + 1 : 1;

  // Real document extraction can replace this section later.
  const extractedData = {
    supplierName: formData.supplierName,
    itemName: formData.itemName,
    quantity: Number(formData.quantity),
    unitPrice: Number(formData.unitPrice),
    totalAmount: Number(formData.totalAmount),
    date: formData.date,
  };

  documentSets.push({
    id: nextId,
    createdAt: new Date().toISOString().slice(0, 10),
    po: {
      ...extractedData,
      poNumber: formData.poNumber,
    },
    so: {
      ...extractedData,
      soNumber: formData.soNumber,
    },
    invoice: {
      ...extractedData,
      invoiceNumber: formData.invoiceNumber,
    },
    uploadedFiles: formData.uploadedFiles,
  });

  saveDocumentSets(documentSets);
}

function getDashboardStats() {
  const documentSets = getDocumentSets();

  return {
    totalPurchaseOrders: documentSets.filter((set) => set.po).length,
    totalSalesOrders: documentSets.filter((set) => set.so).length,
    invoicesChecked: documentSets.filter((set) => set.invoice).length,
    matched: documentSets.filter((set) => set.status === "Matched").length,
    mismatched: documentSets.filter((set) => set.status === "Mismatch").length,
    pending: documentSets.filter((set) => set.status === "Pending review").length,
  };
}

function getReportRows() {
  return getDocumentSets().map((set) => ({
    id: set.id,
    supplierName: set.po?.supplierName || set.invoice?.supplierName || "Unknown supplier",
    poNumber: set.po?.poNumber || "Missing",
    soNumber: set.so?.soNumber || "Missing",
    invoiceNumber: set.invoice?.invoiceNumber || "Missing",
    totalAmount: set.invoice?.totalAmount || set.po?.totalAmount || 0,
    status: set.status,
    issue: set.mismatchFields.length ? set.mismatchFields.join(", ") : "No issue found",
  }));
}

module.exports = {
  addDocumentSet,
  ensureDataFile,
  getDashboardStats,
  getDocumentSets,
  getMismatchFields,
  getReportRows,
};
