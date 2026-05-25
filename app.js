const express = require("express");
const path = require("path");

const supplierPaymentRoutes = require("./controllers/supplierpaymentcontroller");
const { ensureDataFile } = require("./models/supplierpayment");

const app = express();
const PORT = process.env.PORT || 3000;

// Create the JSON data file and upload folder when the project first runs.
ensureDataFile();

app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));

app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

app.use("/", supplierPaymentRoutes);

app.listen(PORT, () => {
  console.log(`Autonomous Supplier Payment Process is running at http://localhost:${PORT}`);
});
n