const crypto = require("crypto");
const db = require("../db");

let schemaReady = false;
const DEFAULT_SANDBOX_RECIPIENT = "supplieracc@business.example.com";

async function columnExists(tableName, columnName) {
  const [rows] = await db.execute(
    `SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
     WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = ? AND COLUMN_NAME = ? LIMIT 1`,
    [tableName, columnName]
  );
  return rows.length > 0;
}

async function columnIsNullable(tableName, columnName) {
  const [rows] = await db.execute(
    `SELECT IS_NULLABLE FROM INFORMATION_SCHEMA.COLUMNS
     WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = ? AND COLUMN_NAME = ? LIMIT 1`,
    [tableName, columnName]
  );
  return rows[0]?.IS_NULLABLE === "YES";
}

async function indexExists(tableName, indexName) {
  const [rows] = await db.execute(
    `SELECT 1 FROM INFORMATION_SCHEMA.STATISTICS
     WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = ? AND INDEX_NAME = ? LIMIT 1`,
    [tableName, indexName]
  );
  return rows.length > 0;
}

function sharedSandboxRecipient() {
  if ((process.env.PAYPAL_MODE || "sandbox").toLowerCase() !== "sandbox") {
    throw new Error("The shared supplier recipient is restricted to PAYPAL_MODE=sandbox.");
  }
  return cleanEmail(
    process.env.PAYPAL_SANDBOX_RECIPIENT_EMAIL ||
    process.env.SUPPLIER_PAYPAL_EMAIL ||
    DEFAULT_SANDBOX_RECIPIENT
  );
}

async function ensurePayPalSchema() {
  if (schemaReady) return;
  if (!(await columnExists("supplier_master", "paypal_email"))) {
    await db.execute("ALTER TABLE supplier_master ADD COLUMN paypal_email VARCHAR(254) NULL");
  }
  if (!(await columnExists("supplier_master", "paypal_recipient_verified"))) {
    await db.execute("ALTER TABLE supplier_master ADD COLUMN paypal_recipient_verified BOOLEAN NOT NULL DEFAULT FALSE");
  }
  await db.execute(`
    CREATE TABLE IF NOT EXISTS paypal_payouts (
      id BIGINT NOT NULL AUTO_INCREMENT,
      payment_due_id INT NOT NULL,
      supplier_id VARCHAR(20) NOT NULL,
      supplier_name VARCHAR(150) NOT NULL,
      sender_batch_id VARCHAR(100) NOT NULL,
      sender_item_id VARCHAR(100) NOT NULL,
      paypal_batch_id VARCHAR(50) NULL,
      paypal_item_id VARCHAR(50) NULL,
      recipient_email VARCHAR(254) NOT NULL,
      amount DECIMAL(12,2) NOT NULL,
      currency VARCHAR(10) NOT NULL,
      status VARCHAR(40) NOT NULL,
      error_code VARCHAR(100) NULL,
      error_message VARCHAR(500) NULL,
      raw_response JSON NULL,
      created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
      updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
      PRIMARY KEY (id),
      UNIQUE KEY uq_paypal_sender_batch (sender_batch_id),
      UNIQUE KEY uq_paypal_sender_item (sender_item_id),
      UNIQUE KEY uq_paypal_batch (paypal_batch_id),
      UNIQUE KEY uq_paypal_item (paypal_item_id),
      KEY idx_paypal_payment_due (payment_due_id),
      CONSTRAINT fk_paypal_payment_due FOREIGN KEY (payment_due_id)
        REFERENCES payment_due_list(payment_due_id) ON UPDATE CASCADE ON DELETE RESTRICT
    ) ENGINE=InnoDB
  `);
  await db.execute(`
    CREATE TABLE IF NOT EXISTS paypal_webhook_events (
      event_id VARCHAR(100) NOT NULL,
      event_type VARCHAR(100) NOT NULL,
      received_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
      PRIMARY KEY (event_id)
    ) ENGINE=InnoDB
  `);
  await db.execute(`
    CREATE TABLE IF NOT EXISTS payment_system_settings (
      setting_key VARCHAR(80) NOT NULL,
      setting_value VARCHAR(255) NOT NULL DEFAULT '',
      decimal_value DECIMAL(14,2) NOT NULL,
      currency VARCHAR(10) NOT NULL DEFAULT 'SGD',
      updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
      PRIMARY KEY (setting_key)
    ) ENGINE=InnoDB
  `);
  if (!(await columnExists("payment_system_settings", "decimal_value"))) {
    await db.execute(
      "ALTER TABLE payment_system_settings ADD COLUMN decimal_value DECIMAL(14,2) NULL"
    );
  }
  if (!(await columnExists("payment_system_settings", "currency"))) {
    await db.execute(
      "ALTER TABLE payment_system_settings ADD COLUMN currency VARCHAR(10) NULL DEFAULT 'SGD'"
    );
  }
  await db.execute(`
    INSERT INTO payment_system_settings (setting_key, setting_value, decimal_value, currency)
    VALUES ('simulated_available_cash', '100000.00', 100000.00, 'SGD')
    ON DUPLICATE KEY UPDATE
      decimal_value = COALESCE(decimal_value, CAST(setting_value AS DECIMAL(14,2))),
      currency = COALESCE(currency, 'SGD')
  `);
  await db.execute(`
    CREATE TABLE IF NOT EXISTS payment_audit_log (
      audit_id BIGINT NOT NULL AUTO_INCREMENT,
      invoice_id VARCHAR(20) NULL,
      payment_due_id INT NULL,
      action_type VARCHAR(80) NOT NULL,
      action_status VARCHAR(30) NOT NULL,
      review_reason VARCHAR(100) NULL,
      details VARCHAR(500) NULL,
      actor_type VARCHAR(30) NOT NULL DEFAULT 'SYSTEM',
      created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
      PRIMARY KEY (audit_id),
      KEY idx_audit_invoice (invoice_id),
      KEY idx_audit_payment_due (payment_due_id),
      CONSTRAINT fk_audit_invoice FOREIGN KEY (invoice_id)
        REFERENCES supplier_invoices(invoice_id) ON UPDATE CASCADE ON DELETE SET NULL,
      CONSTRAINT fk_audit_payment_due FOREIGN KEY (payment_due_id)
        REFERENCES payment_due_list(payment_due_id) ON UPDATE CASCADE ON DELETE SET NULL
    ) ENGINE=InnoDB
  `);
  await db.execute(`
    CREATE TABLE IF NOT EXISTS payment_cash_reservations (
      reservation_id BIGINT NOT NULL AUTO_INCREMENT,
      payment_due_id INT NOT NULL,
      paypal_payout_id BIGINT NOT NULL,
      amount DECIMAL(12,2) NOT NULL,
      currency VARCHAR(10) NOT NULL,
      status ENUM('RESERVED','SETTLED','RELEASED') NOT NULL DEFAULT 'RESERVED',
      created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
      updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
      PRIMARY KEY (reservation_id),
      KEY idx_cash_reservation_payment (payment_due_id),
      UNIQUE KEY uq_cash_reservation_payout (paypal_payout_id),
      CONSTRAINT fk_cash_reservation_payment FOREIGN KEY (payment_due_id)
        REFERENCES payment_due_list(payment_due_id) ON UPDATE CASCADE ON DELETE RESTRICT,
      CONSTRAINT fk_cash_reservation_payout FOREIGN KEY (paypal_payout_id)
        REFERENCES paypal_payouts(id) ON UPDATE CASCADE ON DELETE RESTRICT
    ) ENGINE=InnoDB
  `);
  if (await indexExists("payment_cash_reservations", "uq_cash_reservation_payment")) {
    await db.execute(
      "ALTER TABLE payment_cash_reservations DROP INDEX uq_cash_reservation_payment"
    );
  }
  await db.execute(`
    CREATE TABLE IF NOT EXISTS payment_cashflow_ledger (
      ledger_id BIGINT NOT NULL AUTO_INCREMENT,
      reservation_id BIGINT NOT NULL,
      payment_due_id INT NOT NULL,
      paypal_payout_id BIGINT NOT NULL,
      action_type ENUM('RESERVE','SETTLE','RELEASE') NOT NULL,
      amount DECIMAL(12,2) NOT NULL,
      currency VARCHAR(10) NOT NULL,
      available_cash_after DECIMAL(14,2) NOT NULL,
      created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
      PRIMARY KEY (ledger_id),
      KEY idx_cashflow_payment (payment_due_id),
      KEY idx_cashflow_payout (paypal_payout_id),
      CONSTRAINT fk_cashflow_reservation FOREIGN KEY (reservation_id)
        REFERENCES payment_cash_reservations(reservation_id) ON UPDATE CASCADE ON DELETE RESTRICT
    ) ENGINE=InnoDB
  `);
  if (!(await columnExists("paypal_payouts", "supplier_id"))) {
    await db.execute("ALTER TABLE paypal_payouts ADD COLUMN supplier_id VARCHAR(20) NULL AFTER payment_due_id");
  }
  if (!(await columnExists("paypal_payouts", "supplier_name"))) {
    await db.execute("ALTER TABLE paypal_payouts ADD COLUMN supplier_name VARCHAR(150) NULL AFTER supplier_id");
  }
  await db.execute(`
    UPDATE paypal_payouts pp
    INNER JOIN payment_due_list pdl ON pdl.payment_due_id = pp.payment_due_id
    INNER JOIN supplier_master sm ON sm.supplier_id = pdl.supplier_id
    SET pp.supplier_id = COALESCE(pp.supplier_id, pdl.supplier_id),
        pp.supplier_name = COALESCE(pp.supplier_name, sm.supplier_name)
    WHERE pp.supplier_id IS NULL OR pp.supplier_name IS NULL
  `);
  if (await columnIsNullable("paypal_payouts", "supplier_id")) {
    await db.execute("ALTER TABLE paypal_payouts MODIFY supplier_id VARCHAR(20) NOT NULL");
  }
  if (await columnIsNullable("paypal_payouts", "supplier_name")) {
    await db.execute("ALTER TABLE paypal_payouts MODIFY supplier_name VARCHAR(150) NOT NULL");
  }
  if (await indexExists("paypal_payouts", "idx_paypal_batch")) {
    await db.execute("ALTER TABLE paypal_payouts DROP INDEX idx_paypal_batch");
  }
  if (!(await indexExists("paypal_payouts", "uq_paypal_batch"))) {
    await db.execute("ALTER TABLE paypal_payouts ADD UNIQUE KEY uq_paypal_batch (paypal_batch_id)");
  }
  if (!(await indexExists("paypal_payouts", "uq_paypal_item"))) {
    await db.execute("ALTER TABLE paypal_payouts ADD UNIQUE KEY uq_paypal_item (paypal_item_id)");
  }
  const recipientEmail = sharedSandboxRecipient();
  await db.execute(
    `UPDATE supplier_master SET paypal_email = ?, paypal_recipient_verified = TRUE
     WHERE paypal_email IS NULL OR paypal_email <> ? OR paypal_recipient_verified = FALSE`,
    [recipientEmail, recipientEmail]
  );
  schemaReady = true;
}

function cleanEmail(value) {
  const email = String(value || "").trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || email.length > 254) {
    throw new Error("Enter a valid sandbox supplier PayPal email address.");
  }
  return email;
}

async function preparePayout(invoiceId) {
  await ensurePayPalSchema();
  const connection = await db.getConnection();
  try {
    await connection.beginTransaction();
    const [rows] = await connection.execute(
      `SELECT pdl.payment_due_id, pdl.payment_status, pdl.exception_flag,
              pdl.amount_due, pdl.currency, pdl.supplier_id, sm.supplier_name,
              sm.paypal_email, sm.paypal_recipient_verified, sm.active_flag,
              si.supplier_id AS invoice_supplier_id, si.po_id AS invoice_po_id,
              si.do_id AS invoice_do_id, si.item_id AS invoice_item_id,
              si.qty_invoiced, si.unit_price AS invoice_unit_price,
              si.total_amount AS invoice_total_amount, si.currency AS invoice_currency,
              po.supplier_id AS po_supplier_id, po.item_id AS po_item_id,
              po.qty_ordered, po.unit_price AS po_unit_price,
              po.total_amount AS po_total_amount, po.currency AS po_currency,
              dox.po_id AS delivery_po_id, dox.item_id AS delivery_item_id,
              dox.qty_delivered,
              (SELECT COUNT(*) FROM matching_exceptions me
               WHERE me.invoice_id = si.invoice_id) AS exception_count
       FROM payment_due_list pdl
       INNER JOIN supplier_master sm ON sm.supplier_id = pdl.supplier_id
       INNER JOIN supplier_invoices si ON si.invoice_id = pdl.invoice_id
       INNER JOIN purchase_orders po ON po.po_id = si.po_id
       LEFT JOIN delivery_orders dox ON dox.do_id = si.do_id
       WHERE pdl.invoice_id = ? FOR UPDATE`,
      [invoiceId]
    );
    const payment = rows[0];
    if (!payment) throw new Error("Payment record was not found.");
    if (payment.supplier_id !== payment.invoice_supplier_id) {
      throw new Error("Payment and invoice supplier IDs do not match.");
    }
    const sameMoney = (left, right) => Math.abs(Number(left) - Number(right)) <= 0.01;
    const documentsMatch =
      payment.invoice_do_id &&
      payment.supplier_id === payment.po_supplier_id &&
      payment.invoice_po_id === payment.delivery_po_id &&
      payment.invoice_item_id === payment.po_item_id &&
      payment.invoice_item_id === payment.delivery_item_id &&
      sameMoney(payment.qty_invoiced, payment.qty_ordered) &&
      sameMoney(payment.qty_invoiced, payment.qty_delivered) &&
      sameMoney(payment.invoice_unit_price, payment.po_unit_price) &&
      sameMoney(payment.invoice_total_amount, payment.po_total_amount) &&
      sameMoney(payment.amount_due, payment.invoice_total_amount) &&
      payment.currency === payment.invoice_currency &&
      payment.currency === payment.po_currency &&
      Number(payment.exception_count) === 0;
    const documentsValid =
      Number(payment.qty_invoiced) > 0 &&
      Number(payment.qty_ordered) > 0 &&
      Number(payment.qty_delivered) > 0 &&
      Number(payment.invoice_unit_price) > 0 &&
      Number(payment.po_unit_price) > 0 &&
      Number(payment.invoice_total_amount) > 0;
    if (!documentsMatch || !documentsValid) {
      throw new Error(
        "Payment eligibility changed. The invoice must remain Matched, Valid, and discrepancy-free."
      );
    }
    if (!["PENDING_APPROVAL", "PAYMENT_HELD", "APPROVED_FOR_PAYMENT", "READY", "HOLD"].includes(payment.payment_status)) {
      throw new Error("This payment has already been rejected, submitted, or completed.");
    }
    if (payment.exception_flag !== "N" || payment.active_flag !== "Y") {
      throw new Error("This supplier payment is held and cannot be sent.");
    }
    if (!(Number(payment.amount_due) > 0) || !/^[A-Z]{3}$/.test(payment.currency)) {
      throw new Error("The payment amount or currency is invalid.");
    }

    const recipientEmail = sharedSandboxRecipient();
    if (payment.paypal_email !== recipientEmail || !payment.paypal_recipient_verified) {
      throw new Error("The supplier is not configured for the shared PayPal sandbox recipient.");
    }

    const [existingRows] = await connection.execute(
      `SELECT id, status FROM paypal_payouts
       WHERE payment_due_id = ? AND status NOT IN ('FAILED','DENIED','BLOCKED','RETURNED','REFUNDED','CANCELED','UNCLAIMED')
       ORDER BY id DESC LIMIT 1 FOR UPDATE`,
      [payment.payment_due_id]
    );
    if (existingRows[0]) {
      throw new Error("This invoice already has a payout in progress or completed.");
    }

    const [cashRows] = await connection.execute(
      `SELECT decimal_value, currency FROM payment_system_settings
       WHERE setting_key = 'simulated_available_cash' FOR UPDATE`
    );
    const cash = cashRows[0];
    if (!cash || cash.currency !== payment.currency) {
      throw new Error(`Simulated cashflow is not configured for ${payment.currency}.`);
    }
    if (Number(cash.decimal_value) < Number(payment.amount_due)) {
      throw new Error(
        `Insufficient simulated cash. Available: ${cash.currency} ${Number(cash.decimal_value).toFixed(2)}.`
      );
    }

    const unique = crypto.randomUUID();
    const senderBatchId = `invoice-${invoiceId}-${unique}`.slice(0, 100);
    const senderItemId = `item-${invoiceId}-${unique}`.slice(0, 100);
    const [result] = await connection.execute(
      `INSERT INTO paypal_payouts
       (payment_due_id, supplier_id, supplier_name, sender_batch_id, sender_item_id,
        recipient_email, amount, currency, status)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'INITIATING')`,
      [payment.payment_due_id, payment.supplier_id, payment.supplier_name, senderBatchId,
        senderItemId, recipientEmail, payment.amount_due, payment.currency]
    );
    const availableCashAfter = Number(cash.decimal_value) - Number(payment.amount_due);
    await connection.execute(
      `UPDATE payment_system_settings SET decimal_value = ?, setting_value = ?
       WHERE setting_key = 'simulated_available_cash'`,
      [availableCashAfter, availableCashAfter.toFixed(2)]
    );
    const [reservationResult] = await connection.execute(
      `INSERT INTO payment_cash_reservations
       (payment_due_id, paypal_payout_id, amount, currency, status)
       VALUES (?, ?, ?, ?, 'RESERVED')`,
      [payment.payment_due_id, result.insertId, payment.amount_due, payment.currency]
    );
    await connection.execute(
      `INSERT INTO payment_cashflow_ledger
       (reservation_id, payment_due_id, paypal_payout_id, action_type, amount, currency, available_cash_after)
       VALUES (?, ?, ?, 'RESERVE', ?, ?, ?)`,
      [reservationResult.insertId, payment.payment_due_id, result.insertId,
        payment.amount_due, payment.currency, availableCashAfter]
    );
    await connection.execute(
      "UPDATE payment_due_list SET payment_status = 'PAYMENT_PROCESSING' WHERE payment_due_id = ?",
      [payment.payment_due_id]
    );
    await connection.execute(
      `INSERT INTO payment_audit_log
       (invoice_id, payment_due_id, action_type, action_status, details, actor_type)
       VALUES (?, ?, 'APPROVE_AND_RESERVE', 'SUCCESS', ?, 'MANAGER')`,
      [invoiceId, payment.payment_due_id,
        `Reserved ${payment.currency} ${Number(payment.amount_due).toFixed(2)} and started PayPal Sandbox submission.`]
    );
    await connection.commit();
    return { ...payment, payoutId: result.insertId, recipientEmail, senderBatchId, senderItemId, invoiceId };
  } catch (error) {
    await connection.rollback();
    throw error;
  } finally {
    connection.release();
  }
}

async function recordSubmission(payoutId, response) {
  const batch = response.batch_header || {};
  const connection = await db.getConnection();
  try {
    await connection.beginTransaction();
    const [rows] = await connection.execute(
      `SELECT pp.payment_due_id, pdl.invoice_id FROM paypal_payouts pp
       INNER JOIN payment_due_list pdl ON pdl.payment_due_id = pp.payment_due_id
       WHERE pp.id = ? FOR UPDATE`,
      [payoutId]
    );
    await connection.execute(
      `UPDATE paypal_payouts SET paypal_batch_id = ?, status = ?, raw_response = ? WHERE id = ?`,
      [batch.payout_batch_id || null, batch.batch_status || "PENDING", JSON.stringify(response), payoutId]
    );
    if (rows[0]) {
      await connection.execute(
        `INSERT INTO payment_audit_log
         (invoice_id, payment_due_id, action_type, action_status, details)
         VALUES (?, ?, 'PAYPAL_SUBMISSION', 'SUCCESS', ?)`,
        [rows[0].invoice_id, rows[0].payment_due_id,
          `PayPal Sandbox accepted payout batch ${batch.payout_batch_id || "(pending ID)"}.`]
      );
    }
    await connection.commit();
  } catch (error) {
    await connection.rollback();
    throw error;
  } finally {
    connection.release();
  }
}

async function recordFailure(payoutId, error) {
  const connection = await db.getConnection();
  try {
    await connection.beginTransaction();
    const [rows] = await connection.execute(
      `SELECT pp.payment_due_id, pdl.invoice_id FROM paypal_payouts pp
       INNER JOIN payment_due_list pdl ON pdl.payment_due_id = pp.payment_due_id
       WHERE pp.id = ? FOR UPDATE`, [payoutId]
    );
    const definitiveFailure = Number(error.status) >= 400 && Number(error.status) < 500;
    await connection.execute(
      `UPDATE paypal_payouts SET status = ?, error_code = ?, error_message = ? WHERE id = ?`,
      [definitiveFailure ? "FAILED" : "UNKNOWN", error.details?.issue || error.details?.name || error.name, String(error.message).slice(0, 500), payoutId]
    );
    if (rows[0] && definitiveFailure) {
      const [reservationRows] = await connection.execute(
        `SELECT * FROM payment_cash_reservations
         WHERE paypal_payout_id = ? AND status = 'RESERVED' FOR UPDATE`,
        [payoutId]
      );
      const reservation = reservationRows[0];
      if (reservation) {
        const [cashRows] = await connection.execute(
          `SELECT decimal_value FROM payment_system_settings
           WHERE setting_key = 'simulated_available_cash' FOR UPDATE`
        );
        const availableCashAfter = Number(cashRows[0].decimal_value) + Number(reservation.amount);
        await connection.execute(
          `UPDATE payment_system_settings SET decimal_value = ?, setting_value = ?
           WHERE setting_key = 'simulated_available_cash'`,
          [availableCashAfter, availableCashAfter.toFixed(2)]
        );
        await connection.execute(
          "UPDATE payment_cash_reservations SET status = 'RELEASED' WHERE reservation_id = ?",
          [reservation.reservation_id]
        );
        await connection.execute(
          `INSERT INTO payment_cashflow_ledger
           (reservation_id, payment_due_id, paypal_payout_id, action_type, amount, currency, available_cash_after)
           VALUES (?, ?, ?, 'RELEASE', ?, ?, ?)`,
          [reservation.reservation_id, reservation.payment_due_id, payoutId,
            reservation.amount, reservation.currency, availableCashAfter]
        );
      }
      await connection.execute(
        `UPDATE payment_due_list SET payment_status = 'APPROVED_FOR_PAYMENT'
         WHERE payment_due_id = ? AND payment_status = 'PAYMENT_PROCESSING'`,
        [rows[0].payment_due_id]
      );
    }
    if (rows[0]) {
      await connection.execute(
        `INSERT INTO payment_audit_log
         (invoice_id, payment_due_id, action_type, action_status, details)
         VALUES (?, ?, 'PAYPAL_SUBMISSION', ?, ?)`,
        [rows[0].invoice_id, rows[0].payment_due_id,
          definitiveFailure ? "FAILED" : "UNCERTAIN",
          String(error.message || "PayPal submission error").slice(0, 500)]
      );
    }
    await connection.commit();
  } catch (dbError) {
    await connection.rollback();
    throw dbError;
  } finally {
    connection.release();
  }
}

async function getPayoutByInvoice(invoiceId) {
  await ensurePayPalSchema();
  const [rows] = await db.execute(
    `SELECT pp.* FROM paypal_payouts pp
     INNER JOIN payment_due_list pdl ON pdl.payment_due_id = pp.payment_due_id
     WHERE pdl.invoice_id = ? ORDER BY pp.id DESC LIMIT 1`,
    [invoiceId]
  );
  return rows[0] || null;
}

function itemStatusFromBatch(batch) {
  return batch.items?.[0]?.transaction_status || batch.batch_header?.batch_status || "PENDING";
}

async function applyPayPalStatusByBatch(batchId, status, response = null, itemId = null) {
  await ensurePayPalSchema();
  const connection = await db.getConnection();
  try {
    await connection.beginTransaction();
    const [rows] = await connection.execute(
      `SELECT pp.id, pp.payment_due_id, pdl.invoice_id
       FROM paypal_payouts pp
       INNER JOIN payment_due_list pdl ON pdl.payment_due_id = pp.payment_due_id
       WHERE pp.paypal_batch_id = ? FOR UPDATE`, [batchId]
    );
    if (!rows[0]) {
      await connection.rollback();
      return false;
    }
    const normalized = String(status || "PENDING").toUpperCase();
    await connection.execute(
      `UPDATE paypal_payouts SET status = ?, paypal_item_id = COALESCE(?, paypal_item_id),
       raw_response = COALESCE(?, raw_response) WHERE id = ?`,
      [normalized, itemId, response ? JSON.stringify(response) : null, rows[0].id]
    );
    const succeeded = ["SUCCESS", "SUCCEEDED"].includes(normalized);
    const failed = ["FAILED", "DENIED", "BLOCKED", "RETURNED", "REFUNDED", "CANCELED", "UNCLAIMED"].includes(normalized);
    if (succeeded) {
      await connection.execute("UPDATE payment_due_list SET payment_status = 'PAID' WHERE payment_due_id = ?", [rows[0].payment_due_id]);
    } else if (failed) {
      await connection.execute("UPDATE payment_due_list SET payment_status = 'PAYMENT_HELD' WHERE payment_due_id = ?", [rows[0].payment_due_id]);
    }
    if (succeeded || failed) {
      const [reservationRows] = await connection.execute(
        `SELECT * FROM payment_cash_reservations
         WHERE paypal_payout_id = ? AND status = 'RESERVED' FOR UPDATE`,
        [rows[0].id]
      );
      const reservation = reservationRows[0];
      if (reservation) {
        const [cashRows] = await connection.execute(
          `SELECT decimal_value FROM payment_system_settings
           WHERE setting_key = 'simulated_available_cash' FOR UPDATE`
        );
        let availableCashAfter = Number(cashRows[0].decimal_value);
        const action = succeeded ? "SETTLE" : "RELEASE";
        if (failed) {
          availableCashAfter += Number(reservation.amount);
          await connection.execute(
            `UPDATE payment_system_settings SET decimal_value = ?, setting_value = ?
             WHERE setting_key = 'simulated_available_cash'`,
            [availableCashAfter, availableCashAfter.toFixed(2)]
          );
        }
        await connection.execute(
          "UPDATE payment_cash_reservations SET status = ? WHERE reservation_id = ?",
          [succeeded ? "SETTLED" : "RELEASED", reservation.reservation_id]
        );
        await connection.execute(
          `INSERT INTO payment_cashflow_ledger
           (reservation_id, payment_due_id, paypal_payout_id, action_type, amount, currency, available_cash_after)
           VALUES (?, ?, ?, ?, ?, ?, ?)`,
          [reservation.reservation_id, reservation.payment_due_id, rows[0].id, action,
            reservation.amount, reservation.currency, availableCashAfter]
        );
      }
      await connection.execute(
        `INSERT INTO payment_audit_log
         (invoice_id, payment_due_id, action_type, action_status, details)
         VALUES (?, ?, 'PAYPAL_RECONCILIATION', ?, ?)`,
        [rows[0].invoice_id, rows[0].payment_due_id,
          succeeded ? "COMPLETED" : "FAILED", `PayPal payout item status: ${normalized}.`]
      );
    }
    await connection.commit();
    return true;
  } catch (error) {
    await connection.rollback();
    throw error;
  } finally {
    connection.release();
  }
}

async function recordWebhookOnce(event) {
  await ensurePayPalSchema();
  try {
    await db.execute("INSERT INTO paypal_webhook_events (event_id, event_type) VALUES (?, ?)", [event.id, event.event_type]);
    return true;
  } catch (error) {
    if (error.code === "ER_DUP_ENTRY") return false;
    throw error;
  }
}

async function releaseWebhookEvent(eventId) {
  await db.execute("DELETE FROM paypal_webhook_events WHERE event_id = ?", [eventId]);
}

async function getCashflowSummary() {
  await ensurePayPalSchema();
  const [[cashRows], [reservedRows]] = await Promise.all([
    db.execute(
      `SELECT decimal_value, currency FROM payment_system_settings
       WHERE setting_key = 'simulated_available_cash'`
    ),
    db.execute(
      `SELECT COALESCE(SUM(amount), 0) AS reserved_cash
       FROM payment_cash_reservations WHERE status = 'RESERVED'`
    ),
  ]);
  return {
    availableCash: Number(cashRows[0]?.decimal_value || 0),
    reservedCash: Number(reservedRows[0]?.reserved_cash || 0),
    currency: cashRows[0]?.currency || "SGD",
  };
}

module.exports = {
  applyPayPalStatusByBatch,
  ensurePayPalSchema,
  getCashflowSummary,
  getPayoutByInvoice,
  itemStatusFromBatch,
  preparePayout,
  recordFailure,
  releaseWebhookEvent,
  recordSubmission,
  recordWebhookOnce,
  sharedSandboxRecipient,
};
