CREATE SCHEMA IF NOT EXISTS `fypSQL` DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_0900_ai_ci ;
USE `fypSQL` ;

-- Drop tables first to avoid foreign key errors when re-running script
DROP TABLE IF EXISTS `paypal_webhook_events`;
DROP TABLE IF EXISTS `payment_cashflow_ledger`;
DROP TABLE IF EXISTS `payment_cash_reservations`;
DROP TABLE IF EXISTS `paypal_payouts`;
DROP TABLE IF EXISTS `payment_audit_log`;
DROP TABLE IF EXISTS `matching_exceptions`;
DROP TABLE IF EXISTS `payment_due_list`;
DROP TABLE IF EXISTS `supplier_invoices`;
DROP TABLE IF EXISTS `delivery_orders`;
DROP TABLE IF EXISTS `purchase_orders`;
DROP TABLE IF EXISTS `supplier_master`;
DROP TABLE IF EXISTS `payment_terms`;
DROP TABLE IF EXISTS `payment_system_settings`;

-- 1. Payment Terms Table
CREATE TABLE `payment_terms` (
  `term_code` VARCHAR(20) NOT NULL,
  `description` VARCHAR(100) NOT NULL,
  `days` INT NOT NULL,
  `type` ENUM('FIXED', 'EOM') NOT NULL,

  PRIMARY KEY (`term_code`)
) ENGINE = InnoDB;

-- 2. Supplier Master Table
CREATE TABLE `supplier_master` (
  `supplier_id` VARCHAR(20) NOT NULL,
  `supplier_name` VARCHAR(150) NOT NULL,
  `currency` VARCHAR(10) NOT NULL,
  `payment_term_code` VARCHAR(20) NOT NULL,
  `credit_limit` DECIMAL(12,2) NULL,
  `tax_id` VARCHAR(50) NULL,
  `bank_account` VARCHAR(50) NULL,
  `bank_name` VARCHAR(100) NULL,
  `paypal_email` VARCHAR(254) NULL,
  `paypal_recipient_verified` BOOLEAN NOT NULL DEFAULT FALSE,
  `active_flag` ENUM('Y', 'N') NOT NULL DEFAULT 'Y',

  PRIMARY KEY (`supplier_id`),

  CONSTRAINT `fk_supplier_payment_term`
    FOREIGN KEY (`payment_term_code`)
    REFERENCES `payment_terms` (`term_code`)
    ON UPDATE CASCADE
    ON DELETE RESTRICT
) ENGINE = InnoDB;

-- 3. Purchase Orders Table
CREATE TABLE `purchase_orders` (
  `po_id` VARCHAR(20) NOT NULL,
  `supplier_id` VARCHAR(20) NOT NULL,
  `po_date` DATE NOT NULL,
  `item_id` VARCHAR(20) NOT NULL,
  `description` VARCHAR(255) NOT NULL,
  `qty_ordered` DECIMAL(10,2) NOT NULL,
  `unit_price` DECIMAL(12,2) NOT NULL,
  `currency` VARCHAR(10) NOT NULL,
  `total_amount` DECIMAL(12,2) NOT NULL,

  PRIMARY KEY (`po_id`),

  CONSTRAINT `fk_po_supplier`
    FOREIGN KEY (`supplier_id`)
    REFERENCES `supplier_master` (`supplier_id`)
    ON UPDATE CASCADE
    ON DELETE RESTRICT
) ENGINE = InnoDB;

-- 4. Delivery Orders Table
CREATE TABLE `delivery_orders` (
  `do_id` VARCHAR(20) NOT NULL,
  `po_id` VARCHAR(20) NOT NULL,
  `delivery_date` DATE NOT NULL,
  `item_id` VARCHAR(20) NOT NULL,
  `qty_delivered` DECIMAL(10,2) NOT NULL,

  PRIMARY KEY (`do_id`),

  CONSTRAINT `fk_delivery_po`
    FOREIGN KEY (`po_id`)
    REFERENCES `purchase_orders` (`po_id`)
    ON UPDATE CASCADE
    ON DELETE RESTRICT
) ENGINE = InnoDB;

-- 5. Supplier Invoices Table
CREATE TABLE `supplier_invoices` (
  `invoice_id` VARCHAR(20) NOT NULL,
  `supplier_id` VARCHAR(20) NOT NULL,
  `po_id` VARCHAR(20) NOT NULL,
  `do_id` VARCHAR(20) NULL,
  `invoice_date` DATE NULL,
  `item_id` VARCHAR(20) NOT NULL,
  `qty_invoiced` DECIMAL(10,2) NOT NULL,
  `unit_price` DECIMAL(12,2) NOT NULL,
  `tax_amount` DECIMAL(12,2) NOT NULL DEFAULT 0.00,
  `total_amount` DECIMAL(12,2) NOT NULL,
  `currency` VARCHAR(10) NOT NULL,
  `payment_term_code` VARCHAR(20) NULL,
  `extracted_payment_term` VARCHAR(100) NULL,
  `uploaded_at` TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,

  PRIMARY KEY (`invoice_id`),

  CONSTRAINT `fk_invoice_supplier`
    FOREIGN KEY (`supplier_id`)
    REFERENCES `supplier_master` (`supplier_id`)
    ON UPDATE CASCADE
    ON DELETE RESTRICT,

  CONSTRAINT `fk_invoice_payment_term`
    FOREIGN KEY (`payment_term_code`)
    REFERENCES `payment_terms` (`term_code`)
    ON UPDATE CASCADE
    ON DELETE RESTRICT,

  CONSTRAINT `fk_invoice_po`
    FOREIGN KEY (`po_id`)
    REFERENCES `purchase_orders` (`po_id`)
    ON UPDATE CASCADE
    ON DELETE RESTRICT,

  CONSTRAINT `fk_invoice_delivery`
    FOREIGN KEY (`do_id`)
    REFERENCES `delivery_orders` (`do_id`)
    ON UPDATE CASCADE
    ON DELETE SET NULL
) ENGINE = InnoDB;

-- 6. Payment Due List Table
CREATE TABLE `payment_due_list` (
  `payment_due_id` INT NOT NULL AUTO_INCREMENT,
  `supplier_id` VARCHAR(20) NOT NULL,
  `invoice_id` VARCHAR(20) NOT NULL,
  `invoice_date` DATE NULL,
  `due_date` DATE NULL,
  `amount_due` DECIMAL(12,2) NOT NULL,
  `currency` VARCHAR(10) NOT NULL,
  `payment_status` ENUM(
    'READY',
    'HOLD',
    'PAID',
    'PENDING_APPROVAL',
    'APPROVED_FOR_PAYMENT',
    'PAYMENT_PROCESSING',
    'REJECTED',
    'PAYMENT_HELD'
  ) NOT NULL DEFAULT 'PAYMENT_HELD',
  `exception_flag` ENUM('Y', 'N') NOT NULL DEFAULT 'N',

  PRIMARY KEY (`payment_due_id`),

  UNIQUE KEY `unique_invoice_payment_due` (`invoice_id`),

  CONSTRAINT `fk_payment_due_supplier`
    FOREIGN KEY (`supplier_id`)
    REFERENCES `supplier_master` (`supplier_id`)
    ON UPDATE CASCADE
    ON DELETE RESTRICT,

  CONSTRAINT `fk_payment_due_invoice`
    FOREIGN KEY (`invoice_id`)
    REFERENCES `supplier_invoices` (`invoice_id`)
    ON UPDATE CASCADE
    ON DELETE CASCADE
) ENGINE = InnoDB;

-- 7. Matching Exceptions Table
CREATE TABLE `matching_exceptions` (
  `exception_id` INT NOT NULL AUTO_INCREMENT,
  `invoice_id` VARCHAR(20) NOT NULL,
  `exception_type` VARCHAR(50) NOT NULL,
  `description` VARCHAR(255) NOT NULL,
  `created_at` DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,

  PRIMARY KEY (`exception_id`),

  CONSTRAINT `fk_exception_invoice`
    FOREIGN KEY (`invoice_id`)
    REFERENCES `supplier_invoices` (`invoice_id`)
    ON UPDATE CASCADE
    ON DELETE CASCADE
) ENGINE = InnoDB;

-- 8. Payment workflow audit log
CREATE TABLE `payment_audit_log` (
  `audit_id` BIGINT NOT NULL AUTO_INCREMENT,
  `invoice_id` VARCHAR(20) NULL,
  `payment_due_id` INT NULL,
  `action_type` VARCHAR(80) NOT NULL,
  `action_status` VARCHAR(30) NOT NULL,
  `review_reason` VARCHAR(100) NULL,
  `details` VARCHAR(500) NULL,
  `actor_type` VARCHAR(30) NOT NULL DEFAULT 'SYSTEM',
  `created_at` DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (`audit_id`),
  KEY `idx_audit_invoice` (`invoice_id`),
  KEY `idx_audit_payment_due` (`payment_due_id`),
  CONSTRAINT `fk_audit_invoice`
    FOREIGN KEY (`invoice_id`) REFERENCES `supplier_invoices` (`invoice_id`)
    ON UPDATE CASCADE ON DELETE SET NULL,
  CONSTRAINT `fk_audit_payment_due`
    FOREIGN KEY (`payment_due_id`) REFERENCES `payment_due_list` (`payment_due_id`)
    ON UPDATE CASCADE ON DELETE SET NULL
) ENGINE = InnoDB;

-- 9. PayPal payout audit trail
CREATE TABLE `paypal_payouts` (
  `id` BIGINT NOT NULL AUTO_INCREMENT,
  `payment_due_id` INT NOT NULL,
  `supplier_id` VARCHAR(20) NOT NULL,
  `supplier_name` VARCHAR(150) NOT NULL,
  `sender_batch_id` VARCHAR(100) NOT NULL,
  `sender_item_id` VARCHAR(100) NOT NULL,
  `paypal_batch_id` VARCHAR(50) NULL,
  `paypal_item_id` VARCHAR(50) NULL,
  `recipient_email` VARCHAR(254) NOT NULL,
  `amount` DECIMAL(12,2) NOT NULL,
  `currency` VARCHAR(10) NOT NULL,
  `status` VARCHAR(40) NOT NULL,
  `error_code` VARCHAR(100) NULL,
  `error_message` VARCHAR(500) NULL,
  `raw_response` JSON NULL,
  `created_at` DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  `updated_at` DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`id`),
  UNIQUE KEY `uq_paypal_sender_batch` (`sender_batch_id`),
  UNIQUE KEY `uq_paypal_sender_item` (`sender_item_id`),
  UNIQUE KEY `uq_paypal_batch` (`paypal_batch_id`),
  UNIQUE KEY `uq_paypal_item` (`paypal_item_id`),
  CONSTRAINT `fk_paypal_payment_due` FOREIGN KEY (`payment_due_id`)
    REFERENCES `payment_due_list` (`payment_due_id`) ON UPDATE CASCADE ON DELETE RESTRICT
) ENGINE = InnoDB;

CREATE TABLE `paypal_webhook_events` (
  `event_id` VARCHAR(100) NOT NULL,
  `event_type` VARCHAR(100) NOT NULL,
  `received_at` DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (`event_id`)
) ENGINE = InnoDB;

-- Simulated cash balance used only by the PayPal Sandbox workflow
CREATE TABLE `payment_system_settings` (
  `setting_key` VARCHAR(80) NOT NULL,
  `setting_value` VARCHAR(255) NOT NULL DEFAULT '',
  `decimal_value` DECIMAL(14,2) NOT NULL,
  `currency` VARCHAR(10) NOT NULL DEFAULT 'SGD',
  `updated_at` DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`setting_key`)
) ENGINE = InnoDB;

INSERT INTO `payment_system_settings` (`setting_key`, `setting_value`, `decimal_value`, `currency`)
VALUES ('simulated_available_cash', '100000.00', 100000.00, 'SGD');

CREATE TABLE `payment_cash_reservations` (
  `reservation_id` BIGINT NOT NULL AUTO_INCREMENT,
  `payment_due_id` INT NOT NULL,
  `paypal_payout_id` BIGINT NOT NULL,
  `amount` DECIMAL(12,2) NOT NULL,
  `currency` VARCHAR(10) NOT NULL,
  `status` ENUM('RESERVED','SETTLED','RELEASED') NOT NULL DEFAULT 'RESERVED',
  `created_at` DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  `updated_at` DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`reservation_id`),
  KEY `idx_cash_reservation_payment` (`payment_due_id`),
  UNIQUE KEY `uq_cash_reservation_payout` (`paypal_payout_id`),
  CONSTRAINT `fk_cash_reservation_payment` FOREIGN KEY (`payment_due_id`)
    REFERENCES `payment_due_list` (`payment_due_id`) ON UPDATE CASCADE ON DELETE RESTRICT,
  CONSTRAINT `fk_cash_reservation_payout` FOREIGN KEY (`paypal_payout_id`)
    REFERENCES `paypal_payouts` (`id`) ON UPDATE CASCADE ON DELETE RESTRICT
) ENGINE = InnoDB;

CREATE TABLE `payment_cashflow_ledger` (
  `ledger_id` BIGINT NOT NULL AUTO_INCREMENT,
  `reservation_id` BIGINT NOT NULL,
  `payment_due_id` INT NOT NULL,
  `paypal_payout_id` BIGINT NOT NULL,
  `action_type` ENUM('RESERVE','SETTLE','RELEASE') NOT NULL,
  `amount` DECIMAL(12,2) NOT NULL,
  `currency` VARCHAR(10) NOT NULL,
  `available_cash_after` DECIMAL(14,2) NOT NULL,
  `created_at` DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (`ledger_id`),
  KEY `idx_cashflow_payment` (`payment_due_id`),
  KEY `idx_cashflow_payout` (`paypal_payout_id`),
  CONSTRAINT `fk_cashflow_reservation` FOREIGN KEY (`reservation_id`)
    REFERENCES `payment_cash_reservations` (`reservation_id`) ON UPDATE CASCADE ON DELETE RESTRICT
) ENGINE = InnoDB;

INSERT INTO payment_terms
(term_code, description, days, type)
VALUES
('NET30', 'Pay in 30 days', 30, 'FIXED'),
('NET60', 'Pay in 60 days', 60, 'FIXED'),
('DUEONRECEIPT', 'Due on receipt', 0, 'FIXED'),
('EOM30', 'End of month plus 30 days', 30, 'EOM');

INSERT INTO supplier_master
(supplier_id, supplier_name, currency, payment_term_code, credit_limit, tax_id, bank_account, bank_name, paypal_email, paypal_recipient_verified, active_flag)
VALUES
('SUP001', 'Alpha Industrial Pte Ltd', 'SGD', 'NET30', 100000.00, '201912345Z', 'DBS-123456', 'DBS Bank', 'supplieracc@business.example.com', TRUE, 'Y'),
('SUP002', 'Beta Office Supplies', 'SGD', 'NET60', 50000.00, '202012345A', 'OCBC-888999', 'OCBC Bank', 'supplieracc@business.example.com', TRUE, 'Y'),
('SUP003', 'Gamma Tech Parts', 'USD', 'EOM30', 75000.00, '202112345B', 'UOB-555666', 'UOB Bank', 'supplieracc@business.example.com', TRUE, 'N');

INSERT INTO purchase_orders
(po_id, supplier_id, po_date, item_id, description, qty_ordered, unit_price, currency, total_amount)
VALUES
('PO1001', 'SUP001', '2026-01-10', 'ITEM001', 'Steel Rod', 100.00, 10.00, 'SGD', 1000.00),
('PO1002', 'SUP002', '2026-01-12', 'ITEM002', 'Printer Paper', 200.00, 5.00, 'SGD', 1000.00),
('PO1003', 'SUP003', '2026-01-15', 'ITEM003', 'Circuit Board', 50.00, 20.00, 'USD', 1000.00);

INSERT INTO delivery_orders
(do_id, po_id, delivery_date, item_id, qty_delivered)
VALUES
('DO5001', 'PO1001', '2026-01-15', 'ITEM001', 100.00),
('DO5002', 'PO1002', '2026-01-18', 'ITEM002', 180.00),
('DO5003', 'PO1003', '2026-01-20', 'ITEM003', 50.00);

INSERT INTO supplier_invoices
(invoice_id, supplier_id, po_id, do_id, invoice_date, item_id, qty_invoiced, unit_price,
 tax_amount, total_amount, currency, payment_term_code, extracted_payment_term)
VALUES
('INV9001', 'SUP001', 'PO1001', 'DO5001', '2026-01-20', 'ITEM001', 100.00, 10.00, 90.00, 1090.00, 'SGD', 'NET30', 'NET30'),
('INV9002', 'SUP002', 'PO1002', 'DO5002', '2026-01-22', 'ITEM002', 200.00, 5.00, 90.00, 1090.00, 'SGD', 'NET60', 'NET60'),
('INV9003', 'SUP003', 'PO1003', 'DO5003', '2026-01-25', 'ITEM003', 50.00, 25.00, 0.00, 1250.00, 'USD', 'EOM30', 'EOM30');

INSERT INTO payment_due_list
(supplier_id, invoice_id, invoice_date, due_date, amount_due, currency, payment_status, exception_flag)
VALUES
('SUP001', 'INV9001', '2026-01-20', '2026-02-19', 1090.00, 'SGD', 'READY', 'N'),
('SUP002', 'INV9002', '2026-01-22', '2026-03-23', 1090.00, 'SGD', 'HOLD', 'Y'),
('SUP003', 'INV9003', '2026-01-25', '2026-03-02', 1250.00, 'USD', 'HOLD', 'Y');

INSERT INTO matching_exceptions
(invoice_id, exception_type, description)
VALUES
('INV9002', 'QTY_MISMATCH', 'Invoice quantity is more than delivered quantity'),
('INV9003', 'PRICE_MISMATCH', 'Invoice unit price is higher than PO unit price'),
('INV9003', 'INACTIVE_SUPPLIER', 'Supplier is inactive');
