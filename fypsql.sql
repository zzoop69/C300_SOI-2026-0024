CREATE SCHEMA IF NOT EXISTS `fypSQL` DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_0900_ai_ci ;
USE `fypSQL` ;

-- Drop tables first to avoid foreign key errors when re-running script
DROP TABLE IF EXISTS `matching_exceptions`;
DROP TABLE IF EXISTS `payment_due_list`;
DROP TABLE IF EXISTS `supplier_invoices`;
DROP TABLE IF EXISTS `delivery_orders`;
DROP TABLE IF EXISTS `purchase_orders`;
DROP TABLE IF EXISTS `supplier_master`;
DROP TABLE IF EXISTS `payment_terms`;

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
  `invoice_date` DATE NOT NULL,
  `item_id` VARCHAR(20) NOT NULL,
  `qty_invoiced` DECIMAL(10,2) NOT NULL,
  `unit_price` DECIMAL(12,2) NOT NULL,
  `tax_amount` DECIMAL(12,2) NOT NULL DEFAULT 0.00,
  `total_amount` DECIMAL(12,2) NOT NULL,
  `currency` VARCHAR(10) NOT NULL,

  PRIMARY KEY (`invoice_id`),

  CONSTRAINT `fk_invoice_supplier`
    FOREIGN KEY (`supplier_id`)
    REFERENCES `supplier_master` (`supplier_id`)
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
  `invoice_date` DATE NOT NULL,
  `due_date` DATE NOT NULL,
  `amount_due` DECIMAL(12,2) NOT NULL,
  `currency` VARCHAR(10) NOT NULL,
  `payment_status` ENUM('READY', 'HOLD', 'PAID') NOT NULL DEFAULT 'HOLD',
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

INSERT INTO payment_terms
(term_code, description, days, type)
VALUES
('NET30', 'Pay in 30 days', 30, 'FIXED'),
('NET60', 'Pay in 60 days', 60, 'FIXED'),
('EOM30', 'End of month plus 30 days', 30, 'EOM');

INSERT INTO supplier_master
(supplier_id, supplier_name, currency, payment_term_code, credit_limit, tax_id, bank_account, bank_name, active_flag)
VALUES
('SUP001', 'Alpha Industrial Pte Ltd', 'SGD', 'NET30', 100000.00, '201912345Z', 'DBS-123456', 'DBS Bank', 'Y'),
('SUP002', 'Beta Office Supplies', 'SGD', 'NET60', 50000.00, '202012345A', 'OCBC-888999', 'OCBC Bank', 'Y'),
('SUP003', 'Gamma Tech Parts', 'USD', 'EOM30', 75000.00, '202112345B', 'UOB-555666', 'UOB Bank', 'N');

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
(invoice_id, supplier_id, po_id, do_id, invoice_date, item_id, qty_invoiced, unit_price, tax_amount, total_amount, currency)
VALUES
('INV9001', 'SUP001', 'PO1001', 'DO5001', '2026-01-20', 'ITEM001', 100.00, 10.00, 90.00, 1090.00, 'SGD'),
('INV9002', 'SUP002', 'PO1002', 'DO5002', '2026-01-22', 'ITEM002', 200.00, 5.00, 90.00, 1090.00, 'SGD'),
('INV9003', 'SUP003', 'PO1003', 'DO5003', '2026-01-25', 'ITEM003', 50.00, 25.00, 0.00, 1250.00, 'USD');

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