-- ========================================================
-- Inventory Stock Tracker - MySQL 8.0 Database Schema
-- Database Name: inventory_tracker
-- ========================================================

CREATE DATABASE IF NOT EXISTS inventory_tracker
  CHARACTER SET utf8mb4
  COLLATE utf8mb4_unicode_ci;

USE inventory_tracker;

CREATE TABLE IF NOT EXISTS inventory (
    item_id VARCHAR(50) NOT NULL,
    name VARCHAR(150) NOT NULL,
    category VARCHAR(100) NOT NULL,
    spec VARCHAR(150) DEFAULT NULL,
    brand VARCHAR(100) NOT NULL,
    unit VARCHAR(50) NOT NULL DEFAULT 'Nos',
    qty INT NOT NULL DEFAULT 0,
    reorder INT NOT NULL DEFAULT 50,
    min_stock INT NOT NULL DEFAULT 10,
    max_stock INT NOT NULL DEFAULT 1000,
    cost DECIMAL(12, 2) NOT NULL DEFAULT 0.00,
    value DECIMAL(15, 2) NOT NULL DEFAULT 0.00,
    supplier VARCHAR(150) NOT NULL,
    warehouse VARCHAR(100) NOT NULL,
    bin VARCHAR(50) NOT NULL DEFAULT 'Not Assigned',
    status VARCHAR(50) NOT NULL DEFAULT 'In Stock',
    received DATE DEFAULT NULL,
    remarks VARCHAR(255) DEFAULT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    PRIMARY KEY (item_id),
    INDEX idx_category (category),
    INDEX idx_status (status),
    INDEX idx_brand (brand),
    INDEX idx_supplier (supplier),
    INDEX idx_warehouse (warehouse),
    INDEX idx_qty (qty)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

