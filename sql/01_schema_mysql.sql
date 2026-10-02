-- ============================================================================
--  Customer Data Analysis  |  01 - MySQL schema
--  Project : customer-data-analysis-sql-excel
--  Author  : Gopal Singh Tomer
--
--  Creates a small star-ish schema (customers / products / orders / order_items)
--  that mirrors the CSV files in ../data/, then loads them with LOAD DATA.
--
--  Usage:
--      mysql -u root -p < sql/01_schema_mysql.sql
-- ============================================================================

DROP DATABASE IF EXISTS customer_analysis;
CREATE DATABASE customer_analysis CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
USE customer_analysis;

-- ---------------------------------------------------------------------------
-- Dimension: customers
-- ---------------------------------------------------------------------------
CREATE TABLE customers (
    customer_id   VARCHAR(12)  NOT NULL,
    first_name    VARCHAR(40)  NOT NULL,
    last_name     VARCHAR(40)  NOT NULL,
    email         VARCHAR(90)  DEFAULT NULL,
    phone         VARCHAR(20)  DEFAULT NULL,
    city          VARCHAR(40)  DEFAULT NULL,
    state         VARCHAR(40)  DEFAULT NULL,
    country       VARCHAR(40)  NOT NULL DEFAULT 'India',
    signup_date   DATE         NOT NULL,
    segment       VARCHAR(20)  NOT NULL,
    loyalty_tier  VARCHAR(20)  NOT NULL,
    PRIMARY KEY (customer_id),
    KEY idx_customers_state   (state),
    KEY idx_customers_segment (segment),
    KEY idx_customers_signup  (signup_date)
) ENGINE = InnoDB;

-- ---------------------------------------------------------------------------
-- Dimension: products
-- ---------------------------------------------------------------------------
CREATE TABLE products (
    product_id    VARCHAR(10)   NOT NULL,
    product_name  VARCHAR(60)   NOT NULL,
    category      VARCHAR(40)   NOT NULL,
    unit_price    DECIMAL(10,2) NOT NULL,
    PRIMARY KEY (product_id),
    KEY idx_products_category (category)
) ENGINE = InnoDB;

-- ---------------------------------------------------------------------------
-- Fact: orders
--   order_year / order_month are stored explicitly so the analysis queries stay
--   completely portable between MySQL and SQLite.
-- ---------------------------------------------------------------------------
CREATE TABLE orders (
    order_id        VARCHAR(12) NOT NULL,
    customer_id     VARCHAR(12) NOT NULL,
    order_date      DATE        NOT NULL,
    order_year      SMALLINT    NOT NULL,
    order_month     CHAR(7)     NOT NULL,      -- 'YYYY-MM'
    channel         VARCHAR(20) NOT NULL,
    payment_method  VARCHAR(25) NOT NULL,
    status          VARCHAR(15) NOT NULL,
    PRIMARY KEY (order_id),
    KEY idx_orders_customer (customer_id),
    KEY idx_orders_date     (order_date),
    KEY idx_orders_month    (order_month),
    CONSTRAINT fk_orders_customer FOREIGN KEY (customer_id)
        REFERENCES customers (customer_id)
) ENGINE = InnoDB;

-- ---------------------------------------------------------------------------
-- Fact: order_items
-- ---------------------------------------------------------------------------
CREATE TABLE order_items (
    order_item_id  VARCHAR(14)   NOT NULL,
    order_id       VARCHAR(12)   NOT NULL,
    product_id     VARCHAR(10)   NOT NULL,
    quantity       SMALLINT      NOT NULL,
    unit_price     DECIMAL(10,2) NOT NULL,
    discount_pct   DECIMAL(4,2)  NOT NULL DEFAULT 0,
    line_total     DECIMAL(12,2) NOT NULL,
    PRIMARY KEY (order_item_id),
    KEY idx_items_order   (order_id),
    KEY idx_items_product (product_id),
    CONSTRAINT fk_items_order   FOREIGN KEY (order_id)   REFERENCES orders (order_id),
    CONSTRAINT fk_items_product FOREIGN KEY (product_id) REFERENCES products (product_id)
) ENGINE = InnoDB;

-- ============================================================================
--  Loading the CSV extracts in ../data/
--  (run the python generator first:  python scripts/generate_data.py)
-- ============================================================================
-- LOAD DATA LOCAL INFILE '../data/customers.csv'
--   INTO TABLE customers
--   FIELDS TERMINATED BY ',' OPTIONALLY ENCLOSED BY '"'
--   LINES TERMINATED BY '\n'
--   IGNORE 1 LINES
--   (customer_id, first_name, last_name, email, phone, city, state, country,
--    signup_date, segment, loyalty_tier);

-- LOAD DATA LOCAL INFILE '../data/products.csv'
--   INTO TABLE products
--   FIELDS TERMINATED BY ',' OPTIONALLY ENCLOSED BY '"'
--   LINES TERMINATED BY '\n' IGNORE 1 LINES;

-- LOAD DATA LOCAL INFILE '../data/orders.csv'
--   INTO TABLE orders
--   FIELDS TERMINATED BY ',' OPTIONALLY ENCLOSED BY '"'
--   LINES TERMINATED BY '\n' IGNORE 1 LINES
--   (order_id, customer_id, order_date, order_year, order_month, channel,
--    payment_method, status);

-- LOAD DATA LOCAL INFILE '../data/order_items.csv'
--   INTO TABLE order_items
--   FIELDS TERMINATED BY ',' OPTIONALLY ENCLOSED BY '"'
--   LINES TERMINATED BY '\n' IGNORE 1 LINES
--   (order_item_id, order_id, product_id, quantity, unit_price, discount_pct,
--    line_total);

-- ---------------------------------------------------------------------------
-- Manual INSERT example (instead of the bulk loader)
-- ---------------------------------------------------------------------------
INSERT INTO customers
    (customer_id, first_name, last_name, email, phone, city, state,
     country, signup_date, segment, loyalty_tier)
VALUES
    ('CUST1001', 'Aarav',  'Sharma', 'aarav.sharma@example.com',
     '+91 9876543210', 'New Delhi', 'Delhi', 'India', '2022-03-14', 'Consumer', 'Silver'),
    ('CUST1002', 'Meera',  'Iyer',   'meera.iyer@example.com',
     '+91 9822113344', 'Bengaluru', 'Karnataka', 'India', '2022-07-02', 'Corporate', 'Gold');

-- ---------------------------------------------------------------------------
-- Sanity checks after loading
-- ---------------------------------------------------------------------------
SELECT 'customers'   AS table_name, COUNT(*) AS row_count FROM customers
UNION ALL SELECT 'products',    COUNT(*) FROM products
UNION ALL SELECT 'orders',      COUNT(*) FROM orders
UNION ALL SELECT 'order_items', COUNT(*) FROM order_items;

-- Referential integrity: no orphan order lines
SELECT COUNT(*) AS orphan_order_items
FROM order_items oi
LEFT JOIN orders o ON o.order_id = oi.order_id
WHERE o.order_id IS NULL;
