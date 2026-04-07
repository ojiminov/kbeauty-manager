-- ============================================================
-- KBeauty Manager — Supabase Database Schema
-- Run this entire file in: Supabase Dashboard → SQL Editor → New Query
-- ============================================================

-- 1. CUSTOMERS
CREATE TABLE customers (
  id           UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  name         TEXT NOT NULL,
  phone        TEXT,
  address      TEXT,
  created_at   TIMESTAMPTZ DEFAULT NOW()
);

-- 2. ORDERS
CREATE TABLE orders (
  id                    UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  customer_id           UUID REFERENCES customers(id) ON DELETE SET NULL,
  salesperson           TEXT,
  type                  TEXT DEFAULT 'B2C',          -- 'B2C' | 'B2B'
  status                TEXT DEFAULT 'New',          -- 'New' | 'Purchased' | 'Shipped' | 'Delivered'
  selling_price_usd     NUMERIC(10,2) DEFAULT 0,
  shipping_rate_per_kg  NUMERIC(10,4) DEFAULT 0,
  notes                 TEXT,
  created_at            TIMESTAMPTZ DEFAULT NOW()
);

-- 3. ORDER ITEMS
CREATE TABLE order_items (
  id                UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  order_id          UUID REFERENCES orders(id) ON DELETE CASCADE,
  product_name      TEXT NOT NULL,
  quantity          INTEGER DEFAULT 1,
  purchase_cost_krw NUMERIC(12,2) DEFAULT 0,
  weight_kg         NUMERIC(8,3) DEFAULT 0,
  created_at        TIMESTAMPTZ DEFAULT NOW()
);

-- 4. EXPENSES
CREATE TABLE expenses (
  id          UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  date        DATE NOT NULL,
  type        TEXT,     -- 'Salary' | 'Cargo' | 'Office' | 'Marketing' | 'Other'
  description TEXT,
  amount_usd  NUMERIC(10,2) DEFAULT 0,
  paid_by     TEXT,
  created_at  TIMESTAMPTZ DEFAULT NOW()
);

-- ============================================================
-- DISABLE Row Level Security (for internal team use only)
-- If you add login later, remove these lines and set up RLS policies
-- ============================================================
ALTER TABLE customers  DISABLE ROW LEVEL SECURITY;
ALTER TABLE orders     DISABLE ROW LEVEL SECURITY;
ALTER TABLE order_items DISABLE ROW LEVEL SECURITY;
ALTER TABLE expenses   DISABLE ROW LEVEL SECURITY;

-- ============================================================
-- Done! You should now see 4 tables in your Supabase Table Editor.
-- ============================================================
