-- ============================================================
-- MUNA MANAGER — Supabase Schema v3.0
-- Run this in: Supabase → SQL Editor → New Query → Run
-- ============================================================

-- ─── PRODUCT CATALOG ─────────────────────────────────────────
CREATE TABLE IF NOT EXISTS catalog (
  id                      UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  brand                   TEXT NOT NULL,
  name                    TEXT NOT NULL,
  category                TEXT DEFAULT '',
  skin_type               TEXT DEFAULT 'All',
  weight_grams            NUMERIC DEFAULT 0,
  estimated_cost_krw      NUMERIC DEFAULT 0,
  estimated_shipping_krw  NUMERIC DEFAULT 0,
  suggested_b2c_price_uzs NUMERIC DEFAULT 0,
  suggested_b2b_price_uzs NUMERIC DEFAULT 0,
  notes                   TEXT DEFAULT '',
  created_at              TIMESTAMPTZ DEFAULT NOW()
);

-- ─── ORDERS ──────────────────────────────────────────────────
-- Status flow: Pending → Purchased → Shipped → Delivered → Paid
CREATE TABLE IF NOT EXISTS orders (
  id            UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  customer_name TEXT NOT NULL,
  product_id    UUID REFERENCES catalog(id) ON DELETE SET NULL,
  quantity      INTEGER NOT NULL DEFAULT 1,
  price_uzs     NUMERIC NOT NULL DEFAULT 0,
  type          TEXT DEFAULT 'B2C',
  status        TEXT DEFAULT 'Pending',
  salesperson   TEXT DEFAULT '',
  notes         TEXT DEFAULT '',
  created_at    TIMESTAMPTZ DEFAULT NOW()
);

-- ─── INVENTORY ───────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS inventory (
  id          UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  product_id  UUID REFERENCES catalog(id) ON DELETE CASCADE UNIQUE,
  quantity    INTEGER DEFAULT 0,
  updated_at  TIMESTAMPTZ DEFAULT NOW()
);

-- ─── EXPENSES ────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS expenses (
  id          UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  date        DATE NOT NULL DEFAULT CURRENT_DATE,
  type        TEXT NOT NULL,
  description TEXT NOT NULL,
  amount_krw  NUMERIC DEFAULT 0,
  paid_by     TEXT DEFAULT '',
  created_at  TIMESTAMPTZ DEFAULT NOW()
);

-- ─── INDEXES ─────────────────────────────────────────────────
CREATE INDEX IF NOT EXISTS idx_orders_status     ON orders(status);
CREATE INDEX IF NOT EXISTS idx_orders_created_at ON orders(created_at DESC);
CREATE INDEX IF NOT EXISTS idx_catalog_brand     ON catalog(brand);

-- ─── ROW LEVEL SECURITY ──────────────────────────────────────
ALTER TABLE catalog   ENABLE ROW LEVEL SECURITY;
ALTER TABLE orders    ENABLE ROW LEVEL SECURITY;
ALTER TABLE inventory ENABLE ROW LEVEL SECURITY;
ALTER TABLE expenses  ENABLE ROW LEVEL SECURITY;

DO $$ BEGIN
  IF NOT EXISTS (SELECT 1 FROM pg_policies WHERE tablename='catalog'   AND policyname='allow_all_catalog')   THEN CREATE POLICY allow_all_catalog   ON catalog   FOR ALL USING (true) WITH CHECK (true); END IF;
  IF NOT EXISTS (SELECT 1 FROM pg_policies WHERE tablename='orders'    AND policyname='allow_all_orders')    THEN CREATE POLICY allow_all_orders    ON orders    FOR ALL USING (true) WITH CHECK (true); END IF;
  IF NOT EXISTS (SELECT 1 FROM pg_policies WHERE tablename='inventory' AND policyname='allow_all_inventory') THEN CREATE POLICY allow_all_inventory ON inventory FOR ALL USING (true) WITH CHECK (true); END IF;
  IF NOT EXISTS (SELECT 1 FROM pg_policies WHERE tablename='expenses'  AND policyname='allow_all_expenses')  THEN CREATE POLICY allow_all_expenses  ON expenses  FOR ALL USING (true) WITH CHECK (true); END IF;
END $$;

-- ─── SAMPLE CATALOG DATA ─────────────────────────────────────
INSERT INTO catalog (brand, name, category, skin_type, weight_grams, estimated_cost_krw, estimated_shipping_krw, suggested_b2c_price_uzs, suggested_b2b_price_uzs, notes) VALUES
  ('COSRX',      'Advanced Snail 96 Mucin Power Essence', 'Serum',       'All',      100, 18000, 3000, 280000, 220000, 'Best seller'),
  ('Some By Mi', 'AHA BHA PHA 30 Days Miracle Toner',     'Toner',       'Oily',     150, 15000, 3500, 220000, 170000, ''),
  ('Laneige',    'Lip Sleeping Mask Berry',                'Lip',         'All',       20, 25000, 1500, 350000, 280000, 'Top seller'),
  ('COSRX',      'Low pH Good Morning Gel Cleanser',      'Cleanser',    'Oily',     150, 12000, 3000, 180000, 140000, ''),
  ('Innisfree',  'Green Tea Seed Serum',                   'Serum',       'Dry',       80, 22000, 2500, 300000, 240000, ''),
  ('Klairs',     'Supple Preparation Unscented Toner',     'Toner',       'Sensitive',180, 19000, 3500, 250000, 195000, ''),
  ('Missha',     'Time Revolution Night Repair Serum',     'Serum',       'All',       50, 35000, 2000, 450000, 360000, 'Premium')
ON CONFLICT DO NOTHING;
