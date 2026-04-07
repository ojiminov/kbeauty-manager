# 🚀 KBeauty Manager — Deployment Guide
> Follow these steps in order. Takes about 30–45 minutes total.

---

## STEP 1 — Set Up Supabase Database

1. Go to **supabase.com** → open your project
2. In the left sidebar click **SQL Editor**
3. Click **New Query**
4. Open the file `supabase-schema.sql` from your folder
5. Copy everything and paste it into the SQL Editor
6. Click **Run** (green button)
7. You should see: *"Success. No rows returned"*
8. Click **Table Editor** in the sidebar — you should now see 4 tables:
   - customers
   - orders
   - order_items
   - expenses

---

## STEP 2 — Get Your Supabase Keys

1. In Supabase, go to **Settings** (gear icon) → **API**
2. Copy these two values (you'll need them in Step 3):
   - **Project URL** — looks like: `https://abcdefgh.supabase.co`
   - **anon / public key** — starts with `eyJ...` (long string)

---

## STEP 3 — Add Your Keys to the App

1. Open `kbeauty-manager.html` in any text editor (Notepad is fine)
2. Find these two lines near the top (around line 30):
   ```
   const SUPABASE_URL  = 'YOUR_SUPABASE_URL';
   const SUPABASE_KEY  = 'YOUR_SUPABASE_ANON_KEY';
   ```
3. Replace the placeholder text with your actual values:
   ```
   const SUPABASE_URL  = 'https://abcdefgh.supabase.co';
   const SUPABASE_KEY  = 'eyJhbGciOiJIUzI1NiIsInR5cCI6...';
   ```
4. Save the file

✅ Test it locally: Open `kbeauty-manager.html` in Chrome. If it loads without error, you're ready.

---

## STEP 4 — Upload to GitHub

1. Go to **github.com** and log in
2. Click the **+** button (top right) → **New repository**
3. Name it: `kbeauty-manager`
4. Set it to **Private** (recommended for a business app)
5. Click **Create repository**
6. On the next page, click **uploading an existing file**
7. Drag and drop `kbeauty-manager.html` into the upload area
8. At the bottom, click **Commit changes**

---

## STEP 5 — Deploy to Vercel (Free Hosting)

1. Go to **vercel.com** → Sign up with your GitHub account
2. Click **Add New Project**
3. Find `kbeauty-manager` in the list → click **Import**
4. Leave all settings as default → click **Deploy**
5. Wait ~1 minute
6. Vercel gives you a URL like: `https://kbeauty-manager.vercel.app`

🎉 **Your app is now live!** Share that URL with your 5 team members.

---

## STEP 6 — How to Update the App Later

Whenever you want to make changes:

1. I update the `kbeauty-manager.html` file for you
2. You go to your GitHub repository
3. Click on `kbeauty-manager.html` → click the **pencil icon** (Edit)
4. Select all text → paste the new version
5. Click **Commit changes**
6. Vercel automatically redeploys in ~1 minute
7. Everyone's app updates instantly ✓

---

## Your App URLs

| What | Where |
|------|-------|
| Live app | `https://kbeauty-manager.vercel.app` (after deployment) |
| GitHub code | `https://github.com/YOUR_USERNAME/kbeauty-manager` |
| Supabase data | `https://supabase.com/dashboard` |

---

## Troubleshooting

**"Could not connect to database"** → Your Supabase URL or Key is wrong. Re-check Step 3.

**App loads but data doesn't save** → Go to Supabase → Table Editor and confirm the 4 tables exist. If not, re-run the SQL schema.

**Blank white screen** → Open browser DevTools (F12) → Console tab → Share the red error message.

---

*Need help? Just tell me what error you see and I'll fix it.*
