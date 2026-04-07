/**
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║   K-BEAUTY UZ — Full Accounting System  v2  (Google Apps Script)        ║
 * ║                                                                          ║
 * ║   HOW TO USE:                                                            ║
 * ║   1. Open a new blank Google Sheet                                       ║
 * ║   2. Click  Extensions → Apps Script                                     ║
 * ║   3. Delete all existing code and paste this entire file                 ║
 * ║   4. Save (💾), then Run → buildAll                                      ║
 * ║   5. Approve permissions when prompted                                   ║
 * ║   6. Wait ~45 sec → spreadsheet is ready!                                ║
 * ║   7. A "K-Beauty UZ ▾" menu appears at the top — use it to add data     ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

// ─────────────────────────────────────────────────────────────────────────────
// COLOURS
// ─────────────────────────────────────────────────────────────────────────────
const C = {
  navy:'#1B3A6B', gold:'#C9A84C', lightBlue:'#D6E4F0', lightGold:'#FFF5DC',
  yellow:'#FFFACD', white:'#FFFFFF', green:'#1E8449', red:'#C0392B',
  purple:'#7D3C98', inputBlue:'#EBF5FB', formulaGray:'#F2F3F4',
  rowAlt:'#F8FBFF',
};
const FMT_USD='$#,##0.00;($#,##0.00)', FMT_USD0='$#,##0;($#,##0)',
      FMT_KRW='₩#,##0', FMT_UZS='#,##0" UZS"', FMT_PCT='0.0%',
      FMT_NUM='#,##0', FMT_DATE='dd/mm/yyyy';

// Row where data starts in each sheet (update if you change structure)
const ROW = { prod:6, inv:8, sales:8, exp:8 };


// ═════════════════════════════════════════════════════════════════════════════
// ①  onOpen  — adds the custom K-Beauty UZ menu every time the sheet opens
// ═════════════════════════════════════════════════════════════════════════════
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('K-Beauty UZ ▾')
    .addItem('➕  Add New Product',        'showAddProductForm')
    .addItem('🛒  Log a Sale / Order',     'showAddSaleForm')
    .addItem('💸  Record an Expense',      'showAddExpenseForm')
    .addSeparator()
    .addItem('🔍  Search Records',         'showSearchDialog')
    .addSeparator()
    .addItem('🔄  Rebuild Entire Sheet',   'buildAll')
    .addToUi();
}


// ═════════════════════════════════════════════════════════════════════════════
// ②  SERVER-SIDE DATA WRITERS  (called from the HTML dialogs)
// ═════════════════════════════════════════════════════════════════════════════

/** Appends a new product to 🛍️ Products and returns success / error string */
function saveProduct(d) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('🛍️ Products');
    if (!sh) return 'ERROR: Products sheet not found. Run buildAll() first.';

    // Find next empty row (column A)
    const lastRow = sh.getLastRow();
    const r = lastRow + 1;

    sh.getRange(r,1).setValue(d.id).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial');
    sh.getRange(r,2).setValue(d.brand).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial');
    sh.getRange(r,3).setValue(d.name).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial');
    sh.getRange(r,4).setValue(d.variant).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial');
    sh.getRange(r,5).setValue(d.category).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial');
    sh.getRange(r,6).setValue(d.skinType).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial');
    sh.getRange(r,7).setValue(Number(d.weight)).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial').setNumberFormat(FMT_NUM);
    sh.getRange(r,8).setValue(Number(d.costKrw)).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial').setNumberFormat(FMT_KRW);
    sh.getRange(r,9).setValue(Number(d.sellUzs)).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial').setNumberFormat(FMT_UZS);
    sh.getRange(r,10).setFormula(`=IF(A${r}="","",H${r}*KRW_UZS_RATE+(G${r}/1000)*SHIPPING_KRW_KG*KRW_UZS_RATE)`).setNumberFormat(FMT_UZS).setBackground(C.formulaGray).setFontFamily('Arial');
    sh.getRange(r,11).setFormula(`=IF(A${r}="","",I${r}-J${r})`).setNumberFormat(FMT_UZS).setBackground(C.formulaGray).setFontFamily('Arial');
    sh.getRange(r,1,1,11).setBorder(true,true,true,true,true,true,'#CCCCCC',SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(r,22);
    SpreadsheetApp.flush();
    return 'OK: Product "' + d.name + '" added at row ' + r;
  } catch(e) { return 'ERROR: ' + e.message; }
}

/** Appends a new sale row to 💰 Sales */
function saveSale(d) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('💰 Sales');
    if (!sh) return 'ERROR: Sales sheet not found.';

    // Find next empty row after header
    let r = ROW.sales;
    while (sh.getRange(r,1).getValue() !== '') r++;

    sh.getRange(r,1).setValue(d.orderId).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial');
    sh.getRange(r,2).setValue(new Date(d.date)).setNumberFormat(FMT_DATE).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial');
    sh.getRange(r,3).setValue(d.customer).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial');
    sh.getRange(r,4).setValue(d.contact).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial');
    sh.getRange(r,5).setValue(d.productId).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial');
    sh.getRange(r,6).setFormula(`=IF(E${r}="","",IFERROR(VLOOKUP(E${r},'🛍️ Products'!A:C,2,0),""))`).setBackground(C.formulaGray).setFontFamily('Arial').setFontSize(10);
    sh.getRange(r,7).setValue(Number(d.qty)).setNumberFormat(FMT_NUM).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial');
    sh.getRange(r,8).setFormula(`=IF(E${r}="","",IFERROR(G${r}*VLOOKUP(E${r},'🛍️ Products'!A:I,9,0),0))`).setNumberFormat(FMT_UZS).setBackground(C.formulaGray).setFontFamily('Arial');
    sh.getRange(r,9).setValue(d.status).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial');
    sh.getRange(r,10).setValue(d.shipping).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial');
    sh.getRange(r,11).setValue(d.tracking).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial');
    sh.getRange(r,12).setValue(d.region).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial');
    sh.getRange(r,13).setValue(d.notes).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial');
    sh.getRange(r,1,1,13).setBorder(true,true,true,true,true,true,'#CCCCCC',SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(r,22);
    SpreadsheetApp.flush();
    return 'OK: Order "' + d.orderId + '" logged at row ' + r;
  } catch(e) { return 'ERROR: ' + e.message; }
}

/** Appends a new expense row to 💸 Expenses */
function saveExpense(d) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('💸 Expenses');
    if (!sh) return 'ERROR: Expenses sheet not found.';

    let r = ROW.exp;
    while (sh.getRange(r,1).getValue() !== '') r++;

    sh.getRange(r,1).setValue(d.expId).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial');
    sh.getRange(r,2).setValue(new Date(d.date)).setNumberFormat(FMT_DATE).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial');
    sh.getRange(r,3).setValue(d.category).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial');
    sh.getRange(r,4).setValue(d.description).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial');
    sh.getRange(r,5).setValue(d.vendor).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial');
    sh.getRange(r,6).setValue(Number(d.amount)).setNumberFormat(FMT_USD).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial');
    sh.getRange(r,7).setValue(d.payMethod).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial');
    sh.getRange(r,8).setValue(d.receipt).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial');
    sh.getRange(r,9).setValue(d.paidBy).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial');
    sh.getRange(r,10).setFormula(`=IF(B${r}="","",TEXT(B${r},"YYYY-MM"))`).setBackground(C.formulaGray).setFontFamily('Arial').setFontSize(10);
    sh.getRange(r,11).setValue(d.notes).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial');
    sh.getRange(r,1,1,11).setBorder(true,true,true,true,true,true,'#CCCCCC',SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(r,22);
    SpreadsheetApp.flush();
    return 'OK: Expense "' + d.expId + '" recorded at row ' + r;
  } catch(e) { return 'ERROR: ' + e.message; }
}

/** Returns search results across Products, Sales, Expenses as JSON string */
function searchRecords(query) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const q = query.toString().toLowerCase().trim();
    const results = [];

    // Search Products (cols A-C: ID, Brand, Name)
    const psh = ss.getSheetByName('🛍️ Products');
    if (psh) {
      const pData = psh.getRange(ROW.prod, 1, Math.max(psh.getLastRow()-ROW.prod+1,1), 9).getValues();
      pData.forEach((row, i) => {
        const rowText = row.slice(0,6).join(' ').toLowerCase();
        if (row[0] && rowText.includes(q)) {
          results.push({
            sheet:'🛍️ Products', row: ROW.prod+i,
            col1: row[0], col2: row[1], col3: row[2],
            col4: row[4], col5: '₩'+Number(row[7]).toLocaleString(),
            col6: Number(row[8]).toLocaleString()+' UZS',
            label1:'ID', label2:'Brand', label3:'Name',
            label4:'Category', label5:'Cost (KRW)', label6:'Sell (UZS)'
          });
        }
      });
    }

    // Search Sales (cols A-E: Order#, Date, Customer, Contact, Product ID)
    const ssh = ss.getSheetByName('💰 Sales');
    if (ssh) {
      const sData = ssh.getRange(ROW.sales, 1, Math.max(ssh.getLastRow()-ROW.sales+1,1), 13).getValues();
      sData.forEach((row, i) => {
        const rowText = [row[0],row[2],row[3],row[4],row[5],row[8],row[11]].join(' ').toLowerCase();
        if (row[0] && rowText.includes(q)) {
          const dateVal = row[1] instanceof Date ? Utilities.formatDate(row[1],'Asia/Tashkent','dd/MM/yyyy') : row[1];
          results.push({
            sheet:'💰 Sales', row: ROW.sales+i,
            col1: row[0], col2: dateVal, col3: row[2],
            col4: row[4], col5: row[8], col6: Number(row[7]).toLocaleString()+' UZS',
            label1:'Order #', label2:'Date', label3:'Customer',
            label4:'Product ID', label5:'Status', label6:'Total'
          });
        }
      });
    }

    // Search Expenses (cols A-E: ID, Date, Category, Description, Vendor)
    const esh = ss.getSheetByName('💸 Expenses');
    if (esh) {
      const eData = esh.getRange(ROW.exp, 1, Math.max(esh.getLastRow()-ROW.exp+1,1), 11).getValues();
      eData.forEach((row, i) => {
        const rowText = [row[0],row[2],row[3],row[4],row[8]].join(' ').toLowerCase();
        if (row[0] && rowText.includes(q)) {
          const dateVal = row[1] instanceof Date ? Utilities.formatDate(row[1],'Asia/Tashkent','dd/MM/yyyy') : row[1];
          results.push({
            sheet:'💸 Expenses', row: ROW.exp+i,
            col1: row[0], col2: dateVal, col3: row[2],
            col4: row[3].toString().substring(0,35)+(row[3].toString().length>35?'…':''),
            col5: row[4], col6: '$'+Number(row[5]).toFixed(2),
            label1:'Exp ID', label2:'Date', label3:'Category',
            label4:'Description', label5:'Vendor', label6:'Amount'
          });
        }
      });
    }

    return JSON.stringify({ count: results.length, results: results });
  } catch(e) { return JSON.stringify({ count:0, results:[], error: e.message }); }
}

/** Navigate the active sheet to a specific row */
function goToRow(sheetName, rowNum) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (sh) { ss.setActiveSheet(sh); sh.setActiveRange(sh.getRange(rowNum, 1)); }
}

/** Auto-generate next product/order/expense ID */
function getNextId(type) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh, col, prefix;
  if (type==='product')  { sh=ss.getSheetByName('🛍️ Products'); col=1; prefix='P'; }
  if (type==='sale')     { sh=ss.getSheetByName('💰 Sales');     col=1; prefix='ORD-'; }
  if (type==='expense')  { sh=ss.getSheetByName('💸 Expenses');  col=1; prefix='EXP-'; }
  if (!sh) return prefix+'001';
  const vals = sh.getRange(ROW.prod, col, sh.getLastRow(), 1).getValues().flat().filter(v => v!=='');
  const nums = vals.map(v=>parseInt(v.toString().replace(/\D/g,''))).filter(n=>!isNaN(n));
  const next = nums.length ? Math.max(...nums)+1 : 1;
  if (type==='product') return prefix + String(next).padStart(3,'0');
  return prefix + String(next).padStart(3,'0');
}


// ═════════════════════════════════════════════════════════════════════════════
// ③  HTML DIALOG LAUNCHERS
// ═════════════════════════════════════════════════════════════════════════════
function showAddProductForm() {
  const nextId = getNextId('product');
  const html = HtmlService.createHtmlOutput(getProductFormHtml(nextId))
    .setWidth(520).setHeight(580).setTitle('➕ Add New Product');
  SpreadsheetApp.getUi().showModalDialog(html, '➕ Add New Product');
}

function showAddSaleForm() {
  const nextId = getNextId('sale');
  const html = HtmlService.createHtmlOutput(getSaleFormHtml(nextId))
    .setWidth(520).setHeight(620).setTitle('🛒 Log a Sale / Order');
  SpreadsheetApp.getUi().showModalDialog(html, '🛒 Log a Sale / Order');
}

function showAddExpenseForm() {
  const nextId = getNextId('expense');
  const html = HtmlService.createHtmlOutput(getExpenseFormHtml(nextId))
    .setWidth(520).setHeight(600).setTitle('💸 Record an Expense');
  SpreadsheetApp.getUi().showModalDialog(html, '💸 Record an Expense');
}

function showSearchDialog() {
  const html = HtmlService.createHtmlOutput(getSearchHtml())
    .setWidth(680).setHeight(560).setTitle('🔍 Search Records');
  SpreadsheetApp.getUi().showModalDialog(html, '🔍 Search Records');
}


// ═════════════════════════════════════════════════════════════════════════════
// ④  HTML TEMPLATES
// ═════════════════════════════════════════════════════════════════════════════

function getProductFormHtml(nextId) {
  return `<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
  *{box-sizing:border-box;margin:0;padding:0;font-family:Arial,sans-serif}
  body{background:#f4f6fb;padding:18px}
  h2{color:#1B3A6B;font-size:15px;margin-bottom:14px;border-bottom:2px solid #C9A84C;padding-bottom:6px}
  .row{display:flex;gap:10px;margin-bottom:11px}
  .field{flex:1;display:flex;flex-direction:column}
  label{font-size:11px;font-weight:bold;color:#333;margin-bottom:3px}
  input,select{padding:7px 9px;border:1px solid #ccc;border-radius:5px;font-size:12px;
    background:#fff;color:#1A5276;outline:none;transition:border .2s}
  input:focus,select:focus{border-color:#1B3A6B;box-shadow:0 0 0 2px #d6e4f0}
  .hint{font-size:10px;color:#888;margin-top:2px}
  .btns{display:flex;gap:10px;margin-top:16px}
  .btn-save{flex:1;background:#1B3A6B;color:#fff;border:none;padding:10px;border-radius:6px;
    font-size:13px;font-weight:bold;cursor:pointer;transition:background .2s}
  .btn-save:hover{background:#2E5E9E}
  .btn-cancel{flex:0 0 100px;background:#eee;color:#555;border:none;padding:10px;
    border-radius:6px;font-size:13px;cursor:pointer}
  .msg{margin-top:10px;padding:8px;border-radius:5px;font-size:12px;display:none}
  .msg.ok{background:#d5f5e3;color:#1E8449;display:block}
  .msg.err{background:#fdecea;color:#C0392B;display:block}
  .loader{display:none;text-align:center;color:#888;font-size:12px;margin-top:8px}
</style></head><body>
<h2>➕ Add New Product</h2>
<form id="frm">
  <div class="row">
    <div class="field">
      <label>Product ID *</label>
      <input id="id" value="${nextId}" required>
      <span class="hint">Auto-generated — change if needed</span>
    </div>
    <div class="field">
      <label>Brand *</label>
      <input id="brand" placeholder="e.g. COSRX" required>
    </div>
  </div>
  <div class="row">
    <div class="field">
      <label>Product Name *</label>
      <input id="name" placeholder="e.g. Advanced Snail 96 Mucin Essence" required>
    </div>
  </div>
  <div class="row">
    <div class="field">
      <label>Variant</label>
      <input id="variant" placeholder="e.g. 100 ml / SPF50+">
    </div>
    <div class="field">
      <label>Category *</label>
      <select id="category" required>
        <option value="">— Select —</option>
        <option>Toner</option><option>Serum</option><option>Essence</option>
        <option>Moisturizer</option><option>Sunscreen</option><option>Cleanser</option>
        <option>Mask</option><option>Eye Cream</option><option>Other</option>
      </select>
    </div>
    <div class="field">
      <label>Skin Type</label>
      <select id="skinType">
        <option value="">— Select —</option>
        <option>All</option><option>Dry</option><option>Oily</option>
        <option>Sensitive</option><option>Normal/Dry</option><option>Oily/Combo</option>
      </select>
    </div>
  </div>
  <div class="row">
    <div class="field">
      <label>Weight (g) *</label>
      <input id="weight" type="number" min="1" placeholder="e.g. 100" required>
      <span class="hint">Used to calculate real shipping cost</span>
    </div>
    <div class="field">
      <label>Cost Price (KRW) *</label>
      <input id="costKrw" type="number" min="0" placeholder="e.g. 15000" required>
    </div>
    <div class="field">
      <label>Selling Price (UZS) *</label>
      <input id="sellUzs" type="number" min="0" placeholder="e.g. 220000" required>
    </div>
  </div>
  <div class="btns">
    <button class="btn-save" type="submit">💾  Save Product</button>
    <button class="btn-cancel" type="button" onclick="google.script.host.close()">Cancel</button>
  </div>
  <div id="msg" class="msg"></div>
  <div id="loader" class="loader">⏳ Saving…</div>
</form>
<script>
document.getElementById('frm').onsubmit = function(e) {
  e.preventDefault();
  var d = {
    id:       document.getElementById('id').value.trim(),
    brand:    document.getElementById('brand').value.trim(),
    name:     document.getElementById('name').value.trim(),
    variant:  document.getElementById('variant').value.trim(),
    category: document.getElementById('category').value,
    skinType: document.getElementById('skinType').value,
    weight:   document.getElementById('weight').value,
    costKrw:  document.getElementById('costKrw').value,
    sellUzs:  document.getElementById('sellUzs').value
  };
  if (!d.id||!d.brand||!d.name||!d.category||!d.weight||!d.costKrw||!d.sellUzs) {
    showMsg('Please fill in all required (*) fields.','err'); return;
  }
  document.getElementById('loader').style.display='block';
  google.script.run
    .withSuccessHandler(function(r){
      document.getElementById('loader').style.display='none';
      if(r.startsWith('OK')){
        showMsg('✅ '+r.replace('OK: ',''),'ok');
        document.getElementById('frm').reset();
        document.getElementById('id').value='';
      } else showMsg(r,'err');
    })
    .withFailureHandler(function(e){
      document.getElementById('loader').style.display='none';
      showMsg('Error: '+e.message,'err');
    })
    .saveProduct(d);
};
function showMsg(txt,type){
  var m=document.getElementById('msg');
  m.textContent=txt; m.className='msg '+type;
}
</script></body></html>`;
}

// ─────────────────────────────────────────────────────────────────────────────
function getSaleFormHtml(nextId) {
  const today = new Date().toISOString().split('T')[0];
  return `<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
  *{box-sizing:border-box;margin:0;padding:0;font-family:Arial,sans-serif}
  body{background:#f0faf4;padding:18px}
  h2{color:#1E8449;font-size:15px;margin-bottom:14px;border-bottom:2px solid #1E8449;padding-bottom:6px}
  .row{display:flex;gap:10px;margin-bottom:11px}
  .field{flex:1;display:flex;flex-direction:column}
  label{font-size:11px;font-weight:bold;color:#333;margin-bottom:3px}
  input,select,textarea{padding:7px 9px;border:1px solid #ccc;border-radius:5px;
    font-size:12px;background:#fff;color:#1A5276;outline:none;transition:border .2s}
  input:focus,select:focus{border-color:#1E8449;box-shadow:0 0 0 2px #d5f5e3}
  .hint{font-size:10px;color:#888;margin-top:2px}
  .btns{display:flex;gap:10px;margin-top:16px}
  .btn-save{flex:1;background:#1E8449;color:#fff;border:none;padding:10px;border-radius:6px;
    font-size:13px;font-weight:bold;cursor:pointer}
  .btn-cancel{flex:0 0 100px;background:#eee;color:#555;border:none;padding:10px;border-radius:6px;font-size:13px;cursor:pointer}
  .msg{margin-top:10px;padding:8px;border-radius:5px;font-size:12px;display:none}
  .msg.ok{background:#d5f5e3;color:#1E8449;display:block}
  .msg.err{background:#fdecea;color:#C0392B;display:block}
  .loader{display:none;text-align:center;color:#888;font-size:12px;margin-top:8px}
</style></head><body>
<h2>🛒 Log a Sale / Order</h2>
<form id="frm">
  <div class="row">
    <div class="field">
      <label>Order # *</label>
      <input id="orderId" value="${nextId}" required>
    </div>
    <div class="field">
      <label>Order Date *</label>
      <input id="date" type="date" value="${today}" required>
    </div>
  </div>
  <div class="row">
    <div class="field">
      <label>Customer Name *</label>
      <input id="customer" placeholder="e.g. Dilnoza T." required>
    </div>
    <div class="field">
      <label>Telegram / Contact</label>
      <input id="contact" placeholder="e.g. @dilnoza_uz">
    </div>
  </div>
  <div class="row">
    <div class="field">
      <label>Product ID *</label>
      <input id="productId" placeholder="e.g. P001" required>
      <span class="hint">Name &amp; price auto-fill from Products sheet</span>
    </div>
    <div class="field">
      <label>Quantity *</label>
      <input id="qty" type="number" min="1" value="1" required>
    </div>
  </div>
  <div class="row">
    <div class="field">
      <label>Status *</label>
      <select id="status" required>
        <option>Pending</option><option>Processing</option>
        <option>Shipped</option><option>Delivered</option>
        <option>Cancelled</option><option>Refunded</option>
      </select>
    </div>
    <div class="field">
      <label>Shipping Method</label>
      <input id="shipping" value="Uzbekistan Post">
    </div>
  </div>
  <div class="row">
    <div class="field">
      <label>Tracking Number</label>
      <input id="tracking" placeholder="e.g. UZ123456">
    </div>
    <div class="field">
      <label>Region (Uzbekistan)</label>
      <select id="region">
        <option value="">— Select —</option>
        <option>Tashkent</option><option>Samarkand</option><option>Bukhara</option>
        <option>Fergana</option><option>Andijan</option><option>Namangan</option>
        <option>Nukus</option><option>Termez</option><option>Other</option>
      </select>
    </div>
  </div>
  <div class="row">
    <div class="field">
      <label>Notes</label>
      <input id="notes" placeholder="Any additional notes…">
    </div>
  </div>
  <div class="btns">
    <button class="btn-save" type="submit">💾  Log Order</button>
    <button class="btn-cancel" type="button" onclick="google.script.host.close()">Cancel</button>
  </div>
  <div id="msg" class="msg"></div>
  <div id="loader" class="loader">⏳ Saving…</div>
</form>
<script>
document.getElementById('frm').onsubmit = function(e) {
  e.preventDefault();
  var d = {
    orderId:  document.getElementById('orderId').value.trim(),
    date:     document.getElementById('date').value,
    customer: document.getElementById('customer').value.trim(),
    contact:  document.getElementById('contact').value.trim(),
    productId:document.getElementById('productId').value.trim().toUpperCase(),
    qty:      document.getElementById('qty').value,
    status:   document.getElementById('status').value,
    shipping: document.getElementById('shipping').value.trim(),
    tracking: document.getElementById('tracking').value.trim(),
    region:   document.getElementById('region').value,
    notes:    document.getElementById('notes').value.trim()
  };
  document.getElementById('loader').style.display='block';
  google.script.run
    .withSuccessHandler(function(r){
      document.getElementById('loader').style.display='none';
      if(r.startsWith('OK')){ showMsg('✅ '+r.replace('OK: ',''),'ok'); document.getElementById('frm').reset(); document.getElementById('orderId').value=''; }
      else showMsg(r,'err');
    })
    .withFailureHandler(function(e){ document.getElementById('loader').style.display='none'; showMsg('Error: '+e.message,'err'); })
    .saveSale(d);
};
function showMsg(txt,type){ var m=document.getElementById('msg'); m.textContent=txt; m.className='msg '+type; }
</script></body></html>`;
}

// ─────────────────────────────────────────────────────────────────────────────
function getExpenseFormHtml(nextId) {
  const today = new Date().toISOString().split('T')[0];
  return `<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
  *{box-sizing:border-box;margin:0;padding:0;font-family:Arial,sans-serif}
  body{background:#fdf2f2;padding:18px}
  h2{color:#C0392B;font-size:15px;margin-bottom:14px;border-bottom:2px solid #C0392B;padding-bottom:6px}
  .row{display:flex;gap:10px;margin-bottom:11px}
  .field{flex:1;display:flex;flex-direction:column}
  label{font-size:11px;font-weight:bold;color:#333;margin-bottom:3px}
  input,select{padding:7px 9px;border:1px solid #ccc;border-radius:5px;
    font-size:12px;background:#fff;color:#1A5276;outline:none;transition:border .2s}
  input:focus,select:focus{border-color:#C0392B;box-shadow:0 0 0 2px #fdecea}
  .hint{font-size:10px;color:#888;margin-top:2px}
  .btns{display:flex;gap:10px;margin-top:16px}
  .btn-save{flex:1;background:#C0392B;color:#fff;border:none;padding:10px;border-radius:6px;
    font-size:13px;font-weight:bold;cursor:pointer}
  .btn-cancel{flex:0 0 100px;background:#eee;color:#555;border:none;padding:10px;border-radius:6px;font-size:13px;cursor:pointer}
  .msg{margin-top:10px;padding:8px;border-radius:5px;font-size:12px;display:none}
  .msg.ok{background:#d5f5e3;color:#1E8449;display:block}
  .msg.err{background:#fdecea;color:#C0392B;display:block}
  .loader{display:none;text-align:center;color:#888;font-size:12px;margin-top:8px}
</style></head><body>
<h2>💸 Record an Expense</h2>
<form id="frm">
  <div class="row">
    <div class="field">
      <label>Expense ID *</label>
      <input id="expId" value="${nextId}" required>
    </div>
    <div class="field">
      <label>Date *</label>
      <input id="date" type="date" value="${today}" required>
    </div>
  </div>
  <div class="row">
    <div class="field">
      <label>Category *</label>
      <select id="category" required>
        <option value="">— Select —</option>
        <option>Product Purchase</option>
        <option>Shipping / Freight</option>
        <option>Posting Guy Salary</option>
        <option>Platform Fee</option>
        <option>Marketing</option>
        <option>Packaging</option>
        <option>Customs / Duties</option>
        <option>Other</option>
      </select>
    </div>
    <div class="field">
      <label>Amount (USD) *</label>
      <input id="amount" type="number" min="0" step="0.01" placeholder="e.g. 150.00" required>
    </div>
  </div>
  <div class="row">
    <div class="field">
      <label>Description *</label>
      <input id="description" placeholder="e.g. COSRX Essence × 30 units" required>
    </div>
  </div>
  <div class="row">
    <div class="field">
      <label>Paid To / Vendor</label>
      <input id="vendor" placeholder="e.g. Olive Young KR">
    </div>
    <div class="field">
      <label>Payment Method</label>
      <select id="payMethod">
        <option>Bank Transfer</option>
        <option>Card</option>
        <option>Cash</option>
        <option>UZ Bank Wire</option>
        <option>PayPal</option>
        <option>Other</option>
      </select>
    </div>
  </div>
  <div class="row">
    <div class="field">
      <label>Receipt / Ref #</label>
      <input id="receipt" placeholder="e.g. RC-009">
    </div>
    <div class="field">
      <label>Paid By</label>
      <select id="paidBy">
        <option>Partner 1 (OJ)</option>
        <option>Partner 2</option>
        <option>Partner 3</option>
        <option>Partner 4</option>
        <option>Posting Guy UZ</option>
      </select>
    </div>
  </div>
  <div class="row">
    <div class="field">
      <label>Notes</label>
      <input id="notes" placeholder="Any additional notes…">
    </div>
  </div>
  <div class="btns">
    <button class="btn-save" type="submit">💾  Save Expense</button>
    <button class="btn-cancel" type="button" onclick="google.script.host.close()">Cancel</button>
  </div>
  <div id="msg" class="msg"></div>
  <div id="loader" class="loader">⏳ Saving…</div>
</form>
<script>
document.getElementById('frm').onsubmit = function(e) {
  e.preventDefault();
  var d = {
    expId:      document.getElementById('expId').value.trim(),
    date:       document.getElementById('date').value,
    category:   document.getElementById('category').value,
    description:document.getElementById('description').value.trim(),
    vendor:     document.getElementById('vendor').value.trim(),
    amount:     document.getElementById('amount').value,
    payMethod:  document.getElementById('payMethod').value,
    receipt:    document.getElementById('receipt').value.trim(),
    paidBy:     document.getElementById('paidBy').value,
    notes:      document.getElementById('notes').value.trim()
  };
  if(!d.expId||!d.date||!d.category||!d.amount||!d.description){
    showMsg('Please fill in all required (*) fields.','err'); return;
  }
  document.getElementById('loader').style.display='block';
  google.script.run
    .withSuccessHandler(function(r){
      document.getElementById('loader').style.display='none';
      if(r.startsWith('OK')){ showMsg('✅ '+r.replace('OK: ',''),'ok'); document.getElementById('frm').reset(); document.getElementById('expId').value=''; }
      else showMsg(r,'err');
    })
    .withFailureHandler(function(e){ document.getElementById('loader').style.display='none'; showMsg('Error: '+e.message,'err'); })
    .saveExpense(d);
};
function showMsg(txt,type){ var m=document.getElementById('msg'); m.textContent=txt; m.className='msg '+type; }
</script></body></html>`;
}

// ─────────────────────────────────────────────────────────────────────────────
function getSearchHtml() {
  return `<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
  *{box-sizing:border-box;margin:0;padding:0;font-family:Arial,sans-serif}
  body{background:#f4f6fb;padding:16px}
  h2{color:#1B3A6B;font-size:15px;margin-bottom:12px;border-bottom:2px solid #C9A84C;padding-bottom:6px}
  .search-bar{display:flex;gap:8px;margin-bottom:12px}
  #query{flex:1;padding:9px 12px;border:2px solid #1B3A6B;border-radius:6px;
    font-size:13px;outline:none}
  #query:focus{border-color:#C9A84C;box-shadow:0 0 0 2px #fff5dc}
  .btn-search{background:#1B3A6B;color:#fff;border:none;padding:9px 18px;
    border-radius:6px;font-size:13px;font-weight:bold;cursor:pointer}
  .btn-search:hover{background:#2E5E9E}
  .filter-bar{display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap}
  .filter-btn{padding:4px 12px;border-radius:20px;border:1px solid #ccc;
    background:#fff;font-size:11px;cursor:pointer;transition:all .2s}
  .filter-btn.active{background:#1B3A6B;color:#fff;border-color:#1B3A6B}
  #status-bar{font-size:11px;color:#888;margin-bottom:8px;min-height:16px}
  #results{overflow-y:auto;max-height:320px}
  .card{background:#fff;border:1px solid #e0e0e0;border-radius:6px;
    padding:10px 12px;margin-bottom:7px;cursor:pointer;transition:all .15s}
  .card:hover{border-color:#1B3A6B;box-shadow:0 2px 6px rgba(27,58,107,.1)}
  .card-top{display:flex;justify-content:space-between;align-items:center;margin-bottom:5px}
  .sheet-tag{font-size:10px;font-weight:bold;padding:2px 8px;border-radius:10px;color:#fff}
  .tag-prod{background:#7D3C98}.tag-sales{background:#1E8449}.tag-exp{background:#C0392B}
  .card-id{font-weight:bold;color:#1B3A6B;font-size:12px}
  .card-grid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:4px 12px}
  .card-field{font-size:11px;color:#555}
  .card-field span{font-weight:bold;color:#333}
  .card-right{font-size:12px;font-weight:bold;color:#1E8449}
  .no-results{text-align:center;padding:30px;color:#aaa;font-size:13px}
  .loader{display:none;text-align:center;padding:20px;color:#888;font-size:13px}
  .hint{font-size:10px;color:#aaa;margin-top:4px}
</style></head><body>
<h2>🔍 Search Records</h2>
<div class="search-bar">
  <input id="query" type="text" placeholder="Search products, customers, orders, expenses…" autofocus>
  <button class="btn-search" onclick="doSearch()">Search</button>
</div>
<div class="filter-bar">
  <span style="font-size:11px;color:#666;line-height:24px">Filter:</span>
  <button class="filter-btn active" data-f="all"   onclick="setFilter('all',this)">All</button>
  <button class="filter-btn"        data-f="prod"  onclick="setFilter('prod',this)">🛍️ Products</button>
  <button class="filter-btn"        data-f="sales" onclick="setFilter('sales',this)">💰 Sales</button>
  <button class="filter-btn"        data-f="exp"   onclick="setFilter('exp',this)">💸 Expenses</button>
</div>
<div id="status-bar">Type a keyword and click Search (or press Enter)</div>
<div id="loader" class="loader">⏳ Searching…</div>
<div id="results"></div>
<div class="hint">💡 Click any result card to jump to that row in the sheet</div>
<script>
var allResults = [];
var activeFilter = 'all';

document.getElementById('query').addEventListener('keydown', function(e){
  if(e.key==='Enter') doSearch();
});

function setFilter(f, btn) {
  activeFilter = f;
  document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  renderResults();
}

function doSearch() {
  var q = document.getElementById('query').value.trim();
  if (!q) { document.getElementById('status-bar').textContent='Please enter a search term.'; return; }
  document.getElementById('loader').style.display='block';
  document.getElementById('results').innerHTML='';
  document.getElementById('status-bar').textContent='Searching…';
  google.script.run
    .withSuccessHandler(function(raw){
      document.getElementById('loader').style.display='none';
      var data = JSON.parse(raw);
      allResults = data.results;
      document.getElementById('status-bar').textContent =
        data.count===0 ? 'No results found for "'+q+'"'
        : 'Found '+data.count+' result'+(data.count!==1?'s':'');
      renderResults();
    })
    .withFailureHandler(function(e){
      document.getElementById('loader').style.display='none';
      document.getElementById('status-bar').textContent='Error: '+e.message;
    })
    .searchRecords(q);
}

function renderResults() {
  var container = document.getElementById('results');
  container.innerHTML='';
  var filtered = allResults.filter(function(r){
    if(activeFilter==='all') return true;
    if(activeFilter==='prod')  return r.sheet==='🛍️ Products';
    if(activeFilter==='sales') return r.sheet==='💰 Sales';
    if(activeFilter==='exp')   return r.sheet==='💸 Expenses';
    return true;
  });
  if(filtered.length===0){
    container.innerHTML='<div class="no-results">😕 No results in this category</div>';
    return;
  }
  filtered.forEach(function(r){
    var tagClass = r.sheet.includes('Products')?'tag-prod': r.sheet.includes('Sales')?'tag-sales':'tag-exp';
    var div = document.createElement('div');
    div.className='card';
    div.innerHTML=
      '<div class="card-top">'
       +'<div><span class="sheet-tag '+tagClass+'">'+r.sheet+'</span>'
       +'&nbsp;&nbsp;<span class="card-id">'+r.col1+'</span></div>'
       +'<span class="card-right">'+r.col6+'</span>'
      +'</div>'
      +'<div class="card-grid">'
       +'<div class="card-field">'+r.label2+': <span>'+r.col2+'</span></div>'
       +'<div class="card-field">'+r.label3+': <span>'+r.col3+'</span></div>'
       +'<div class="card-field">'+r.label4+': <span>'+r.col4+'</span></div>'
       +'<div class="card-field">'+r.label5+': <span>'+r.col5+'</span></div>'
      +'</div>';
    div.onclick = function(){ google.script.run.goToRow(r.sheet, r.row); };
    container.appendChild(div);
  });
}
</script></body></html>`;
}


// ═════════════════════════════════════════════════════════════════════════════
// ⑤  SPREADSHEET BUILDER  (run once)
// ═════════════════════════════════════════════════════════════════════════════

function buildAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetName('K-Beauty UZ — Accounting');
  buildSettings(ss);
  buildProductDB(ss);
  buildInventory(ss);
  buildSales(ss);
  buildExpenses(ss);
  buildPartnerDashboard(ss);
  buildDashboard(ss);
  const def = ss.getSheetByName('Sheet1');
  if (def) ss.deleteSheet(def);
  ss.setActiveSheet(ss.getSheetByName('📊 Dashboard'));
  SpreadsheetApp.flush();
  onOpen(); // register menu immediately
  SpreadsheetApp.getUi().alert(
    '✅ Done!\n\n' +
    'Your K-Beauty UZ spreadsheet is ready.\n\n' +
    '▸ Use the "K-Beauty UZ ▾" menu at the top to:\n' +
    '  • Add products\n  • Log sales\n  • Record expenses\n  • Search anything\n\n' +
    'Blue cells = your inputs.  Gray cells = auto-calculated.'
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────────────────────────────────────
function getOrCreate(ss,name,color){
  let sh=ss.getSheetByName(name);
  if(!sh) sh=ss.insertSheet(name);
  sh.setTabColor(color); sh.clear(); sh.clearFormats();
  sh.setFrozenRows(0); sh.setFrozenColumns(0); return sh;
}
function titleBlock(sh,title,subtitle){
  sh.getRange(1,1,1,13).merge().setValue(title)
    .setBackground(C.navy).setFontColor(C.white).setFontSize(16)
    .setFontWeight('bold').setFontFamily('Arial').setVerticalAlignment('middle');
  sh.setRowHeight(1,42);
  sh.getRange(2,1,1,13).merge().setValue(subtitle)
    .setBackground('#2C4F85').setFontColor('#BDC3C7').setFontSize(9)
    .setFontFamily('Arial').setFontStyle('italic').setVerticalAlignment('middle');
  sh.setRowHeight(2,18);
}
function sectionHeader(sh,row,col,span,text,bg,tc){
  sh.getRange(row,col,1,span).merge().setValue(text)
    .setBackground(bg||C.navy).setFontColor(tc||C.white)
    .setFontSize(11).setFontWeight('bold').setFontFamily('Arial')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.setRowHeight(row,28);
}
function headerRow(sh,row,headers,bg,tc){
  headers.forEach((h,i)=>{
    sh.getRange(row,i+1).setValue(h).setBackground(bg||C.navy)
      .setFontColor(tc||C.white).setFontWeight('bold').setFontFamily('Arial')
      .setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  });
  sh.setRowHeight(row,36);
}
function altRow(sh,row,cols,alt){
  sh.getRange(row,1,1,cols).setBackground(alt?C.rowAlt:C.white);
}
function borders(sh,row,col,rows,cols){
  sh.getRange(row,col,rows,cols).setBorder(true,true,true,true,true,true,'#CCCCCC',SpreadsheetApp.BorderStyle.SOLID);
}


// ─────────────────────────────────────────────────────────────────────────────
// ⚙️ SETTINGS SHEET
// ─────────────────────────────────────────────────────────────────────────────
function buildSettings(ss){
  const sh=getOrCreate(ss,'⚙️ Settings',C.purple);
  titleBlock(sh,'⚙️  Settings & Configuration','Update these values — all sheets recalculate automatically');
  sh.setColumnWidth(1,240); sh.setColumnWidth(2,200); sh.setColumnWidth(3,320);
  sectionHeader(sh,4,1,3,'  💱  EXCHANGE RATES & SHIPPING',C.navy);
  const rows=[
    ['KRW → UZS Exchange Rate',10.5,'UZS per 1 KRW  ←  update when rate changes',FMT_NUM,'KRW_UZS_RATE'],
    ['Korea → UZ Shipping (KRW/kg)',13000,'KRW per kg  ←  update when shipping changes',FMT_NUM,'SHIPPING_KRW_KG'],
    ['USD → UZS Exchange Rate',12700,'UZS per 1 USD',FMT_NUM,'USD_UZS_RATE'],
  ];
  rows.forEach(([lbl,val,note,fmt,nm],i)=>{
    const r=5+i; sh.setRowHeight(r,22);
    sh.getRange(r,1).setValue(lbl).setFontWeight('bold').setFontFamily('Arial').setFontSize(10);
    sh.getRange(r,2).setValue(val).setBackground(C.yellow).setFontColor('#1A5276')
      .setFontFamily('Arial').setFontSize(10).setNumberFormat(fmt).setNote(nm);
    sh.getRange(r,3).setValue(note).setFontFamily('Arial').setFontSize(9).setFontColor('#666').setFontStyle('italic');
  });
  try{
    ss.setNamedRange('KRW_UZS_RATE',   sh.getRange('B5'));
    ss.setNamedRange('SHIPPING_KRW_KG',sh.getRange('B6'));
    ss.setNamedRange('USD_UZS_RATE',   sh.getRange('B7'));
  }catch(e){}

  sectionHeader(sh,10,1,3,'  🏢  BUSINESS INFO',C.navy);
  const biz=[
    ['Business Name','Muna Cosmetics'],['Launch Date','01/01/2026'],
    ['Founder / Partner 1','OJ'],['Partner 2 Name','Partner 2'],
    ['Partner 3 Name','Partner 3'],['Partner 4 Name','Partner 4'],
    ['Posting Guy (UZ)','Posting Guy UZ'],['Monthly Salary – Posting Guy (USD)',150],
    ['Sales Channel','Instagram / Telegram'],['Telegram Handle','@muna_cosmetics_uz'],
  ];
  biz.forEach(([lbl,val],i)=>{
    const r=11+i; sh.setRowHeight(r,22);
    sh.getRange(r,1).setValue(lbl).setFontWeight('bold').setFontFamily('Arial').setFontSize(10);
    sh.getRange(r,2).setValue(val).setBackground(C.inputBlue).setFontColor('#1A5276')
      .setFontFamily('Arial').setFontSize(10);
  });
  borders(sh,4,1,20,3);
  sh.setFrozenRows(2);
}

// ─────────────────────────────────────────────────────────────────────────────
// 🛍️ PRODUCT DATABASE
// ─────────────────────────────────────────────────────────────────────────────
function buildProductDB(ss){
  const sh=getOrCreate(ss,'🛍️ Products',C.purple);
  titleBlock(sh,'🛍️  Product Database',
    'Your Korean cosmetics catalogue  |  Use menu "K-Beauty UZ ▾ → Add New Product" to add rows');
  sh.getRange(4,1,1,11).merge()
    .setValue('⚡  Real Cost and Profit auto-calculate using exchange rate & shipping from ⚙️ Settings. Add rows freely — formulas are pre-filled 100 rows down.')
    .setBackground('#EAF2FF').setFontColor('#1A5276').setFontSize(9).setFontStyle('italic').setFontFamily('Arial');
  sh.setRowHeight(4,20);
  [50,160,220,80,110,80,80,130,160,130,130].forEach((w,i)=>sh.setColumnWidth(i+1,w));
  headerRow(sh,5,['ID','Brand','Product Name','Variant','Category','Skin Type',
    'Weight\n(g)','Cost Price\n(KRW)','Selling Price\n(UZS)','Real Cost\n(UZS) ⚡','Profit\n(UZS) ⚡'],C.navy);

  const products=[
    ['P001','Round Lab',        'Birch Juice Moisturizing Toner',          '200 ml', 'Toner',    'All',        200,15000,220000],
    ['P002','Beauty of Joseon', 'Revive Serum: Ginseng + Snail Mucin',     '30 ml',  'Serum',    'All',         55,18000,270000],
    ['P003','Beauty of Joseon', 'Relief Sun: Rice + Probiotics SPF50+',    'SPF 50+','Sunscreen','Sensitive',   65,16000,240000],
    ['P004','Innisfree',        'Green Tea Seed Serum',                    '80 ml',  'Serum',    'Normal/Dry', 125,22000,330000],
    ['P005','COSRX',            'Advanced Snail 96 Mucin Essence',         '100 ml', 'Essence',  'All',        185,20000,300000],
    ['P006','Klairs',           'Supple Preparation Unscented Toner',      '180 ml', 'Toner',    'Sensitive',  225,17000,255000],
    ['P007','Some By Mi',       'AHA BHA PHA 30 Days Miracle Toner',       '150 ml', 'Toner',    'Oily/Combo', 195,14000,210000],
    ['P008','Purito',           'Centella Green Level Sunscreen',          'SPF 50+','Sunscreen','Sensitive',   65,15500,235000],
  ];
  const DS=ROW.prod;
  products.forEach((p,i)=>{
    const r=DS+i; altRow(sh,r,11,i%2===1); sh.setRowHeight(r,22);
    p.forEach((val,ci)=>{
      const c=sh.getRange(r,ci+1);
      c.setValue(val).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
      if(ci<6)  c.setFontColor('#1A5276').setBackground(C.inputBlue);
      if(ci===7) c.setNumberFormat(FMT_KRW).setFontColor('#1A5276').setBackground(C.inputBlue);
      if(ci===8) c.setNumberFormat(FMT_UZS).setFontColor('#1A5276').setBackground(C.inputBlue);
    });
    sh.getRange(r,10).setFormula(`=IF(A${r}="","",H${r}*KRW_UZS_RATE+(G${r}/1000)*SHIPPING_KRW_KG*KRW_UZS_RATE)`)
      .setNumberFormat(FMT_UZS).setBackground(C.formulaGray).setFontFamily('Arial');
    sh.getRange(r,11).setFormula(`=IF(A${r}="","",I${r}-J${r})`)
      .setNumberFormat(FMT_UZS).setBackground(C.formulaGray).setFontFamily('Arial');
  });
  // 100 pre-formatted blank rows
  for(let i=0;i<100;i++){
    const r=DS+products.length+i; altRow(sh,r,11,i%2===1); sh.setRowHeight(r,20);
    [1,2,3,4,5,6,7,8,9].forEach(c=>sh.getRange(r,c).setBackground(i%2===1?C.rowAlt:C.white));
    sh.getRange(r,7).setBackground(i%2===1?C.rowAlt:C.white).setNumberFormat(FMT_NUM);
    sh.getRange(r,8).setBackground(i%2===1?C.rowAlt:C.white).setNumberFormat(FMT_KRW);
    sh.getRange(r,9).setBackground(i%2===1?C.rowAlt:C.white).setNumberFormat(FMT_UZS);
    sh.getRange(r,10).setFormula(`=IF(A${r}="","",H${r}*KRW_UZS_RATE+(G${r}/1000)*SHIPPING_KRW_KG*KRW_UZS_RATE)`)
      .setNumberFormat(FMT_UZS).setBackground(C.formulaGray).setFontFamily('Arial');
    sh.getRange(r,11).setFormula(`=IF(A${r}="","",I${r}-J${r})`)
      .setNumberFormat(FMT_UZS).setBackground(C.formulaGray).setFontFamily('Arial');
  }
  // Totals
  const TR=DS+products.length+100;
  sh.getRange(TR,1,1,11).setBackground(C.navy).setFontColor(C.white).setFontWeight('bold').setFontFamily('Arial');
  sh.getRange(TR,1).setValue('TOTALS');
  [[7,FMT_NUM],[8,FMT_KRW],[9,FMT_UZS],[10,FMT_UZS],[11,FMT_UZS]].forEach(([col,fmt])=>{
    sh.getRange(TR,col).setFormula(`=SUM(${String.fromCharCode(64+col)}${DS}:${String.fromCharCode(64+col)}${TR-1})`)
      .setNumberFormat(fmt).setBackground(C.navy).setFontColor(C.white).setFontFamily('Arial');
  });
  sh.setRowHeight(TR,24);
  borders(sh,5,1,TR-4,11);
  sh.setFrozenRows(5); sh.setFrozenColumns(1);
}

// ─────────────────────────────────────────────────────────────────────────────
// 📦 INVENTORY
// ─────────────────────────────────────────────────────────────────────────────
function buildInventory(ss){
  const sh=getOrCreate(ss,'📦 Inventory','#2E86C1');
  titleBlock(sh,'📦  Stock & Inventory',
    'Just type a Product ID — name & prices fill automatically via lookup');
  [55,190,110,100,100,100,130,150,120,110,130].forEach((w,i)=>sh.setColumnWidth(i+1,w));
  // KPI bar
  sh.setRowHeight(4,24); sh.setRowHeight(5,32); sh.setRowHeight(6,6);
  [['Total SKUs',`=COUNTA(A${ROW.inv}:A508)`,'#1B3A6B'],
   ['Units In Stock',`=SUMIF(A${ROW.inv}:A508,"<>",F${ROW.inv}:F508)`,'#2E86C1'],
   ['Low Stock / Out',`=COUNTIFS(A${ROW.inv}:A508,"<>",F${ROW.inv}:F508,"<="&J${ROW.inv}:J508)`,'#C0392B']
  ].forEach(([lbl,f,bg],i)=>{
    sh.getRange(4,i*4+1,1,4).merge().setValue(lbl).setBackground(bg).setFontColor(C.white).setFontWeight('bold').setFontFamily('Arial').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sh.getRange(5,i*4+1,1,4).merge().setFormula(f).setNumberFormat(FMT_NUM).setBackground('#EBF5FB').setFontColor(bg).setFontWeight('bold').setFontFamily('Arial').setFontSize(16).setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
  headerRow(sh,7,['Product ID','Product Name','Category','Units\nBought','Units\nSold',
    'Units\nLeft','Cost\n(KRW)','Sell Price\n(UZS)','Last\nRestocked','Reorder\nLevel','Status'],C.navy);
  const invData=[
    ['P001',50,20,'05/01/2026',10],['P002',40,15,'05/01/2026',10],
    ['P003',35,12,'05/01/2026', 8],['P004',25, 8,'05/01/2026', 5],
    ['P005',30,10,'05/01/2026', 8],['P006',20, 6,'05/01/2026', 5],
    ['P007',28,12,'05/01/2026', 8],['P008',22, 7,'05/01/2026', 5],
  ];
  const DS=ROW.inv;
  invData.forEach((d,i)=>{
    const r=DS+i; altRow(sh,r,11,i%2===1); sh.setRowHeight(r,22);
    sh.getRange(r,1).setValue(d[0]).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial').setFontSize(10);
    sh.getRange(r,2).setFormula(`=IF(A${r}="","",IFERROR(VLOOKUP(A${r},'🛍️ Products'!A:C,2,0),"?"))`).setBackground(C.formulaGray).setFontFamily('Arial').setFontSize(10);
    sh.getRange(r,3).setFormula(`=IF(A${r}="","",IFERROR(VLOOKUP(A${r},'🛍️ Products'!A:E,5,0),""))`).setBackground(C.formulaGray).setFontFamily('Arial').setFontSize(10);
    sh.getRange(r,4).setValue(d[1]).setBackground(C.inputBlue).setFontColor('#1A5276').setNumberFormat(FMT_NUM).setFontFamily('Arial');
    sh.getRange(r,5).setValue(d[2]).setBackground(C.inputBlue).setFontColor('#1A5276').setNumberFormat(FMT_NUM).setFontFamily('Arial');
    sh.getRange(r,6).setFormula(`=IF(A${r}="","",D${r}-E${r})`).setBackground(C.formulaGray).setNumberFormat(FMT_NUM).setFontFamily('Arial');
    sh.getRange(r,7).setFormula(`=IF(A${r}="","",IFERROR(VLOOKUP(A${r},'🛍️ Products'!A:H,8,0),""))`).setBackground(C.formulaGray).setNumberFormat(FMT_KRW).setFontFamily('Arial');
    sh.getRange(r,8).setFormula(`=IF(A${r}="","",IFERROR(VLOOKUP(A${r},'🛍️ Products'!A:I,9,0),""))`).setBackground(C.formulaGray).setNumberFormat(FMT_UZS).setFontFamily('Arial');
    sh.getRange(r,9).setValue(d[3]).setBackground(C.inputBlue).setFontColor('#1A5276').setNumberFormat(FMT_DATE).setFontFamily('Arial');
    sh.getRange(r,10).setValue(d[4]).setBackground(C.inputBlue).setFontColor('#1A5276').setNumberFormat(FMT_NUM).setFontFamily('Arial');
    sh.getRange(r,11).setFormula(`=IF(A${r}="","",IF(F${r}<=0,"🔴 Out of Stock",IF(F${r}<=J${r},"🟡 Low Stock","🟢 In Stock")))`).setBackground(C.formulaGray).setFontFamily('Arial').setFontSize(10);
  });
  for(let i=0;i<200;i++){
    const r=DS+invData.length+i; altRow(sh,r,11,i%2===1); sh.setRowHeight(r,20);
    sh.getRange(r,1).setBackground(i%2===1?C.rowAlt:C.white);
    sh.getRange(r,2).setFormula(`=IF(A${r}="","",IFERROR(VLOOKUP(A${r},'🛍️ Products'!A:C,2,0),"?"))`).setBackground(C.formulaGray).setFontFamily('Arial').setFontSize(10);
    sh.getRange(r,3).setFormula(`=IF(A${r}="","",IFERROR(VLOOKUP(A${r},'🛍️ Products'!A:E,5,0),""))`).setBackground(C.formulaGray).setFontFamily('Arial').setFontSize(10);
    [4,5].forEach(c=>sh.getRange(r,c).setBackground(C.inputBlue).setNumberFormat(FMT_NUM));
    sh.getRange(r,6).setFormula(`=IF(A${r}="","",D${r}-E${r})`).setBackground(C.formulaGray).setNumberFormat(FMT_NUM).setFontFamily('Arial');
    sh.getRange(r,7).setFormula(`=IF(A${r}="","",IFERROR(VLOOKUP(A${r},'🛍️ Products'!A:H,8,0),""))`).setBackground(C.formulaGray).setNumberFormat(FMT_KRW).setFontFamily('Arial');
    sh.getRange(r,8).setFormula(`=IF(A${r}="","",IFERROR(VLOOKUP(A${r},'🛍️ Products'!A:I,9,0),""))`).setBackground(C.formulaGray).setNumberFormat(FMT_UZS).setFontFamily('Arial');
    sh.getRange(r,9).setBackground(C.inputBlue).setNumberFormat(FMT_DATE);
    sh.getRange(r,10).setBackground(C.inputBlue).setNumberFormat(FMT_NUM);
    sh.getRange(r,11).setFormula(`=IF(A${r}="","",IF(F${r}<=0,"🔴 Out of Stock",IF(F${r}<=J${r},"🟡 Low Stock","🟢 In Stock")))`).setBackground(C.formulaGray).setFontFamily('Arial').setFontSize(10);
  }
  borders(sh,7,1,DS+invData.length+200-7,11);
  sh.setFrozenRows(7); sh.setFrozenColumns(1);
}

// ─────────────────────────────────────────────────────────────────────────────
// 💰 SALES
// ─────────────────────────────────────────────────────────────────────────────
function buildSales(ss){
  const sh=getOrCreate(ss,'💰 Sales','#1E8449');
  titleBlock(sh,'💰  Sales & Revenue Log',
    'Use menu "K-Beauty UZ ▾ → Log a Sale" to add orders  |  Total auto-calculates from product price × qty');
  [70,90,180,160,75,200,55,130,100,150,110,130,200].forEach((w,i)=>sh.setColumnWidth(i+1,w));
  sh.setRowHeight(4,24); sh.setRowHeight(5,32); sh.setRowHeight(6,6); sh.setRowHeight(7,36);
  [['Total Orders',`=COUNTA(A${ROW.sales}:A2007)`,'#1B3A6B'],
   ['Total Revenue',`=SUM(H${ROW.sales}:H2007)`,  '#1E8449'],
   ['This Month',   `=SUMIF(B${ROW.sales}:B2007,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),H${ROW.sales}:H2007)`,'#2E86C1'],
   ['Pending / Shipped',`=COUNTIFS(I${ROW.sales}:I2007,"<>Delivered",I${ROW.sales}:I2007,"<>Cancelled",I${ROW.sales}:I2007,"<>",A${ROW.sales}:A2007,"<>")`,'#C0392B']
  ].forEach(([lbl,f,bg],i)=>{
    sh.getRange(4,i*3+1,1,3).merge().setValue(lbl).setBackground(bg).setFontColor(C.white).setFontWeight('bold').setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sh.getRange(5,i*3+1,1,3).merge().setFormula(f).setNumberFormat(i===0?FMT_NUM:FMT_UZS).setBackground('#EBF5FB').setFontColor(bg).setFontWeight('bold').setFontFamily('Arial').setFontSize(i===0||i===3?16:13).setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
  headerRow(sh,7,['Order #','Date','Customer Name','Telegram / Contact','Product ID',
    'Product Name','Qty','Total (UZS)','Status','Shipping Method','Tracking #','Region (UZ)','Notes'],C.navy);
  const sampleSales=[
    ['ORD-001','05/01/2026','Dilnoza T.','@dilnoza_uz','P001',2,'Delivered','Uzbekistan Post','UZ123456','Tashkent',''],
    ['ORD-002','10/01/2026','Malika K.', '@malika_k', 'P002',1,'Delivered','Uzbekistan Post','UZ123457','Samarkand',''],
    ['ORD-003','15/01/2026','Nodira R.', '@nodira_r', 'P003',3,'Shipped',  'Uzbekistan Post','UZ123458','Tashkent',''],
    ['ORD-004','20/01/2026','Zulfiya A.','@zulfiya_a','P005',1,'Pending',  'Uzbekistan Post','','Fergana',''],
    ['ORD-005','25/01/2026','Barno S.',  '@barno_s',  'P004',2,'Delivered','Uzbekistan Post','UZ123460','Tashkent',''],
  ];
  const DS=ROW.sales;
  sampleSales.forEach((d,i)=>{
    const r=DS+i; altRow(sh,r,13,i%2===1); sh.setRowHeight(r,22);
    [d[0],d[1],d[2],d[3],d[4]].forEach((v,ci)=>{
      sh.getRange(r,ci+1).setValue(v).setFontFamily('Arial').setFontSize(10).setFontColor('#1A5276').setBackground(C.inputBlue);
      if(ci===1) sh.getRange(r,ci+1).setNumberFormat(FMT_DATE);
    });
    sh.getRange(r,6).setFormula(`=IF(E${r}="","",IFERROR(VLOOKUP(E${r},'🛍️ Products'!A:C,2,0),""))`).setBackground(C.formulaGray).setFontFamily('Arial').setFontSize(10);
    sh.getRange(r,7).setValue(d[5]).setBackground(C.inputBlue).setFontColor('#1A5276').setNumberFormat(FMT_NUM).setFontFamily('Arial');
    sh.getRange(r,8).setFormula(`=IF(E${r}="","",IFERROR(G${r}*VLOOKUP(E${r},'🛍️ Products'!A:I,9,0),0))`).setNumberFormat(FMT_UZS).setBackground(C.formulaGray).setFontFamily('Arial');
    [d[6],d[7],d[8],d[9],d[10]].forEach((v,ci)=>sh.getRange(r,9+ci).setValue(v).setFontFamily('Arial').setFontSize(10).setFontColor('#1A5276').setBackground(C.inputBlue));
  });
  for(let i=0;i<2000;i++){
    const r=DS+sampleSales.length+i; altRow(sh,r,13,i%2===1); sh.setRowHeight(r,20);
    [1,2,3,4,5,7,9,10,11,12,13].forEach(c=>sh.getRange(r,c).setBackground(i%2===1?C.rowAlt:C.white));
    sh.getRange(r,2).setBackground(i%2===1?C.rowAlt:C.white).setNumberFormat(FMT_DATE);
    sh.getRange(r,6).setFormula(`=IF(E${r}="","",IFERROR(VLOOKUP(E${r},'🛍️ Products'!A:C,2,0),""))`).setBackground(C.formulaGray).setFontFamily('Arial').setFontSize(10);
    sh.getRange(r,7).setBackground(i%2===1?C.rowAlt:C.white).setNumberFormat(FMT_NUM);
    sh.getRange(r,8).setFormula(`=IF(E${r}="","",IFERROR(G${r}*VLOOKUP(E${r},'🛍️ Products'!A:I,9,0),0))`).setNumberFormat(FMT_UZS).setBackground(C.formulaGray).setFontFamily('Arial');
  }
  borders(sh,7,1,DS+sampleSales.length+2000-7,13);
  sh.setFrozenRows(7); sh.setFrozenColumns(1);
}

// ─────────────────────────────────────────────────────────────────────────────
// 💸 EXPENSES
// ─────────────────────────────────────────────────────────────────────────────
function buildExpenses(ss){
  const sh=getOrCreate(ss,'💸 Expenses',C.red);
  titleBlock(sh,'💸  Expenses & Costs Log',
    'Use menu "K-Beauty UZ ▾ → Record an Expense" to add costs');
  [70,100,165,340,200,110,150,90,140,80,250].forEach((w,i)=>sh.setColumnWidth(i+1,w));
  sh.setRowHeight(4,24); sh.setRowHeight(5,32); sh.setRowHeight(6,6); sh.setRowHeight(7,36);
  [['Total Expenses (USD)',`=SUM(F${ROW.exp}:F1007)`,'#C0392B'],
   ['Product Costs',`=SUMIF(C${ROW.exp}:C1007,"Product Purchase",F${ROW.exp}:F1007)`,'#943126'],
   ['Shipping',`=SUMIF(C${ROW.exp}:C1007,"Shipping / Freight",F${ROW.exp}:F1007)`,'#E74C3C'],
   ['Salaries',`=SUMIF(C${ROW.exp}:C1007,"Posting Guy Salary",F${ROW.exp}:F1007)`,'#922B21'],
  ].forEach(([lbl,f,bg],i)=>{
    sh.getRange(4,i*3+1,1,3).merge().setValue(lbl).setBackground(bg).setFontColor(C.white).setFontWeight('bold').setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sh.getRange(5,i*3+1,1,3).merge().setFormula(f).setNumberFormat(FMT_USD).setBackground('#FDEDEC').setFontColor(bg).setFontWeight('bold').setFontFamily('Arial').setFontSize(14).setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
  headerRow(sh,7,['Expense ID','Date','Category','Description','Paid To / Vendor',
    'Amount (USD)','Payment Method','Receipt #','Paid By','Month','Notes'],C.navy);
  const sampleExp=[
    ['EXP-001','02/01/2026','Product Purchase',   'Round Lab Toner × 50 units',           'Olive Young KR',   452.00,'Bank Transfer','RC-001','Partner 1 (OJ)'],
    ['EXP-002','02/01/2026','Product Purchase',   'Beauty of Joseon Serum × 40 units',    'Olive Young KR',   576.00,'Bank Transfer','RC-002','Partner 1 (OJ)'],
    ['EXP-003','02/01/2026','Product Purchase',   'COSRX Essence × 30 units',             'Gmarket KR',       483.00,'Bank Transfer','RC-003','Partner 1 (OJ)'],
    ['EXP-004','03/01/2026','Shipping / Freight', 'Air freight KR → UZ (January batch)',  'Korea Post / EMS', 320.00,'Bank Transfer','SH-001','Partner 1 (OJ)'],
    ['EXP-005','05/01/2026','Posting Guy Salary', 'January salary – Posting Guy UZ',      'Posting Guy UZ',   150.00,'UZ Bank Wire', 'PAY-01','Partner 2'],
    ['EXP-006','05/01/2026','Platform Fee',       'Instagram / Telegram boost',            'Meta / Telegram',  30.00,'Card',         'PF-001','Partner 3'],
    ['EXP-007','06/01/2026','Packaging',          'Branded bags & bubble wrap',            'Local UZ supplier',45.00,'Cash',         'PK-001','Posting Guy'],
    ['EXP-008','07/01/2026','Customs / Duties',   'UZ customs clearance – Jan batch',     'UZ Customs',        80.00,'Bank Transfer','CD-001','Partner 4'],
  ];
  const DS=ROW.exp;
  sampleExp.forEach((d,i)=>{
    const r=DS+i; altRow(sh,r,11,i%2===1); sh.setRowHeight(r,22);
    d.slice(0,9).forEach((v,ci)=>{
      const c=sh.getRange(r,ci+1);
      c.setValue(v).setFontFamily('Arial').setFontSize(10).setFontColor('#1A5276').setBackground(C.inputBlue);
      if(ci===1) c.setNumberFormat(FMT_DATE);
      if(ci===5) c.setNumberFormat(FMT_USD);
    });
    sh.getRange(r,10).setFormula(`=IF(B${r}="","",TEXT(B${r},"YYYY-MM"))`).setBackground(C.formulaGray).setFontFamily('Arial').setFontSize(10);
    sh.getRange(r,11).setBackground(C.inputBlue).setFontFamily('Arial').setFontSize(10);
  });
  for(let i=0;i<1000;i++){
    const r=DS+sampleExp.length+i; altRow(sh,r,11,i%2===1); sh.setRowHeight(r,20);
    [1,2,3,4,5,6,7,8,9,11].forEach(c=>sh.getRange(r,c).setBackground(i%2===1?C.rowAlt:C.white));
    sh.getRange(r,2).setBackground(i%2===1?C.rowAlt:C.white).setNumberFormat(FMT_DATE);
    sh.getRange(r,6).setBackground(i%2===1?C.rowAlt:C.white).setNumberFormat(FMT_USD);
    sh.getRange(r,10).setFormula(`=IF(B${r}="","",TEXT(B${r},"YYYY-MM"))`).setBackground(C.formulaGray).setFontFamily('Arial').setFontSize(10);
  }
  borders(sh,7,1,DS+sampleExp.length+1000-7,11);
  sh.setFrozenRows(7); sh.setFrozenColumns(1);
}

// ─────────────────────────────────────────────────────────────────────────────
// 🤝 PARTNER DASHBOARD
// ─────────────────────────────────────────────────────────────────────────────
function buildPartnerDashboard(ss){
  const sh=getOrCreate(ss,'🤝 Partners',C.gold);
  titleBlock(sh,'🤝  Partner Dashboard',
    'P&L summary and profit split  |  All values pull live — no editing needed');
  [240,140,90,160,160,140].forEach((w,i)=>sh.setColumnWidth(i+1,w));
  sectionHeader(sh,4,1,6,'  📊  PROFIT & LOSS SUMMARY',C.navy);
  sh.setRowHeight(5,8);
  const pl=[
    ['Total Revenue (UZS)',          `=SUM('💰 Sales'!H${ROW.sales}:H2007)`,FMT_UZS,true,false],
    ['Cost of Goods Sold (UZS)',      `=SUMPRODUCT(('📦 Inventory'!A${ROW.inv}:A208<>"")*'📦 Inventory'!E${ROW.inv}:E208*IFERROR(VLOOKUP('📦 Inventory'!A${ROW.inv}:A208,'🛍️ Products'!A:J,10,0),0))`,FMT_UZS,false,false],
    ['GROSS PROFIT (UZS)',            '=B6-B7',FMT_UZS,true,false],
    ['Gross Margin %',                '=IFERROR(B8/B6,0)',FMT_PCT,false,false],
    ['','','',false,false],
    ['Total Expenses (USD)',          `=SUM('💸 Expenses'!F${ROW.exp}:F1007)`,FMT_USD,false,false],
    ['  – Product Purchases',         `=SUMIF('💸 Expenses'!C${ROW.exp}:C1007,"Product Purchase",'💸 Expenses'!F${ROW.exp}:F1007)`,FMT_USD,false,true],
    ['  – Shipping & Freight',        `=SUMIF('💸 Expenses'!C${ROW.exp}:C1007,"Shipping / Freight",'💸 Expenses'!F${ROW.exp}:F1007)`,FMT_USD,false,true],
    ['  – Posting Guy Salary',        `=SUMIF('💸 Expenses'!C${ROW.exp}:C1007,"Posting Guy Salary",'💸 Expenses'!F${ROW.exp}:F1007)`,FMT_USD,false,true],
    ['  – Platform Fees',             `=SUMIF('💸 Expenses'!C${ROW.exp}:C1007,"Platform Fee",'💸 Expenses'!F${ROW.exp}:F1007)`,FMT_USD,false,true],
    ['  – Other',                     `=SUMIFS('💸 Expenses'!F${ROW.exp}:F1007,'💸 Expenses'!C${ROW.exp}:C1007,"<>Product Purchase",'💸 Expenses'!C${ROW.exp}:C1007,"<>Shipping / Freight",'💸 Expenses'!C${ROW.exp}:C1007,"<>Posting Guy Salary",'💸 Expenses'!C${ROW.exp}:C1007,"<>Platform Fee")`,FMT_USD,false,true],
    ['','','',false,false],
    ['NET PROFIT (USD)',              '=B8/USD_UZS_RATE-B11',FMT_USD,true,false],
    ['Net Margin %',                  '=IFERROR(B19/B6*USD_UZS_RATE,0)',FMT_PCT,false,false],
  ];
  pl.forEach(([lbl,f,fmt,bold,indent],i)=>{
    const r=6+i; sh.setRowHeight(r,lbl===''?8:22);
    if(lbl==='') return;
    sh.getRange(r,1).setValue(lbl).setFontFamily('Arial').setFontSize(10)
      .setFontWeight(bold?'bold':'normal').setFontColor(indent?'#666666':'#1A1A2E')
      .setBackground(bold?C.lightBlue:C.white).setVerticalAlignment('middle');
    if(f) sh.getRange(r,2).setFormula(f).setNumberFormat(fmt).setFontFamily('Arial').setFontSize(10)
      .setFontWeight(bold?'bold':'normal').setBackground(bold?C.lightBlue:C.formulaGray)
      .setHorizontalAlignment('right');
  });
  // Partner split
  sectionHeader(sh,24,1,6,'  👥  PARTNER PROFIT SPLIT  (yellow cells = your inputs)',C.gold,'#1A1A2E');
  sh.setRowHeight(25,36);
  headerRow(sh,25,['Partner','Role','Share %','Net Profit Share (USD)','Capital Invested (USD)','ROI (%)'],C.navy);
  [['Partner 1 (OJ)','Founder / Overseas Procurement',0.25],
   ['Partner 2','Operations',0.25],
   ['Partner 3','Marketing',0.25],
   ['Partner 4','Finance / Admin',0.25],
  ].forEach(([name,role,share],i)=>{
    const r=26+i; altRow(sh,r,6,i%2===1); sh.setRowHeight(r,22);
    sh.getRange(r,1).setValue(name).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial').setFontSize(10);
    sh.getRange(r,2).setValue(role).setBackground(C.inputBlue).setFontColor('#1A5276').setFontFamily('Arial').setFontSize(10);
    sh.getRange(r,3).setValue(share).setNumberFormat(FMT_PCT).setBackground(C.yellow).setFontColor('#1A5276').setFontFamily('Arial').setFontSize(10).setNote('Change this % if ownership differs');
    sh.getRange(r,4).setFormula(`=B19*C${r}`).setNumberFormat(FMT_USD).setBackground(C.formulaGray).setFontFamily('Arial');
    sh.getRange(r,5).setValue(0).setNumberFormat(FMT_USD).setBackground(C.yellow).setFontColor('#1A5276').setFontFamily('Arial').setNote('Enter capital invested by this partner');
    sh.getRange(r,6).setFormula(`=IFERROR(D${r}/E${r},0)`).setNumberFormat(FMT_PCT).setBackground(C.formulaGray).setFontFamily('Arial');
  });
  borders(sh,4,1,32,6);
  sh.setFrozenRows(3);
}

// ─────────────────────────────────────────────────────────────────────────────
// 📊 DASHBOARD (built last)
// ─────────────────────────────────────────────────────────────────────────────
function buildDashboard(ss){
  const sh=getOrCreate(ss,'📊 Dashboard',C.navy);
  titleBlock(sh,'📊  K-Beauty UZ — Business Dashboard',
    'Live overview — all numbers update automatically  |  Use "K-Beauty UZ ▾" menu to add data');
  for(let i=1;i<=10;i++) sh.setColumnWidth(i,168);
  sh.setColumnWidth(1,200);
  sh.setRowHeight(3,10);

  // KPI row
  [
    {lbl:'💰 Total Revenue',  f:`=SUM('💰 Sales'!H${ROW.sales}:H2007)`,      fmt:FMT_UZS, bg:'#1E8449'},
    {lbl:'📦 Units Sold',     f:`=SUM('📦 Inventory'!E${ROW.inv}:E208)`,     fmt:FMT_NUM, bg:'#2E86C1'},
    {lbl:'📈 Gross Profit',   f:`='🤝 Partners'!B8`,                          fmt:FMT_UZS, bg:'#1A5276'},
    {lbl:'💸 Total Expenses', f:`=SUM('💸 Expenses'!F${ROW.exp}:F1007)`,     fmt:FMT_USD, bg:C.red},
    {lbl:'🏆 Net Profit',     f:`='🤝 Partners'!B19`,                         fmt:FMT_USD, bg:'#117A65'},
    {lbl:'📊 Gross Margin',   f:`='🤝 Partners'!B9`,                          fmt:FMT_PCT, bg:C.gold},
  ].forEach((kpi,i)=>{
    sh.setRowHeight(4,22); sh.setRowHeight(5,52);
    sh.getRange(4,i+1).setValue(kpi.lbl).setBackground(kpi.bg).setFontColor(C.white).setFontWeight('bold').setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
    sh.getRange(5,i+1).setFormula(kpi.f).setNumberFormat(kpi.fmt).setBackground('#F0F3F4').setFontColor(kpi.bg).setFontWeight('bold').setFontFamily('Arial').setFontSize(15).setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
  sh.setRowHeight(6,10);

  // Top Products
  sectionHeader(sh,7,1,6,'  🛍️  PRODUCT CATALOGUE SNAPSHOT',C.navy);
  sh.setRowHeight(8,36);
  headerRow(sh,8,['Product ID','Product Name','Category','Sell Price (UZS)','Real Cost (UZS)','Profit / Unit (UZS)'],C.navy);
  for(let i=0;i<8;i++){
    const r=9+i, srcR=ROW.prod+i;
    altRow(sh,r,6,i%2===1); sh.setRowHeight(r,22);
    ['A','C','E','I','J','K'].forEach((col,ci)=>{
      sh.getRange(r,ci+1).setFormula(`='🛍️ Products'!${col}${srcR}`).setBackground(C.formulaGray).setFontFamily('Arial').setFontSize(10);
      if(ci>=3) sh.getRange(r,ci+1).setNumberFormat(FMT_UZS);
    });
  }
  sh.setRowHeight(18,10);

  // Recent Sales summary
  sectionHeader(sh,19,1,6,'  💰  SALES SUMMARY',C.navy);
  sh.setRowHeight(20,22);
  [['Total Orders',`=COUNTA('💰 Sales'!A${ROW.sales}:A2007)`,FMT_NUM],
   ['Total Revenue (UZS)',`=SUM('💰 Sales'!H${ROW.sales}:H2007)`,FMT_UZS],
   ['This Month (UZS)',`=SUMIF('💰 Sales'!B${ROW.sales}:B2007,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),'💰 Sales'!H${ROW.sales}:H2007)`,FMT_UZS],
  ].forEach(([lbl,f,fmt],i)=>{
    sh.getRange(20,i*2+1).setValue(lbl).setFontWeight('bold').setFontFamily('Arial').setFontSize(10).setBackground(C.lightBlue);
    sh.getRange(20,i*2+2).setFormula(f).setNumberFormat(fmt).setBackground(C.formulaGray).setFontFamily('Arial').setFontSize(10).setHorizontalAlignment('right');
  });

  sh.setRowHeight(21,10);

  // How-to
  sectionHeader(sh,22,1,6,'  ℹ️  HOW TO USE',  '#2C4F85',C.white);
  ['① Open the  "K-Beauty UZ ▾"  menu at the top of the screen.',
   '② Add New Product  →  fills Products sheet. Real cost & profit auto-calculate.',
   '③ Log a Sale  →  fills Sales sheet. Product name & total price auto-fill from Product ID.',
   '④ Record an Expense  →  fills Expenses sheet. Month column auto-fills from date.',
   '⑤ Use 🔍 Search  →  finds anything across Products, Sales, and Expenses instantly.',
   '⑥ Update exchange rate in ⚙️ Settings  →  all prices recalculate across every sheet.',
   '⑦ This Dashboard is read-only — it updates by itself, no changes needed here.',
  ].forEach((step,i)=>{
    const r=23+i; sh.setRowHeight(r,24);
    sh.getRange(r,1,1,6).merge().setValue(step)
      .setBackground(i%2===0?'#EBF5FB':'#F8F9FA')
      .setFontFamily('Arial').setFontSize(10).setFontColor('#1A1A2E')
      .setVerticalAlignment('middle').setHorizontalAlignment('left');
  });
  borders(sh,4,1,29,6);
  sh.setFrozenRows(3);
}
