/* script.js — Spice Vault Traders
  - Export: Product column = names only
  - Quantity column = "value UNIT" semicolon-separated (e.g., "2 KG; 500 G")
  - Improved UI interactions preserved
*/

let inventory = JSON.parse(localStorage.getItem('spiceVaultInventory')) || [];
let bills = JSON.parse(localStorage.getItem('spiceVaultBills')) || [];
let billCounter = bills.length > 0 ? Math.max(...bills.map(b => b.id)) + 1 : 1001;
let currentPaymentBill = null;

/* -------- DOM -------- */
const navLinks = document.querySelectorAll('.nav-link');
const sections = document.querySelectorAll('.section');
const addProductBtn = document.getElementById('add-product-btn');
const productForm = document.getElementById('product-form');
const cancelProductBtn = document.getElementById('cancel-product');
const saveProductBtn = document.getElementById('save-product');
const inventoryBody = document.getElementById('inventory-body');
const addProductBillBtn = document.getElementById('add-product-bill');
const billItemsBody = document.getElementById('bill-items-body');
const generateBillBtn = document.getElementById('generate-bill');
const amountPaidInput = document.getElementById('amount-paid');
const paymentsTableBody = document.querySelector('#payments-table tbody');
const billsTableBody = document.querySelector('#bills-table tbody');
const recentTransactionsBody = document.querySelector('#recent-transactions tbody');
const viewAllBillsBtn = document.getElementById('view-all-bills');
const refreshBillsBtn = document.getElementById('refresh-bills');
const paymentModal = document.getElementById('payment-modal');
const closeModalBtns = document.querySelectorAll('.close-modal');
const cancelPaymentBtn = document.getElementById('cancel-payment');
const savePaymentBtn = document.getElementById('save-payment');

/* Export DOM */
const exportBillsBtn = document.getElementById('export-bills');
const exportModal = document.getElementById('export-modal');
const closeExportModalBtn = document.getElementById('close-export-modal');
const cancelExportBtn = document.getElementById('cancel-export');
const downloadMonthExportBtn = document.getElementById('download-month-export');
const exportMonthSelect = document.getElementById('export-month');
const exportYearSelect = document.getElementById('export-year');
const exportCurrentMonthBtn = document.getElementById('export-current-month');

const toastEl = document.getElementById('toast');
document.getElementById('current-year').textContent = new Date().getFullYear();

/* -------- Init -------- */
document.addEventListener('DOMContentLoaded', () => {
  const billDateEl = document.getElementById('bill-date');
  if (billDateEl) billDateEl.valueAsDate = new Date();

  renderInventory();
  renderPayments();
  renderBills();
  updateDashboard();
  populateMonthYearSelects();
});

/* -------- Navigation -------- */
navLinks.forEach(link => {
  link.addEventListener('click', (e) => {
   e.preventDefault();
   navLinks.forEach(l => l.classList.remove('active'));
   link.classList.add('active');
   const target = link.dataset.target;
   sections.forEach(s => s.classList.remove('active'));
   const el = document.getElementById(target);
   if (el) el.classList.add('active');
  });
});

/* -------- Inventory handlers -------- */
addProductBtn?.addEventListener('click', () => {
  productForm.style.display = 'block';
  addProductBtn.style.display = 'none';
});
cancelProductBtn?.addEventListener('click', () => {
  productForm.style.display = 'none';
  addProductBtn.style.display = 'inline-flex';
  resetProductForm();
});
saveProductBtn?.addEventListener('click', saveProduct);

function resetProductForm() {
  ['product-name','product-quantity','product-unit','buying-rate','selling-rate','product-description']
   .forEach(id => { const el = document.getElementById(id); if (el) el.value = ''; });
  document.getElementById('product-unit').value = 'KG';
}

function saveProduct() {
  const name = document.getElementById('product-name').value.trim();
  const quantity = parseFloat(document.getElementById('product-quantity').value);
  const unit = document.getElementById('product-unit').value;
  const buyingRate = parseFloat(document.getElementById('buying-rate').value);
  const sellingRate = parseFloat(document.getElementById('selling-rate').value) || (buyingRate * 1.2);
  const description = document.getElementById('product-description').value.trim();

  if (!name || isNaN(quantity) || isNaN(buyingRate)) { showToast('Please fill required fields', 'error'); return; }

  const product = { id: Date.now(), name, quantity, unit, buyingRate, sellingRate, description, createdAt: new Date().toISOString() };
  inventory.push(product);
  localStorage.setItem('spiceVaultInventory', JSON.stringify(inventory));
  showToast('Product added');
  resetProductForm();
  productForm.style.display = 'none';
  addProductBtn.style.display = 'inline-flex';
  renderInventory();
  updateDashboard();
}

function renderInventory() {
  inventoryBody.innerHTML = '';
  if (inventory.length === 0) {
   inventoryBody.innerHTML = `<tr><td colspan="7" class="empty-state"><i class="fas fa-box-open"></i><p>No products in inventory.</p></td></tr>`;
   return;
  }
  inventory.forEach(p => {
   const tr = document.createElement('tr');
   tr.innerHTML = `
    <td>SVT-${p.id.toString().slice(-4)}</td>
    <td>${p.name}</td>
    <td>${p.description || '-'}</td>
    <td>${p.quantity} ${p.unit}</td>
    <td>₹${p.buyingRate.toFixed(2)}</td>
    <td>₹${p.sellingRate.toFixed(2)}</td>
    <td><button class="action-btn delete-btn" data-id="${p.id}"><i class="fas fa-trash"></i></button></td>
   `;
   inventoryBody.appendChild(tr);
  });

  document.querySelectorAll('.delete-btn').forEach(btn => btn.addEventListener('click', (e) => {
   const id = parseInt(e.currentTarget.dataset.id); if (confirm('Delete product?')) deleteProduct(id);
  }));
}

function deleteProduct(id) {
  inventory = inventory.filter(p => p.id !== id);
  localStorage.setItem('spiceVaultInventory', JSON.stringify(inventory));
  renderInventory(); updateDashboard(); showToast('Product deleted');
}

/* -------- Billing -------- */
addProductBillBtn?.addEventListener('click', addProductToBill);
generateBillBtn?.addEventListener('click', generateBill);
amountPaidInput?.addEventListener('input', updateBillSummary);

function addProductToBill() {
  if (inventory.length === 0) { showToast('Add products first', 'error'); return; }
  const row = document.createElement('tr');
  row.innerHTML = `
   <td><select class="bill-product-select"><option value="">Select Product</option>
    ${inventory.map(p => `<option value="${p.id}" data-rate="${p.sellingRate}" data-unit="${p.unit}" data-qty="${p.quantity}" data-desc="${p.description||''}">${p.name}</option>`).join('')}
   </select></td>
   <td class="bill-description">-</td>
   <td class="bill-available">-</td>
   <td><input type="number" class="bill-quantity" value="1" step="0.01"></td>
   <td class="bill-unit">-</td>
   <td><input type="number" class="bill-rate" step="0.01"></td>
   <td class="bill-total">₹0.00</td>
   <td><button class="action-btn delete-btn"><i class="fas fa-trash"></i></button></td>
  `;
  billItemsBody.appendChild(row);

  const sel = row.querySelector('.bill-product-select');
  const qty = row.querySelector('.bill-quantity');
  const rate = row.querySelector('.bill-rate');
  const del = row.querySelector('.delete-btn');

  sel.addEventListener('change', () => {
   const opt = sel.options[sel.selectedIndex];
   const unit = opt.dataset.unit || '-';
   const desc = opt.dataset.desc || '-';
   const available = opt.dataset.qty || '-';
   const r = parseFloat(opt.dataset.rate || 0);
   row.querySelector('.bill-description').textContent = desc || '-';
   row.querySelector('.bill-available').textContent = `${available} ${unit}`;
   row.querySelector('.bill-unit').textContent = unit;
   rate.value = r.toFixed(2);
   calculateBillTotal(row);
   updateBillSummary();
  });

  qty.addEventListener('input', () => { calculateBillTotal(row); updateBillSummary(); });
  rate.addEventListener('input', () => { calculateBillTotal(row); updateBillSummary(); });
  del.addEventListener('click', () => { row.remove(); updateBillSummary(); });
}

function calculateBillTotal(row) {
  const q = parseFloat(row.querySelector('.bill-quantity').value) || 0;
  const r = parseFloat(row.querySelector('.bill-rate').value) || 0;
  const total = q * r;
  row.querySelector('.bill-total').textContent = `₹${total.toFixed(2)}`;
}

function updateBillSummary() {
  let subtotal = 0;
  document.querySelectorAll('#bill-items-body tr').forEach(row => {
   const t = parseFloat((row.querySelector('.bill-total').textContent || '₹0').replace('₹','')) || 0;
   subtotal += t;
  });
  const paid = parseFloat(amountPaidInput.value) || 0;
  const pending = subtotal - paid;
  document.getElementById('subtotal').textContent = `₹${subtotal.toFixed(2)}`;
  document.getElementById('total-amount').textContent = `₹${subtotal.toFixed(2)}`;
  document.getElementById('pending-amount').textContent = `₹${pending.toFixed(2)}`;
}

function generateBill() {
  const customer = document.getElementById('customer-name').value.trim();
  const date = document.getElementById('bill-date').value;
  const paid = parseFloat(amountPaidInput.value) || 0;
  const total = parseFloat(document.getElementById('total-amount').textContent.replace('₹','')) || 0;
  if (!customer) { showToast('Enter company name', 'error'); return; }

  const items = [];
  let valid = true;

  document.querySelectorAll('#bill-items-body tr').forEach(row => {
   const pid = parseInt(row.querySelector('.bill-product-select').value);
   const product = inventory.find(p => p.id === pid);
   if (!product) { valid = false; return; }
   const quantity = parseFloat(row.querySelector('.bill-quantity').value) || 0;
   const rate = parseFloat(row.querySelector('.bill-rate').value) || 0;
   if (quantity > product.quantity) { showToast(`Not enough ${product.name}`, 'error'); valid = false; return; }
   items.push({ productId: product.id, name: product.name, unit: product.unit, quantity, rate, total: quantity * rate });
  });

  if (items.length === 0) { showToast('Add products to bill', 'error'); return; }
  if (!valid) return;

  const bill = {
   id: billCounter++,
   customer,
   date,
   items,
   total,
   paid,
   pending: total - paid,
   status: paid === 0 ? 'Pending' : (paid < total ? 'Partial' : 'Paid')
  };

  bills.push(bill);
  localStorage.setItem('spiceVaultBills', JSON.stringify(bills));

  // decrement inventory
  items.forEach(it => {
   const prod = inventory.find(p => p.id === it.productId);
   if (prod) prod.quantity -= it.quantity;
  });
  localStorage.setItem('spiceVaultInventory', JSON.stringify(inventory));

  showToast('Bill created');
  billItemsBody.innerHTML = '';
  document.getElementById('customer-name').value = '';
  amountPaidInput.value = '0';
  renderInventory(); renderPayments(); renderBills(); updateDashboard();
}

/* -------- Payments -------- */
function renderPayments() {
  paymentsTableBody.innerHTML = '';
  if (bills.length === 0) {
   paymentsTableBody.innerHTML = `<tr><td colspan="9" class="empty-state"><p>No payments</p></td></tr>`;
   return;
  }
  bills.forEach(b => {
   const due = new Date(b.date); due.setDate(due.getDate() + 30);
   const tr = document.createElement('tr');
   tr.innerHTML = `
    <td>SVT-${b.id}</td>
    <td>${b.customer}</td>
    <td>${b.date}</td>
    <td>₹${b.total.toFixed(2)}</td>
    <td>₹${b.paid.toFixed(2)}</td>
    <td>₹${b.pending.toFixed(2)}</td>
    <td>${due.toISOString().split('T')[0]}</td>
    <td>${b.status}</td>
    <td>${b.pending > 0 ? `<button class="action-btn payment-btn" data-id="${b.id}"><i class="fas fa-rupee-sign"></i></button>` : '<span class="badge">Paid</span>'}</td>
   `;
   paymentsTableBody.appendChild(tr);
  });

  document.querySelectorAll('.payment-btn').forEach(btn => btn.addEventListener('click', (e) => {
   openPaymentModal(parseInt(e.currentTarget.dataset.id));
  }));
}

/* -------- Bills UI -------- */
function renderBills() {
  billsTableBody.innerHTML = '';
  if (bills.length === 0) {
   billsTableBody.innerHTML = `<tr><td colspan="7" class="empty-state"><p>No bills yet</p></td></tr>`;
   return;
  }
  bills.forEach(b => {
   const tr = document.createElement('tr');
   tr.innerHTML = `<td>SVT-${b.id}</td><td>${b.date}</td><td>${b.customer}</td><td>₹${b.total.toFixed(2)}</td><td>₹${b.paid.toFixed(2)}</td><td>₹${b.pending.toFixed(2)}</td><td>${b.status}</td>`;
   billsTableBody.appendChild(tr);
  });
}

/* -------- Dashboard -------- */
function updateDashboard() {
  document.getElementById('total-products').textContent = inventory.length;
  const now = new Date();
  const curMonth = now.toISOString().slice(0,7);
  const monthlySales = bills.filter(b => b.date && b.date.startsWith(curMonth)).reduce((s,b) => s + b.total, 0);
  document.getElementById('monthly-sales').textContent = monthlySales.toLocaleString('en-IN');
  const pendingTotal = bills.reduce((s,b) => s + b.pending, 0);
  document.getElementById('pending-payments').textContent = pendingTotal.toLocaleString('en-IN');

  recentTransactionsBody.innerHTML = '';
  const recent = [...bills].sort((a,b) => new Date(b.date) - new Date(a.date)).slice(0,5);
  if (recent.length === 0) {
   recentTransactionsBody.innerHTML = `<tr><td colspan="5" class="empty-state"><p>No recent transactions</p></td></tr>`;
   return;
  }
  recent.forEach(b => {
   const tr = document.createElement('tr');
   tr.innerHTML = `<td>SVT-${b.id}</td><td>${b.customer}</td><td>${b.date}</td><td>₹${b.total.toFixed(2)}</td><td>${b.status}</td>`;
   recentTransactionsBody.appendChild(tr);
  });
}

/* -------- Payment modal -------- */
function openPaymentModal(billId) {
  const bill = bills.find(b => b.id === billId); if (!bill) return;
  currentPaymentBill = bill;
  document.getElementById('payment-bill-no').value = `SVT-${bill.id}`;
  document.getElementById('payment-customer').value = bill.customer;
  document.getElementById('payment-total').value = bill.total.toFixed(2);
  document.getElementById('payment-paid').value = bill.paid.toFixed(2);
  document.getElementById('payment-pending').value = bill.pending.toFixed(2);
  document.getElementById('payment-amount').value = '';
  document.getElementById('payment-date').valueAsDate = new Date();
  document.getElementById('payment-notes').value = '';
  paymentModal.style.display = 'flex';
}
closeModalBtns.forEach(b => b.addEventListener('click', () => document.querySelectorAll('.modal').forEach(m => m.style.display = 'none')));
cancelPaymentBtn?.addEventListener('click', () => document.querySelectorAll('.modal').forEach(m => m.style.display = 'none'));
savePaymentBtn?.addEventListener('click', () => {
  const amt = parseFloat(document.getElementById('payment-amount').value) || 0;
  const dt = document.getElementById('payment-date').value;
  if (!amt || amt <= 0 || !dt) { showToast('Enter valid amount & date', 'error'); return; }
  if (amt > currentPaymentBill.pending) { showToast('Amount exceeds pending', 'error'); return; }
  currentPaymentBill.paid += amt;
  currentPaymentBill.pending = currentPaymentBill.total - currentPaymentBill.paid;
  currentPaymentBill.status = currentPaymentBill.pending === 0 ? 'Paid' : 'Partial';
  localStorage.setItem('spiceVaultBills', JSON.stringify(bills));
  document.querySelectorAll('.modal').forEach(m => m.style.display = 'none');
  renderPayments(); renderBills(); updateDashboard(); showToast('Payment recorded');
});

/* -------- Toast -------- */
function showToast(msg, type='success') {
  toastEl.textContent = msg;
  toastEl.className = 'toast show';
  if (type === 'error') toastEl.style.background = '#ef4444';
  else toastEl.style.background = 'linear-gradient(90deg,#3b82f6,#60a5fa)';
  setTimeout(() => toastEl.className = 'toast', 3000);
}

/* -------- Export modal wiring -------- */
exportBillsBtn?.addEventListener('click', () => { openExportModal(); });
closeExportModalBtn?.addEventListener('click', () => { closeExportModal(); });
cancelExportBtn?.addEventListener('click', () => { closeExportModal(); });
downloadMonthExportBtn?.addEventListener('click', () => { exportSelectedMonth(); });
exportCurrentMonthBtn?.addEventListener('click', () => { generateCurrentMonthExport(); });

function openExportModal() { populateMonthYearSelects(); exportModal.style.display = 'flex'; }
function closeExportModal() { exportModal.style.display = 'none'; }

function populateMonthYearSelects() {
  const months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  exportMonthSelect.innerHTML = months.map((m,i) => `<option value="${i+1}">${m}</option>`).join('');
  const currentYear = new Date().getFullYear();
  let minYear = currentYear - 3;
  if (bills.length > 0) { const yrs = bills.map(b => new Date(b.date).getFullYear()); minYear = Math.min(...yrs, minYear); }
  exportYearSelect.innerHTML = '';
  for (let y = currentYear; y >= minYear; y--) exportYearSelect.innerHTML += `<option value="${y}">${y}</option>`;
  const now = new Date(); exportMonthSelect.value = now.getMonth() + 1; exportYearSelect.value = now.getFullYear();
}

function exportSelectedMonth() {
  const month = parseInt(exportMonthSelect.value, 10);
  const year = parseInt(exportYearSelect.value, 10);
  exportModal.style.display = 'none';
  exportMonthlyData(month, year);
}

function generateCurrentMonthExport() {
  const now = new Date();
  exportMonthlyData(now.getMonth() + 1, now.getFullYear());
}

/* -------- Export implementation (final)
  - Title (merged A1:G1): (Month / Year) - Spice Vault Traders
  - Header row A2:G2: S.no | Product | Customer | Quantity | Date | Total | Pending
  - Product column: names only (semicolon separated)
  - Quantity column: "value UNIT" semicolon separated (e.g., "2 KG; 500 G")
  - Numeric types for S.no, Total, Pending
*/
function exportMonthlyData(month, year) {
  if (!bills || bills.length === 0) { showToast('No bills to export', 'error'); return; }

  const filtered = bills.filter(b => {
   if (!b.date) return false;
   const d = new Date(b.date + 'T00:00:00');
   return (d.getMonth() + 1) === month && d.getFullYear() === year;
  });

  if (filtered.length === 0) { showToast('No bills for selected month', 'error'); return; }

  const monthName = getMonthName(month);
  const title = `(${monthName} / ${year}) - Spice Vault Traders`;

  const aoa = [];
  aoa.push([title, '', '', '', '', '', '']); // Row 1 title (merge A1:G1)
  aoa.push(['S.no', 'Product', 'Customer', 'Quantity', 'Date', 'Total', 'Pending']); // Row 2 header

  filtered.forEach((b, idx) => {
   // Product names only (semicolon-separated)
   const productNames = b.items.map(it => it.name).join('; ');
   // Quantity column: value + unit per item, semicolon-separated
   const qtyStr = b.items.map(it => `${Number(it.quantity)} ${it.unit}`).join('; ');
   // Ensure numeric totals are numbers
   const totalNum = (typeof b.total === 'number') ? b.total : Number(b.total || 0);
   const pendingNum = (typeof b.pending === 'number') ? b.pending : Number(b.pending || 0);

   aoa.push([ idx + 1, productNames, b.customer, qtyStr, b.date, totalNum, pendingNum ]);
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);

  // Merge title A1:G1
  ws['!merges'] = ws['!merges'] || [];
  ws['!merges'].push({ s: { r: 0, c: 0 }, e: { r: 0, c: 6 } });

  // Column widths
  ws['!cols'] = [
   { wch: 6 },   // S.no
   { wch: 40 },  // Product
   { wch: 28 },  // Customer
   { wch: 20 },  // Quantity (value + unit)
   { wch: 14 },  // Date
   { wch: 14 },  // Total
   { wch: 14 }   // Pending
  ];

  // Make sure the range is set
  const range = XLSX.utils.decode_range(ws['!ref']);
  ws['!ref'] = XLSX.utils.encode_range(range.s, range.e);

  // Optionally bold header row and title (some viewers preserve cell.s; many don't — but harmless)
  try {
   // Title style
   ws['A1'].s = { font: { name: "Calibri", sz: 14, bold: true, color: { rgb: "FF0B57A3" } }, alignment: { horizontal: "center", vertical: "center" } };
   // Header styles (A2:G2)
   const headerCols = ['A','B','C','D','E','F','G'];
   headerCols.forEach((col) => {
    const cell = ws[`${col}2`];
    if (cell) cell.s = { font: { bold: true }, fill: { fgColor: { rgb: "FFDDEEF9" } }, alignment: { horizontal: "center" } };
   });
  } catch(e){
   // styling is optional — ignore if not supported
   console.debug('Cell styling not applied', e);
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Export');

  const filename = `(${monthName} / ${year}) - Spice Vault Traders.xlsx`;

  try {
   XLSX.writeFile(wb, filename);
   showToast(`Exported ${filtered.length} bill(s): ${filename}`, 'success');
  } catch (err) {
   console.error('Export failed', err);
   showToast('Export failed', 'error');
  }
}

function getMonthName(m) {
  const arr = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  return arr[m - 1] || '';
}

/* Overlay and accessibility */
window.addEventListener('click', (ev) => {
  document.querySelectorAll('.modal').forEach(m => {
   if (m.style.display === 'flex' && ev.target === m) m.style.display = 'none';
  });
});
window.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') document.querySelectorAll('.modal').forEach(m => m.style.display = 'none');
});

/* Shortcuts for navigation */
document.getElementById('view-all-bills')?.addEventListener('click', () => {
  navLinks.forEach(l => l.classList.remove('active'));
  document.querySelector('.nav-link[data-target="bills"]').classList.add('active');
  sections.forEach(s => s.classList.remove('active'));
  document.getElementById('bills').classList.add('active');
});
document.getElementById('refresh-bills')?.addEventListener('click', () => renderBills());
