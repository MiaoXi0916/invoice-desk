const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const fs = require('fs');
const { DatabaseSync } = require('node:sqlite');
const XLSX = require('xlsx-js-style');
const { clean, money, invoiceFileName } = require('./invoice-rules.cjs');

let db;
const defaultFolder = path.join(process.cwd(), '我的发票');
const now = () => new Date().toISOString();
function styleWorkbookSheet(sheet, theme = 'project') {
  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1');
  const headerFill = theme === 'all' ? '1F6654' : '365A78';
  const alternateFill = theme === 'all' ? 'F0F7F3' : 'F1F5F8';
  const widths = [];
  for (let c = range.s.c; c <= range.e.c; c += 1) {
    let max = 10;
    for (let r = range.s.r; r <= range.e.r; r += 1) {
      const address = XLSX.utils.encode_cell({ r, c }); const cell = sheet[address];
      if (!cell) continue;
      const text = String(cell.v ?? ''); max = Math.max(max, Math.min(30, text.length + 2));
      cell.s = { alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, font: { name: 'Microsoft YaHei', sz: 10 }, fill: { fgColor: { rgb: r === 0 ? headerFill : (r % 2 ? alternateFill : 'FFFFFF') } }, border: { bottom: { style: 'thin', color: { rgb: 'D9E2DE' } } } };
      if (r === 0) cell.s = { ...cell.s, font: { name: 'Microsoft YaHei', sz: 10, bold: true, color: { rgb: 'FFFFFF' } } };
      if (typeof cell.v === 'number' && /金额|税额/.test(String(sheet[XLSX.utils.encode_cell({ r: 0, c })]?.v || ''))) cell.z = '¥#,##0.00';
    }
    widths.push(Math.min(30, max));
  }
  sheet['!cols'] = widths.map((wch) => ({ wch }));
  sheet['!rows'] = [{ hpt: 26 }, ...Array(Math.max(0, range.e.r)).fill({ hpt: 22 })];
  sheet['!freeze'] = { xSplit: 0, ySplit: 1 };
}

function initDb() {
  const file = path.join(app.getPath('userData'), 'invoices.sqlite');
  db = new DatabaseSync(file);
  db.exec(`
    PRAGMA foreign_keys = ON;
    CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS statuses (id INTEGER PRIMARY KEY, name TEXT UNIQUE NOT NULL, preset INTEGER NOT NULL DEFAULT 0);
    CREATE TABLE IF NOT EXISTS tags (id INTEGER PRIMARY KEY, name TEXT UNIQUE NOT NULL);
    CREATE TABLE IF NOT EXISTS categories (id INTEGER PRIMARY KEY, name TEXT UNIQUE NOT NULL);
    CREATE TABLE IF NOT EXISTS projects (id INTEGER PRIMARY KEY, name TEXT NOT NULL, note TEXT DEFAULT '', created_at TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS invoices (
      id INTEGER PRIMARY KEY, original_path TEXT NOT NULL, current_path TEXT NOT NULL UNIQUE, file_key TEXT UNIQUE,
      status_id INTEGER, category_id INTEGER, product_name TEXT DEFAULT '', total_amount REAL DEFAULT 0, tax_amount REAL DEFAULT 0,
      invoice_number TEXT DEFAULT '', invoice_date TEXT DEFAULT '', seller_name TEXT DEFAULT '', parse_state TEXT NOT NULL DEFAULT '待解析',
      parse_error TEXT DEFAULT '', manual INTEGER NOT NULL DEFAULT 0, notes TEXT DEFAULT '', created_at TEXT NOT NULL,
      FOREIGN KEY(status_id) REFERENCES statuses(id), FOREIGN KEY(category_id) REFERENCES categories(id)
    );
    CREATE TABLE IF NOT EXISTS invoice_tags (invoice_id INTEGER, tag_id INTEGER, PRIMARY KEY(invoice_id, tag_id), FOREIGN KEY(invoice_id) REFERENCES invoices(id) ON DELETE CASCADE, FOREIGN KEY(tag_id) REFERENCES tags(id));
    CREATE TABLE IF NOT EXISTS project_invoices (project_id INTEGER, invoice_id INTEGER, is_primary INTEGER NOT NULL DEFAULT 0, sequence INTEGER, note TEXT DEFAULT '', PRIMARY KEY(project_id, invoice_id), FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE, FOREIGN KEY(invoice_id) REFERENCES invoices(id) ON DELETE CASCADE);
    CREATE TABLE IF NOT EXISTS audit_log (id INTEGER PRIMARY KEY, invoice_id INTEGER, action TEXT NOT NULL, detail TEXT DEFAULT '', created_at TEXT NOT NULL);
  `);
  for (const name of ['未报销', '报销中', '已报销', '不能报销']) db.prepare('INSERT OR IGNORE INTO statuses(name, preset) VALUES (?, 1)').run(name);
  for (const name of ['材料费', '设备费', '办公用品', '差旅费', '服务费', '其他']) db.prepare('INSERT OR IGNORE INTO categories(name) VALUES (?)').run(name);
  db.prepare('INSERT OR IGNORE INTO settings(key, value) VALUES (?, ?)').run('invoice_directory', defaultFolder);
  db.prepare('INSERT OR IGNORE INTO settings(key, value) VALUES (?, ?)').run('theme', 'light');
  try { db.exec("ALTER TABLE project_invoices ADD COLUMN file_path TEXT DEFAULT ''"); } catch { /* Column already exists. */ }
  try { db.exec("ALTER TABLE projects ADD COLUMN archived INTEGER NOT NULL DEFAULT 0"); } catch { /* Column already exists. */ }
  ensureInvoiceLayout();
}

function all(sql, ...args) { return db.prepare(sql).all(...args); }
function one(sql, ...args) { return db.prepare(sql).get(...args); }
function getDirectory() { return one("SELECT value FROM settings WHERE key='invoice_directory'").value; }
function setDirectory(folder) { db.prepare("UPDATE settings SET value=? WHERE key='invoice_directory'").run(folder); }
function ensureInvoiceLayout() {
  let root = getDirectory();
  const migratedRoot = path.join(process.cwd(), '我的发票');
  if ((!fs.existsSync(root) || path.basename(root) === '我的所有发票') && fs.existsSync(migratedRoot)) { root = migratedRoot; setDirectory(root); }
  fs.mkdirSync(root, { recursive: true });
  const master = path.join(root, '所有发票');
  fs.mkdirSync(master, { recursive: true });
  for (const name of fs.readdirSync(root)) {
    const source = path.join(root, name);
    if (source !== master && fs.statSync(source).isFile() && name.toLowerCase().endsWith('.pdf')) {
      let target = path.join(master, name); let suffix = 2;
      while (fs.existsSync(target)) target = path.join(master, `${path.basename(name, '.pdf')}-${suffix++}.pdf`);
      fs.renameSync(source, target);
    }
  }
  for (const invoice of all('SELECT id,current_path FROM invoices')) {
    if (!fs.existsSync(invoice.current_path)) {
      const relocated = path.join(master, path.basename(invoice.current_path));
      if (fs.existsSync(relocated)) db.prepare('UPDATE invoices SET current_path=? WHERE id=?').run(relocated, invoice.id);
    }
  }
  return { root, master };
}
function masterDirectory() { return ensureInvoiceLayout().master; }
function projectDirectory(name) { const folder = path.join(ensureInvoiceLayout().root, clean(name)); fs.mkdirSync(folder, { recursive: true }); return folder; }
function statusId(name = '未报销') { return one('SELECT id FROM statuses WHERE name=?', name)?.id || one('SELECT id FROM statuses ORDER BY id LIMIT 1').id; }

async function pdfText(filePath) {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
  const doc = await pdfjs.getDocument({ data: new Uint8Array(fs.readFileSync(filePath)), disableWorker: true }).promise;
  const pages = [];
  for (let index = 1; index <= doc.numPages; index += 1) {
    const content = await (await doc.getPage(index)).getTextContent();
    const items = content.items.filter((item) => item.str?.trim()).map((item) => ({ text: item.str, x: item.transform[4], y: item.transform[5] })).sort((a, b) => b.y - a.y || a.x - b.x);
    const lines = []; let row = []; let rowY = null;
    for (const item of items) {
      if (rowY === null || Math.abs(item.y - rowY) < 3) row.push(item);
      else { lines.push(row.map((entry) => entry.text).join(' ')); row = [item]; }
      rowY = item.y;
    }
    if (row.length) lines.push(row.map((entry) => entry.text).join(' '));
    pages.push(lines.join('\n'));
  }
  return pages.join('\n');
}

function match(text, pattern) { return text.match(pattern)?.[1]?.replace(/\s+/g, ' ').trim() || ''; }
function suggestCategory(product) {
  const p = product.toLowerCase();
  if (/螺丝|元器件|电路|芯片|核心板|屏幕|电子/.test(p)) return '材料费';
  if (/打印|纸|笔|办公/.test(p)) return '办公用品';
  if (/软件|服务|技术|咨询/.test(p)) return '服务费';
  return '其他';
}
function parseInvoice(text) {
  const normalized = text.replace(/\s+/g, ' ');
  const number = match(normalized, /发票(?:号码|号)\s*[：:]?\s*([A-Z0-9]{8,})/i) || match(normalized, /([A-Z0-9]{8,})\s*发票(?:号码|号)/i);
  const date = match(normalized, /开票日期\s*[：:]?\s*(\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日|\d{4}[-/]\d{1,2}[-/]\d{1,2})/);
  const currencyValues = [...normalized.matchAll(/[¥￥]\s*(\d+(?:\.\d{1,2})?)/g)].map((entry) => entry[1]);
  // PDF text order sometimes places the number before the "价税合计（小写）" label.
  const totalBeforeLabel = normalized.match(/[¥￥]\s*(\d+(?:\.\d{1,2})?)\s*价税合计[\s\S]{0,40}?小写/)?.[1];
  const totalAfterSmallLabel = normalized.match(/小写[^¥￥\d]{0,20}[¥￥]\s*(\d+(?:\.\d{1,2})?)/)?.[1];
  const total = totalBeforeLabel || totalAfterSmallLabel || currencyValues.at(-1);
  const subtotalMatch = normalized.match(/合\s*计\s*[¥￥]?\s*(\d+(?:\.\d{1,2})?)\s*[¥￥]?\s*(\d+(?:\.\d{1,2})?)/);
  const itemBlock = match(normalized, /(?:项目名称|货物或应税劳务、服务名称)[\s\S]{0,500}?(\*[^*]+\*[^\n]*?)(?:合\s*计|价税合计)/);
  const rawProduct = (itemBlock || '').replace(/^\*[^*]+\*/, '');
  const product = clean((rawProduct.match(/^(.+?)(?=\s+(?:\d+(?:\.\d+)?|个|件|把|pcs|台)\b)/i)?.[1] || rawProduct).replace(/\s+(?:个|件|把|pcs|台)\s+\d.*$/i, '') || '待人工补充', '待人工补充');
  const seller = match(normalized, /销\s*售\s*方[\s\S]{0,50}?名称\s*[：:]?\s*([^\s]+)/);
  const subtotal = subtotalMatch ? money(subtotalMatch[1]) : 0;
  const tax = subtotalMatch ? money(subtotalMatch[2]) : 0;
  const amount = money(total || (subtotal + tax));
  if (!number || !amount || product === '待人工补充') return { parse_state: '待人工补充', parse_error: '未能完整识别票面字段', product_name: product, total_amount: amount, tax_amount: tax, invoice_number: number, invoice_date: date, seller_name: seller, category: '其他' };
  return { parse_state: '已解析', parse_error: '', product_name: product, total_amount: amount, tax_amount: tax, invoice_number: number, invoice_date: date, seller_name: seller, category: suggestCategory(product) };
}

function primaryFor(invoiceId) { return one(`SELECT p.*, pi.sequence FROM project_invoices pi JOIN projects p ON p.id=pi.project_id WHERE pi.invoice_id=? AND pi.is_primary=1`, invoiceId); }
function assignSequence(projectId, invoiceId) {
  let row = one('SELECT sequence FROM project_invoices WHERE project_id=? AND invoice_id=?', projectId, invoiceId);
  if (row?.sequence) return row.sequence;
  const next = (one('SELECT COALESCE(MAX(sequence), 0) + 1 AS n FROM project_invoices WHERE project_id=?', projectId)?.n || 1);
  db.prepare('UPDATE project_invoices SET sequence=? WHERE project_id=? AND invoice_id=?').run(next, projectId, invoiceId);
  return next;
}
function unassignedSequence(invoiceId) {
  const invoice = one('SELECT created_at FROM invoices WHERE id=?', invoiceId);
  if (!invoice) return 1;
  return one(`SELECT COUNT(*) + 1 AS n FROM invoices i WHERE NOT EXISTS (SELECT 1 FROM project_invoices pi WHERE pi.invoice_id=i.id) AND (i.created_at < ? OR (i.created_at = ? AND i.id < ?))`, invoice.created_at, invoice.created_at, invoiceId).n;
}
function syncProjectCopy(projectId, invoiceId) {
  const project = one('SELECT * FROM projects WHERE id=?', projectId);
  const relation = one('SELECT * FROM project_invoices WHERE project_id=? AND invoice_id=?', projectId, invoiceId);
  const invoice = one(`SELECT i.*,s.name AS status_name FROM invoices i LEFT JOIN statuses s ON s.id=i.status_id WHERE i.id=?`, invoiceId);
  if (!project || !relation || !invoice || !fs.existsSync(invoice.current_path)) return;
  const sequence = assignSequence(projectId, invoiceId);
  const target = path.join(projectDirectory(project.name), invoiceFileName(invoice, project.archived ? { name: '已报销项目' } : project, sequence));
  if (relation.file_path && relation.file_path !== target && fs.existsSync(relation.file_path)) fs.rmSync(relation.file_path);
  fs.copyFileSync(invoice.current_path, target);
  db.prepare('UPDATE project_invoices SET file_path=? WHERE project_id=? AND invoice_id=?').run(target, projectId, invoiceId);
}
function syncInvoiceCopies(invoiceId) { for (const row of all('SELECT project_id FROM project_invoices WHERE invoice_id=?', invoiceId)) syncProjectCopy(row.project_id, invoiceId); }
function rebuildProjectCopies(projectId) {
  const project = one('SELECT * FROM projects WHERE id=?', projectId);
  if (!project) return;
  const folder = projectDirectory(project.name);
  for (const name of fs.readdirSync(folder)) {
    const file = path.join(folder, name);
    if (fs.statSync(file).isFile() && name.toLowerCase().endsWith('.pdf')) fs.rmSync(file);
  }
  db.prepare('UPDATE project_invoices SET file_path=? WHERE project_id=?').run('', projectId);
  for (const row of all('SELECT invoice_id FROM project_invoices WHERE project_id=? ORDER BY sequence', projectId)) syncProjectCopy(projectId, row.invoice_id);
}
function renameInvoice(id) {
  const invoice = one(`SELECT i.*, s.name AS status_name FROM invoices i LEFT JOIN statuses s ON s.id=i.status_id WHERE i.id=?`, id);
  if (!invoice || !fs.existsSync(invoice.current_path)) return { ok: false, error: '主票据文件不存在，未执行重命名' };
  const primary = primaryFor(id);
  const targetBase = invoiceFileName(invoice, primary ? (primary.archived ? { name: '已报销项目' } : primary) : { name: '无项目' }, primary ? assignSequence(primary.id, id) : unassignedSequence(id));
  const target = path.join(masterDirectory(), targetBase);
  if (target !== invoice.current_path) {
    if (fs.existsSync(target)) fs.rmSync(target);
    fs.renameSync(invoice.current_path, target);
  }
  db.prepare('UPDATE invoices SET current_path=? WHERE id=?').run(target, id);
  syncInvoiceCopies(id);
  return { ok: true, path: target };
}
function syncUnassignedFileNames() {
  for (const invoice of all('SELECT id FROM invoices i WHERE NOT EXISTS (SELECT 1 FROM project_invoices pi WHERE pi.invoice_id=i.id)')) renameInvoice(invoice.id);
}
function decorateInvoices() {
  return all(`SELECT i.*, s.name AS status_name, c.name AS category_name, GROUP_CONCAT(DISTINCT t.name) AS tags, GROUP_CONCAT(DISTINCT p.name) AS project_names, GROUP_CONCAT(DISTINCT p.id || ':' || pi.sequence) AS project_sequences
    FROM invoices i LEFT JOIN statuses s ON s.id=i.status_id LEFT JOIN categories c ON c.id=i.category_id
    LEFT JOIN invoice_tags it ON it.invoice_id=i.id LEFT JOIN tags t ON t.id=it.tag_id LEFT JOIN project_invoices pi ON pi.invoice_id=i.id LEFT JOIN projects p ON p.id=pi.project_id
    GROUP BY i.id ORDER BY i.created_at DESC`).map((i) => ({ ...i, tags: i.tags ? i.tags.split(',') : [], project_names: i.project_names ? i.project_names.split(',') : [], project_sequences: Object.fromEntries((i.project_sequences ? i.project_sequences.split(',') : []).map((entry) => entry.split(':'))) }));
}
async function importOne(filePath, projectId) {
  const originalPath = filePath;
  const sourceKey = `${path.resolve(filePath)}:${fs.statSync(filePath).size}`;
  if (one('SELECT id FROM invoices WHERE file_key=?', sourceKey)) return { skipped: true, file: filePath };
  const master = masterDirectory();
  if (path.dirname(path.resolve(filePath)) !== path.resolve(master)) {
    let stored = path.join(master, clean(path.basename(filePath), `发票-${Date.now()}.pdf`)); let suffix = 2;
    while (fs.existsSync(stored)) stored = path.join(master, `${path.basename(path.basename(filePath), '.pdf')}-${suffix++}.pdf`);
    fs.copyFileSync(filePath, stored); filePath = stored;
  }
  const key = `${path.resolve(filePath)}:${fs.statSync(filePath).size}`;
  if (one('SELECT id FROM invoices WHERE current_path=? OR file_key=?', filePath, key)) return { skipped: true, file: filePath };
  let parsed;
  try { parsed = parseInvoice(await pdfText(filePath)); } catch (error) { parsed = { parse_state: '待人工补充', parse_error: error.message, product_name: '待人工补充', total_amount: 0, tax_amount: 0, invoice_number: '', invoice_date: '', seller_name: '', category: '其他' }; }
  if (parsed.invoice_number) {
    const duplicate = one('SELECT id,current_path,product_name FROM invoices WHERE invoice_number=?', parsed.invoice_number);
    if (duplicate) return { duplicate: true, file: filePath, invoice_number: parsed.invoice_number, existingId: duplicate.id, existingPath: duplicate.current_path };
  }
  const category = one('SELECT id FROM categories WHERE name=?', parsed.category)?.id;
  const result = db.prepare(`INSERT INTO invoices(original_path,current_path,file_key,status_id,category_id,product_name,total_amount,tax_amount,invoice_number,invoice_date,seller_name,parse_state,parse_error,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)`)
    .run(originalPath, filePath, key, statusId(), category, parsed.product_name, parsed.total_amount, parsed.tax_amount, parsed.invoice_number, parsed.invoice_date, parsed.seller_name, parsed.parse_state, parsed.parse_error, now());
  const id = Number(result.lastInsertRowid);
  if (projectId) linkInvoice(projectId, id, true);
  renameInvoice(id);
  return { id, parsed };
}
function linkInvoice(projectId, invoiceId, makePrimary) {
  db.prepare('INSERT OR IGNORE INTO project_invoices(project_id,invoice_id,is_primary) VALUES(?,?,0)').run(projectId, invoiceId);
  assignSequence(projectId, invoiceId);
  if (makePrimary || !primaryFor(invoiceId)) {
    db.prepare('UPDATE project_invoices SET is_primary=0 WHERE invoice_id=?').run(invoiceId);
    db.prepare('UPDATE project_invoices SET is_primary=1 WHERE project_id=? AND invoice_id=?').run(projectId, invoiceId);
  }
  const result = renameInvoice(invoiceId); syncUnassignedFileNames(); return result;
}

function normalizedTheme() { const saved=one("SELECT value FROM settings WHERE key='theme'")?.value || 'light'; const legacy={light:'green-light',dark:'green-dark',ocean:'blue-light',warm:'orange-light',system:'green-light'}; return legacy[saved] || saved; }
function bootstrap() { return { invoices: decorateInvoices(), projects: all('SELECT p.*, COUNT(pi.invoice_id) AS invoice_count, COALESCE(SUM(i.total_amount),0) AS total FROM projects p LEFT JOIN project_invoices pi ON pi.project_id=p.id LEFT JOIN invoices i ON i.id=pi.invoice_id GROUP BY p.id ORDER BY p.created_at DESC'), statuses: all('SELECT * FROM statuses ORDER BY preset DESC,id'), tags: all('SELECT * FROM tags ORDER BY name'), categories: all('SELECT * FROM categories ORDER BY name'), directory: getDirectory(), theme: normalizedTheme() }; }
function reconcileFiles() {
  ensureInvoiceLayout(); const failures = [];
  for (const invoice of all('SELECT id FROM invoices')) { const result = renameInvoice(invoice.id); if (!result.ok) failures.push(invoice.id); }
  for (const project of all('SELECT id FROM projects')) rebuildProjectCopies(project.id);
  return { checked: all('SELECT id FROM invoices').length, failures, data: bootstrap() };
}
async function reparseInvoice(id) {
  const invoice = one('SELECT * FROM invoices WHERE id=?', id);
  if (!invoice || !fs.existsSync(invoice.current_path)) return;
  try {
    const parsed = parseInvoice(await pdfText(invoice.current_path));
    const categoryId = one('SELECT id FROM categories WHERE name=?', parsed.category)?.id;
    db.prepare('UPDATE invoices SET product_name=?,total_amount=?,tax_amount=?,invoice_number=?,invoice_date=?,seller_name=?,parse_state=?,parse_error=?,category_id=? WHERE id=?').run(parsed.product_name, parsed.total_amount, parsed.tax_amount, parsed.invoice_number, parsed.invoice_date, parsed.seller_name, parsed.parse_state, parsed.parse_error, categoryId, id);
    renameInvoice(id);
  } catch { /* Keep the prior record when a PDF cannot be read. */ }
}
async function repairLegacyParsing() {
  if (one("SELECT value FROM settings WHERE key='parser_version'")?.value === '2') return;
  for (const invoice of all('SELECT id FROM invoices WHERE manual=0')) await reparseInvoice(invoice.id);
  db.prepare("INSERT OR REPLACE INTO settings(key,value) VALUES('parser_version','2')").run();
}
function handlers() {
  ipcMain.handle('bootstrap', () => bootstrap());
  ipcMain.handle('choose-directory', async () => { const r = await dialog.showOpenDialog({ properties: ['openDirectory'] }); if (!r.canceled) { setDirectory(r.filePaths[0]); ensureInvoiceLayout(); } return bootstrap(); });
  ipcMain.handle('choose-files', async () => { const r = await dialog.showOpenDialog({ properties: ['openFile', 'multiSelections'], filters: [{ name: 'PDF 发票', extensions: ['pdf'] }] }); return r.canceled ? [] : r.filePaths; });
  ipcMain.handle('import-files', async (_, paths, projectId) => { const results=[]; for(const p of paths) results.push(await importOne(p, projectId)); return { results, data: bootstrap() }; });
  ipcMain.handle('scan-directory', async () => { const dir=masterDirectory(); const files=fs.readdirSync(dir).filter((f)=>f.toLowerCase().endsWith('.pdf')).map((f)=>path.join(dir,f)); const results=[]; for(const p of files) results.push(await importOne(p)); return { results, data: bootstrap() }; });
  ipcMain.handle('update-invoice', (_, id, patch) => { const current=one('SELECT * FROM invoices WHERE id=?', id); const categoryId=patch.category_name ? one('SELECT id FROM categories WHERE name=?',patch.category_name)?.id : current.category_id; const st=patch.status_name ? statusId(patch.status_name) : current.status_id; const product=patch.product_name ?? current.product_name, total=money(patch.total_amount ?? current.total_amount), number=patch.invoice_number ?? current.invoice_number, date=patch.invoice_date ?? current.invoice_date; const complete=product && product!=='待人工补充' && total>0 && number && date; db.prepare('UPDATE invoices SET product_name=?,total_amount=?,tax_amount=?,invoice_number=?,invoice_date=?,notes=?,category_id=?,status_id=?,parse_state=?,parse_error=?,manual=1 WHERE id=?').run(product,total,money(patch.tax_amount ?? current.tax_amount),number,date,patch.notes ?? current.notes,categoryId,st,complete?'已解析':'待人工补充',complete?'':'请补充必填票据信息',id); db.prepare('DELETE FROM invoice_tags WHERE invoice_id=?').run(id); for(const tag of patch.tags||[]) { const row=one('SELECT id FROM tags WHERE name=?',tag); if(row) db.prepare('INSERT INTO invoice_tags(invoice_id,tag_id) VALUES(?,?)').run(id,row.id); } const r=renameInvoice(id); return { ...r, data: bootstrap() }; });
  ipcMain.handle('reparse-invoice', async (_, id) => { await reparseInvoice(id); return bootstrap(); });
  ipcMain.handle('delete-invoice', (_, id, trash) => { const inv=one('SELECT * FROM invoices WHERE id=?',id); for(const row of all('SELECT file_path FROM project_invoices WHERE invoice_id=?',id)) if(row.file_path && fs.existsSync(row.file_path)) fs.rmSync(row.file_path); if(trash && inv && fs.existsSync(inv.current_path)) shell.trashItem(inv.current_path); db.prepare('DELETE FROM invoices WHERE id=?').run(id); syncUnassignedFileNames(); return bootstrap(); });
  ipcMain.handle('create-project', (_, name, note) => { const projectName=clean(name); db.prepare('INSERT INTO projects(name,note,created_at) VALUES(?,?,?)').run(projectName,note||'',now()); projectDirectory(projectName); return bootstrap(); });
  ipcMain.handle('update-project', (_, id, patch) => { const old=one('SELECT * FROM projects WHERE id=?',id); const nextName=clean(patch.name); const renamed=old && old.name!==nextName; if(renamed) { const oldDir=path.join(ensureInvoiceLayout().root,clean(old.name)); const nextDir=path.join(ensureInvoiceLayout().root,nextName); if(fs.existsSync(oldDir) && !fs.existsSync(nextDir)) fs.renameSync(oldDir,nextDir); } db.prepare('UPDATE projects SET name=?,note=? WHERE id=?').run(nextName,patch.note||'',id); for(const row of all('SELECT invoice_id FROM project_invoices WHERE project_id=?',id)) renameInvoice(row.invoice_id); if(renamed) rebuildProjectCopies(id); else for(const row of all('SELECT invoice_id FROM project_invoices WHERE project_id=?',id)) syncProjectCopy(id,row.invoice_id); return bootstrap(); });
  ipcMain.handle('archive-project', (_, id, archived) => { db.prepare('UPDATE projects SET archived=? WHERE id=?').run(archived ? 1 : 0, id); for(const row of all('SELECT invoice_id FROM project_invoices WHERE project_id=?',id)) renameInvoice(row.invoice_id); rebuildProjectCopies(id); return bootstrap(); });
  ipcMain.handle('delete-project', async (_, id) => { const project=one('SELECT * FROM projects WHERE id=?',id); if(!project) return bootstrap(); const invoiceIds=all('SELECT invoice_id,is_primary FROM project_invoices WHERE project_id=?',id); const folder=path.join(ensureInvoiceLayout().root,clean(project.name)); if(fs.existsSync(folder)) await shell.trashItem(folder); db.prepare('DELETE FROM projects WHERE id=?').run(id); for(const row of invoiceIds) { if(row.is_primary) { const next=one('SELECT project_id FROM project_invoices WHERE invoice_id=? ORDER BY sequence LIMIT 1',row.invoice_id); if(next) db.prepare('UPDATE project_invoices SET is_primary=1 WHERE project_id=? AND invoice_id=?').run(next.project_id,row.invoice_id); } renameInvoice(row.invoice_id); } syncUnassignedFileNames(); return bootstrap(); });
  ipcMain.handle('link-invoice', (_, projectId, invoiceId, primary) => ({ result: linkInvoice(projectId,invoiceId,primary), data: bootstrap() }));
  ipcMain.handle('unlink-invoice', (_, projectId, invoiceId) => { const relation=one('SELECT file_path,is_primary FROM project_invoices WHERE project_id=? AND invoice_id=?',projectId,invoiceId); if(relation?.file_path && fs.existsSync(relation.file_path)) fs.rmSync(relation.file_path); db.prepare('DELETE FROM project_invoices WHERE project_id=? AND invoice_id=?').run(projectId,invoiceId); if(relation?.is_primary) { const next=one('SELECT project_id FROM project_invoices WHERE invoice_id=? ORDER BY sequence LIMIT 1',invoiceId); if(next) db.prepare('UPDATE project_invoices SET is_primary=1 WHERE project_id=? AND invoice_id=?').run(next.project_id,invoiceId); } renameInvoice(invoiceId); syncUnassignedFileNames(); return bootstrap(); });
  ipcMain.handle('reconcile-files', () => reconcileFiles());
  ipcMain.handle('open-invoice', (_, p) => shell.openPath(p));
  ipcMain.handle('add-status', (_, name) => { db.prepare('INSERT OR IGNORE INTO statuses(name,preset) VALUES(?,0)').run(clean(name)); return bootstrap(); });
  ipcMain.handle('add-tag', (_, name) => { db.prepare('INSERT OR IGNORE INTO tags(name) VALUES(?)').run(clean(name)); return bootstrap(); });
  ipcMain.handle('add-category', (_, name) => { db.prepare('INSERT OR IGNORE INTO categories(name) VALUES(?)').run(clean(name)); return bootstrap(); });
  ipcMain.handle('set-theme', (_, theme) => { const value=['green-light','green-dark','blue-light','blue-dark','orange-light','orange-dark','pink-light','pink-dark'].includes(theme)?theme:'green-light'; db.prepare("INSERT OR REPLACE INTO settings(key,value) VALUES('theme',?)").run(value); return bootstrap(); });
  ipcMain.handle('export-project', async (_, projectId) => { const project=one('SELECT * FROM projects WHERE id=?',projectId); const rawRows=all(`SELECT pi.sequence AS '序号', i.product_name AS '商品名称', c.name AS '类目', i.total_amount AS '价税合计', i.tax_amount AS '税额', i.invoice_date AS '开票日期', i.invoice_number AS '发票号码', s.name AS '报销状态', i.current_path, i.notes AS '备注', GROUP_CONCAT(p2.name) AS '关联项目' FROM project_invoices pi JOIN invoices i ON i.id=pi.invoice_id LEFT JOIN categories c ON c.id=i.category_id LEFT JOIN statuses s ON s.id=i.status_id LEFT JOIN project_invoices pi2 ON pi2.invoice_id=i.id LEFT JOIN projects p2 ON p2.id=pi2.project_id WHERE pi.project_id=? GROUP BY i.id ORDER BY pi.sequence`,projectId); const rows=rawRows.map(({current_path,...row})=>({'序号':row['序号'],'文件名':path.basename(current_path||''),...row})); const outputPath=path.join(projectDirectory(project.name),`${clean(project.name)}-报销清单.xlsx`); const sheet=XLSX.utils.json_to_sheet(rows); styleWorkbookSheet(sheet,'project'); const book=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(book,sheet,'报销清单'); XLSX.writeFile(book,outputPath); return outputPath; });
  ipcMain.handle('export-all-invoices', async () => { const projects=all('SELECT * FROM projects ORDER BY created_at'); const rowsFor = (where='', args=[]) => all(`SELECT i.product_name AS '商品名称', c.name AS '类目', i.total_amount AS '价税合计', i.tax_amount AS '税额', i.invoice_date AS '开票日期', i.invoice_number AS '发票号码', s.name AS '报销状态', i.notes AS '备注', GROUP_CONCAT(DISTINCT p.name) AS '关联项目' FROM invoices i LEFT JOIN categories c ON c.id=i.category_id LEFT JOIN statuses s ON s.id=i.status_id LEFT JOIN project_invoices pi ON pi.invoice_id=i.id LEFT JOIN projects p ON p.id=pi.project_id ${where} GROUP BY i.id ORDER BY i.created_at`, ...args).map((row, index) => ({ '序号': index + 1, ...row })); const book=XLSX.utils.book_new(); const allSheet=XLSX.utils.json_to_sheet(rowsFor()); styleWorkbookSheet(allSheet,'all'); XLSX.utils.book_append_sheet(book,allSheet,'全部发票'); for(const project of projects) { const projectRows=rowsFor('WHERE EXISTS (SELECT 1 FROM project_invoices px WHERE px.invoice_id=i.id AND px.project_id=?)',[project.id]); let sheetName=clean(project.name).slice(0,31)||`项目${project.id}`; if(book.SheetNames.includes(sheetName)) sheetName=`${sheetName.slice(0,27)}-${project.id}`; const projectSheet=XLSX.utils.json_to_sheet(projectRows); styleWorkbookSheet(projectSheet,'project'); XLSX.utils.book_append_sheet(book,projectSheet,sheetName); } const outputPath=path.join(masterDirectory(),'全部发票-报销清单.xlsx'); XLSX.writeFile(book,outputPath); return outputPath; });
}
app.whenReady().then(async () => { initDb(); await repairLegacyParsing(); reconcileFiles(); handlers(); const win=new BrowserWindow({ width:1440,height:940,minWidth:1100,minHeight:700,webPreferences:{preload:path.join(__dirname,'preload.cjs'),contextIsolation:true,nodeIntegration:false}}); if(process.argv.includes('--dev')) win.loadURL('http://127.0.0.1:5173'); else win.loadFile(path.join(__dirname,'../../发布产物/dist/index.html')); });
app.on('window-all-closed', () => { if(process.platform !== 'darwin') app.quit(); });
