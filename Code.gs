/**
 * CADU - Catatan Keuangan Digital Anda
 * Backend Google Apps Script + Google Sheets
 * Versi 1.3.0
 *
 * PENTING:
 * 1. Hubungkan project Apps Script ke Google Spreadsheet.
 * 2. Jalankan setupCADU() satu kali.
 * 3. Deploy sebagai Web App. Untuk isolasi data per akun,
 *    gunakan Execute as: User accessing the web app.
 * 4. Session.getActiveUser().getEmail() dapat kosong pada konfigurasi
 *    akun/deployment tertentu. Aplikasi akan menampilkan pesan yang jelas.
 */

const CADU = {
  NAME: 'CADU',
  VERSION: '1.4.0',
  TZ: 'Asia/Jakarta',
  TABLES: {
    TRANSACTIONS: [
      'Transaction_ID','User_ID','Date','Type','Category_ID',
      'Description','Amount','Payment_Method','Created_At','Updated_At','Status','Account_ID'
    ],
    CATEGORIES: ['Category_ID','Category_Name','Type','Status'],
    ACCOUNTS: ['Account_ID','User_ID','Account_Name','Account_Type','Bank_Name','Opening_Balance','Status','Created_At'],
    BUDGETS: ['Budget_ID','User_ID','Month','Category_ID','Limit','Created_At'],
    GOALS: [
      'Goal_ID','User_ID','Goal_Name','Target_Amount','Current_Amount',
      'Target_Date','Status','Created_At'
    ],
    SAVINGS: ['Saving_ID','User_ID','Goal_ID','Date','Amount','Note','Created_At','Status'],
    AUTO_SAVINGS: ['Auto_ID','User_ID','Goal_ID','Title','Amount','Day_Of_Month','Start_Date','Last_Run_Month','Status','Created_At'],
    REMINDERS: ['Reminder_ID','User_ID','Title','Reminder_Date','Reminder_Time','Frequency','Email','Last_Sent_Date','Status','Created_At'],
    SETTINGS: ['Setting_Key','Setting_Value'],
    USERS: ['User_ID','Email','Display_Name','Created_At','Status'],
    AUDIT_LOG: ['Log_ID','User_ID','Action','Record_ID','Timestamp','Description']
  },
  DEFAULT_CATEGORIES: [
    ['CAT001','Gaji','INCOME','ACTIVE'],
    ['CAT002','Bonus','INCOME','ACTIVE'],
    ['CAT003','Investasi','INCOME','ACTIVE'],
    ['CAT004','Makanan','EXPENSE','ACTIVE'],
    ['CAT005','Transportasi','EXPENSE','ACTIVE'],
    ['CAT006','Belanja','EXPENSE','ACTIVE'],
    ['CAT007','Tagihan','EXPENSE','ACTIVE'],
    ['CAT008','Hiburan','EXPENSE','ACTIVE'],
    ['CAT009','Kesehatan','EXPENSE','ACTIVE'],
    ['CAT010','Pendidikan','EXPENSE','ACTIVE'],
    ['CAT011','Lain-lain','EXPENSE','ACTIVE']
  ]
};

/* ========================= WEB APP & SETUP ========================= */

function doGet() {
  return ContentService.createTextOutput('CADU Demo API aktif. Gunakan frontend Vercel untuk membuka aplikasi.');
}

function setupCADU() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Buka Apps Script dari Google Spreadsheet (Extensions > Apps Script).');

  PropertiesService.getScriptProperties().setProperty('CADU_DATABASE_ID', ss.getId());

  Object.keys(CADU.TABLES).forEach(name => {
    const headers = CADU.TABLES[name];
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);

    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    } else {
      const current = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), headers.length))
        .getValues()[0];
      const existing = current.filter(String);
      const missing = headers.filter(h => !existing.includes(h));
      if (missing.length) {
        sh.getRange(1, sh.getLastColumn() + 1, 1, missing.length).setValues([missing]);
      }
      const reordered = headers.every(h => existing.includes(h));
      if (!reordered) throw new Error('Header sheet ' + name + ' tidak cocok. Periksa baris pertama.');
    }

    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length)
      .setBackground('#17365D')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
    sh.autoResizeColumns(1, headers.length);
  });

  const categorySheet = ss.getSheetByName('CATEGORIES');
  if (categorySheet.getLastRow() === 1) {
    categorySheet.getRange(2, 1, CADU.DEFAULT_CATEGORIES.length, 4)
      .setValues(CADU.DEFAULT_CATEGORIES);
  }

  const settingsSheet = ss.getSheetByName('SETTINGS');
  if (settingsSheet.getLastRow() === 1) {
    settingsSheet.getRange(2, 1, 3, 2).setValues([
      ['APP_NAME', CADU.NAME],
      ['CURRENCY', 'IDR'],
      ['APP_VERSION', CADU.VERSION]
    ]);
  }

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('CADU berhasil disiapkan. Sheet tabungan, rencana menabung otomatis, dan pengingat tersedia.');
}

/* ========================= DATABASE HELPERS ========================= */

function db_() {
  const id = PropertiesService.getScriptProperties().getProperty('CADU_DATABASE_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Database belum terhubung. Jalankan setupCADU() terlebih dahulu.');
  return ss;
}

function sh_(name) {
  const sheet = db_().getSheetByName(name);
  if (!sheet) throw new Error('Sheet ' + name + ' tidak ditemukan. Jalankan setupCADU().');
  return sheet;
}

function all_(name) {
  const values = sh_(name).getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1).filter(row => row.some(v => v !== '' && v !== null))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
}

function append_(name, obj) {
  const headers = CADU.TABLES[name];
  sh_(name).appendRow(headers.map(h => obj[h] !== undefined ? obj[h] : ''));
}

function rowBy_(name, column, value) {
  const values = sh_(name).getDataRange().getValues();
  if (values.length < 2) return -1;
  const col = values[0].indexOf(column);
  if (col < 0) throw new Error('Kolom ' + column + ' tidak ditemukan di ' + name);
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][col]) === String(value)) return i + 1;
  }
  return -1;
}

function update_(name, rowNumber, patch) {
  const headers = CADU.TABLES[name];
  const range = sh_(name).getRange(rowNumber, 1, 1, headers.length);
  const old = range.getValues()[0];
  range.setValues([headers.map((h, i) => patch[h] !== undefined ? patch[h] : old[i])]);
}

function uid_(prefix) {
  return prefix + Utilities.getUuid().replace(/-/g, '').substring(0, 12).toUpperCase();
}

function now_() { return new Date(); }

function isoDate_(value) {
  if (!value) return '';
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, CADU.TZ, 'yyyy-MM-dd');
  }
  const s = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const d = new Date(value);
  return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, CADU.TZ, 'yyyy-MM-dd');
}

function validDate_(s) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(s || ''))) return false;
  const d = new Date(String(s) + 'T00:00:00');
  return !isNaN(d.getTime()) && isoDate_(d) === s;
}

function num_(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function monthNow_() {
  return Utilities.formatDate(now_(), CADU.TZ, 'yyyy-MM');
}

function audit_(userId, action, recordId, description) {
  append_('AUDIT_LOG', {
    Log_ID: uid_('LOG'),
    User_ID: userId,
    Action: action,
    Record_ID: recordId || '',
    Timestamp: now_(),
    Description: description || ''
  });
}

/* ========================= USER IDENTITY & PROFILE ========================= */

function currentUser_() {
  // Seluruh pengunjung memakai identitas DEMO yang sama.
  // Database wajib berupa spreadsheet DEMO terpisah.
  return { id: 'CADU_PUBLIC_DEMO', email: 'demo@cadu.local', name: 'Pengguna Demo' };
}

function getSetting_(key) {
  const row = all_('SETTINGS').find(x => String(x.Setting_Key) === String(key));
  return row ? row.Setting_Value : '';
}

function putSetting_(key, value) {
  const row = rowBy_('SETTINGS', 'Setting_Key', key);
  if (row > 0) update_('SETTINGS', row, { Setting_Value: value });
  else append_('SETTINGS', { Setting_Key: key, Setting_Value: value });
}

function owned_(table, userId) {
  return all_(table).filter(x => String(x.User_ID) === String(userId));
}

function getMyProfile() {
  const u = currentUser_();
  return {
    User_ID: u.id,
    Display_Name: u.name,
    Email: u.email,
    App_Name: CADU.NAME,
    App_Version: CADU.VERSION
  };
}

/* ========================= CATEGORIES ========================= */

function getCategories() {
  currentUser_();
  return all_('CATEGORIES')
    .filter(c => c.Status === 'ACTIVE')
    .map(c => ({
      Category_ID: String(c.Category_ID),
      Category_Name: String(c.Category_Name),
      Type: String(c.Type),
      Status: String(c.Status)
    }));
}

/* ========================= MASTER KATEGORI & AKUN ========================= */
function saveCategory(data) {
  currentUser_();
  const name=String(data&&data.Category_Name||'').trim(), type=String(data&&data.Type||'');
  if(!name || name.length>60) throw new Error('Nama kategori wajib diisi (maksimal 60 karakter).');
  if(!['INCOME','EXPENSE'].includes(type)) throw new Error('Jenis kategori tidak valid.');
  const all=all_('CATEGORIES');
  const id=String(data.Category_ID||'');
  const duplicate=all.find(c=>String(c.Category_Name).toLowerCase()===name.toLowerCase()&&c.Type===type&&String(c.Category_ID)!==id&&c.Status==='ACTIVE');
  if(duplicate) throw new Error('Kategori dengan nama tersebut sudah ada.');
  if(id){const row=rowBy_('CATEGORIES','Category_ID',id);const old=all.find(c=>String(c.Category_ID)===id);if(row<0||!old)throw new Error('Kategori tidak ditemukan.');update_('CATEGORIES',row,{Category_Name:name,Type:type,Status:String(data.Status||'ACTIVE')});return {success:true};}
  const newId=uid_('CAT');append_('CATEGORIES',{Category_ID:newId,Category_Name:name,Type:type,Status:'ACTIVE'});return {success:true,id:newId};
}
function deleteCategory(id){currentUser_();const cat=all_('CATEGORIES').find(c=>String(c.Category_ID)===String(id));if(!cat)throw new Error('Kategori tidak ditemukan.');const used=all_('TRANSACTIONS').some(t=>String(t.Category_ID)===String(id)&&t.Status==='ACTIVE')||all_('BUDGETS').some(b=>String(b.Category_ID)===String(id));if(used)throw new Error('Kategori sudah digunakan. Nonaktifkan kategori agar riwayat tetap aman.');update_('CATEGORIES',rowBy_('CATEGORIES','Category_ID',id),{Status:'INACTIVE'});return {success:true};}
function getAccounts(){const u=currentUser_();return owned_('ACCOUNTS',u.id).filter(a=>a.Status==='ACTIVE').map(a=>({Account_ID:String(a.Account_ID),Account_Name:String(a.Account_Name),Account_Type:String(a.Account_Type),Bank_Name:String(a.Bank_Name||''),Opening_Balance:num_(a.Opening_Balance),Status:String(a.Status)}));}
function saveAccount(data){const u=currentUser_();const name=String(data&&data.Account_Name||'').trim(),type=String(data&&data.Account_Type||'');if(!name||name.length>80)throw new Error('Nama akun wajib diisi (maksimal 80 karakter).');if(!['CASH','BANK','EWALLET','CREDIT'].includes(type))throw new Error('Jenis akun tidak valid.');const id=String(data.Account_ID||'');const duplicate=owned_('ACCOUNTS',u.id).find(a=>a.Status==='ACTIVE'&&String(a.Account_Name).toLowerCase()===name.toLowerCase()&&String(a.Account_ID)!==id);if(duplicate)throw new Error('Nama akun sudah digunakan.');const payload={Account_Name:name,Account_Type:type,Bank_Name:String(data.Bank_Name||''),Opening_Balance:num_(data.Opening_Balance),Status:'ACTIVE'};if(payload.Opening_Balance<0)throw new Error('Saldo awal tidak boleh negatif.');if(id){const old=owned_('ACCOUNTS',u.id).find(a=>String(a.Account_ID)===id);if(!old)throw new Error('Akun tidak ditemukan atau akses ditolak.');update_('ACCOUNTS',rowBy_('ACCOUNTS','Account_ID',id),payload);}else append_('ACCOUNTS',Object.assign({Account_ID:uid_('ACC'),User_ID:u.id,Created_At:now_()},payload));return {success:true};}
function deleteAccount(id){const u=currentUser_();const a=owned_('ACCOUNTS',u.id).find(x=>String(x.Account_ID)===String(id)&&x.Status==='ACTIVE');if(!a)throw new Error('Akun tidak ditemukan.');update_('ACCOUNTS',rowBy_('ACCOUNTS','Account_ID',id),{Status:'INACTIVE'});return {success:true};}
function seedCADUDemoData(){const u=currentUser_();if(owned_('TRANSACTIONS',u.id).some(t=>t.Status==='ACTIVE'))throw new Error('Data transaksi sudah ada. Data dummy tidak ditambahkan agar data asli aman.');const cats=all_('CATEGORIES');const find=(n,t)=>{let c=cats.find(x=>x.Category_Name===n&&x.Type===t&&x.Status==='ACTIVE');if(!c){const id=uid_('CAT');append_('CATEGORIES',{Category_ID:id,Category_Name:n,Type:t,Status:'ACTIVE'});c={Category_ID:id,Category_Name:n,Type:t,Status:'ACTIVE'};cats.push(c);}return c.Category_ID;};const inc=find('Gaji','INCOME'),bonus=find('Bonus','INCOME'),food=find('Makanan','EXPENSE'),transport=find('Transportasi','EXPENSE'),bill=find('Tagihan','EXPENSE'),shop=find('Belanja','EXPENSE'),health=find('Kesehatan','EXPENSE'),ent=find('Hiburan','EXPENSE');const examples=[[-50,'INCOME',inc,'Gaji bulanan',6500000,'Transfer Bank'],[-45,'EXPENSE',food,'Belanja kebutuhan dapur',650000,'QRIS'],[-43,'EXPENSE',transport,'Transportasi kerja',250000,'E-Wallet'],[-40,'EXPENSE',bill,'Tagihan listrik dan internet',550000,'Transfer Bank'],[-35,'EXPENSE',shop,'Belanja bulanan',720000,'Kartu Debit'],[-30,'INCOME',bonus,'Bonus kinerja',500000,'Transfer Bank'],[-27,'EXPENSE',health,'Vitamin dan kesehatan',180000,'Tunai'],[-22,'EXPENSE',ent,'Nonton dan hiburan',150000,'E-Wallet'],[-15,'EXPENSE',food,'Makan bersama keluarga',300000,'QRIS'],[-8,'EXPENSE',transport,'Bensin',200000,'Tunai'],[-4,'EXPENSE',bill,'Pulsa dan internet',150000,'E-Wallet'],[-2,'INCOME',inc,'Pendapatan tambahan',350000,'Transfer Bank']];examples.forEach(([offset,type,cat,desc,amount,method])=>{const d=new Date();d.setDate(d.getDate()+offset);const date=Utilities.formatDate(d,CADU.TZ,'yyyy-MM-dd');append_('TRANSACTIONS',{Transaction_ID:uid_('TRX'),User_ID:u.id,Date:date,Type:type,Category_ID:cat,Description:desc,Amount:amount,Payment_Method:method,Account_ID:'',Created_At:now_(),Updated_At:now_(),Status:'ACTIVE'});});const existing=owned_('ACCOUNTS',u.id);if(!existing.length){[['Tunai','CASH','',500000],['BCA - Tabungan','BANK','BCA',2500000],['GoPay','EWALLET','GoPay',250000]].forEach(([n,t,b,v])=>append_('ACCOUNTS',{Account_ID:uid_('ACC'),User_ID:u.id,Account_Name:n,Account_Type:t,Bank_Name:b,Opening_Balance:v,Status:'ACTIVE',Created_At:now_()}));}return {success:true,message:'Data contoh berhasil dibuat: 12 transaksi dan 3 akun contoh.'};}

/* ========================= TRANSACTIONS ========================= */

function getTransactions() {
  const user = currentUser_();
  const categories = all_('CATEGORIES');

  return owned_('TRANSACTIONS', user.id)
    .filter(t => t.Status === 'ACTIVE')
    .map(t => {
      const c = categories.find(x => String(x.Category_ID) === String(t.Category_ID));
      return {
        Transaction_ID: String(t.Transaction_ID),
        User_ID: String(t.User_ID),
        Date: isoDate_(t.Date),
        Type: String(t.Type),
        Category_ID: String(t.Category_ID),
        Category_Name: c ? String(c.Category_Name) : 'Kategori tidak ditemukan',
        Description: String(t.Description || ''),
        Amount: num_(t.Amount),
        Payment_Method: String(t.Payment_Method || ''),
        Account_ID: String(t.Account_ID || ''),
        Status: String(t.Status)
      };
    })
    .sort((a, b) => b.Date.localeCompare(a.Date));
}

function saveTransaction(data) {
  if (!data || typeof data !== 'object') throw new Error('Data transaksi tidak valid.');
  const user = currentUser_();
  const type = String(data.Type || '');
  const date = String(data.Date || '');
  const amount = num_(data.Amount);
  const description = String(data.Description || '').trim();
  const categoryId = String(data.Category_ID || '');

  if (!['INCOME', 'EXPENSE'].includes(type)) throw new Error('Jenis transaksi tidak valid.');
  if (!validDate_(date)) throw new Error('Tanggal transaksi tidak valid.');
  if (!description) throw new Error('Deskripsi wajib diisi.');
  if (description.length > 250) throw new Error('Deskripsi maksimal 250 karakter.');
  if (amount <= 0) throw new Error('Nominal harus lebih besar dari nol.');

  const category = all_('CATEGORIES').find(c =>
    String(c.Category_ID) === categoryId && c.Type === type && c.Status === 'ACTIVE'
  );
  if (!category) throw new Error('Kategori tidak sesuai dengan jenis transaksi.');

  const txId = String(data.Transaction_ID || '').trim();
  const row = txId ? rowBy_('TRANSACTIONS', 'Transaction_ID', txId) : -1;
  const timestamp = now_();

  if (row > 0) {
    const old = all_('TRANSACTIONS').find(t => String(t.Transaction_ID) === txId);
    if (!old || String(old.User_ID) !== user.id || old.Status !== 'ACTIVE') {
      throw new Error('Transaksi tidak ditemukan atau akses ditolak.');
    }
    update_('TRANSACTIONS', row, {
      Date: date, Type: type, Category_ID: categoryId,
      Description: description, Amount: amount,
      Payment_Method: String(data.Payment_Method || 'Tunai'),
      Account_ID: String(data.Account_ID || ''),
      Updated_At: timestamp
    });
    audit_(user.id, 'UPDATE_TRANSACTION', txId, description);
    return { success: true, message: 'Transaksi diperbarui.' };
  }

  const newId = uid_('TRX');
  append_('TRANSACTIONS', {
    Transaction_ID: newId, User_ID: user.id, Date: date, Type: type,
    Category_ID: categoryId, Description: description, Amount: amount,
    Payment_Method: String(data.Payment_Method || 'Tunai'),
    Account_ID: String(data.Account_ID || ''),
    Created_At: timestamp, Updated_At: timestamp, Status: 'ACTIVE'
  });
  audit_(user.id, 'CREATE_TRANSACTION', newId, description);
  return { success: true, id: newId, message: 'Transaksi disimpan.' };
}

function deleteTransaction(transactionId) {
  const user = currentUser_();
  const id = String(transactionId || '');
  const row = rowBy_('TRANSACTIONS', 'Transaction_ID', id);
  if (row < 0) throw new Error('Transaksi tidak ditemukan.');

  const item = all_('TRANSACTIONS').find(t => String(t.Transaction_ID) === id);
  if (!item || String(item.User_ID) !== user.id || item.Status !== 'ACTIVE') {
    throw new Error('Transaksi tidak ditemukan atau akses ditolak.');
  }

  update_('TRANSACTIONS', row, { Status: 'DELETED', Updated_At: now_() });
  audit_(user.id, 'DELETE_TRANSACTION', id, item.Description);
  return { success: true };
}

/* ========================= DASHBOARD ========================= */

function getDashboardData() {
  const user = currentUser_();
  const tx = owned_('TRANSACTIONS', user.id).filter(t => t.Status === 'ACTIVE');
  const budgets = owned_('BUDGETS', user.id);
  const savings = owned_('SAVINGS', user.id).filter(x => x.Status === 'ACTIVE');
  const categories = all_('CATEGORIES');
  const thisMonth = monthNow_();

  let lifetimeIncome = 0;
  let lifetimeExpense = 0;
  let monthIncome = 0;
  let monthExpense = 0;

  tx.forEach(t => {
    const amount = num_(t.Amount);
    const month = isoDate_(t.Date).slice(0, 7);
    if (t.Type === 'INCOME') {
      lifetimeIncome += amount;
      if (month === thisMonth) monthIncome += amount;
    } else if (t.Type === 'EXPENSE') {
      lifetimeExpense += amount;
      if (month === thisMonth) monthExpense += amount;
    }
  });

  const monthBudget = budgets
    .filter(b => String(b.Month) === thisMonth)
    .reduce((sum, b) => sum + num_(b.Limit), 0);

  const monthlyData = [];
  for (let i = 5; i >= 0; i--) {
    const d = new Date();
    d.setDate(1);
    d.setMonth(d.getMonth() - i);
    const key = Utilities.formatDate(d, CADU.TZ, 'yyyy-MM');
    const list = tx.filter(t => isoDate_(t.Date).slice(0, 7) === key);
    monthlyData.push({
      month: Utilities.formatDate(d, CADU.TZ, 'MMM yy'),
      income: list.filter(t => t.Type === 'INCOME').reduce((s, t) => s + num_(t.Amount), 0),
      expense: list.filter(t => t.Type === 'EXPENSE').reduce((s, t) => s + num_(t.Amount), 0),
      savings: savings.filter(x => isoDate_(x.Date).slice(0, 7) === key).reduce((s, x) => s + num_(x.Amount), 0)
    });
  }

  const categoryExpenses = categories
    .filter(c => c.Type === 'EXPENSE' && c.Status === 'ACTIVE')
    .map(c => ({
      name: c.Category_Name,
      amount: tx.filter(t =>
        t.Type === 'EXPENSE' &&
        String(t.Category_ID) === String(c.Category_ID) &&
        isoDate_(t.Date).slice(0, 7) === thisMonth
      ).reduce((s, t) => s + num_(t.Amount), 0)
    }))
    .filter(c => c.amount > 0)
    .sort((a, b) => b.amount - a.amount);

  return {
    totalBalance: lifetimeIncome - lifetimeExpense,
    totalIncome: monthIncome,
    totalExpense: monthExpense,
    remainingBudget: monthBudget - monthExpense,
    monthlyData: monthlyData,
    categoryExpenses: categoryExpenses,
    totalSavings: savings.reduce((s, x) => s + num_(x.Amount), 0),
    savingRate: monthIncome > 0 ? Math.round((savings.filter(x => isoDate_(x.Date).slice(0,7) === thisMonth).reduce((s,x)=>s+num_(x.Amount),0) / monthIncome) * 1000) / 10 : 0,
    monthSavings: savings.filter(x => isoDate_(x.Date).slice(0,7) === thisMonth).reduce((s,x)=>s+num_(x.Amount),0)
  };
}

/* ========================= BUDGETS ========================= */

function getBudgets() {
  const user = currentUser_();
  const categories = all_('CATEGORIES');
  const transactions = owned_('TRANSACTIONS', user.id)
    .filter(t => t.Status === 'ACTIVE');

  return owned_('BUDGETS', user.id).map(b => {
    const cat = categories.find(c => String(c.Category_ID) === String(b.Category_ID));
    const used = transactions.filter(t =>
      t.Type === 'EXPENSE' &&
      String(t.Category_ID) === String(b.Category_ID) &&
      isoDate_(t.Date).slice(0, 7) === String(b.Month)
    ).reduce((s, t) => s + num_(t.Amount), 0);

    return {
      Budget_ID: String(b.Budget_ID),
      Month: String(b.Month),
      Category_ID: String(b.Category_ID),
      Category_Name: cat ? String(cat.Category_Name) : 'Kategori',
      Limit: num_(b.Limit),
      Used_Amount: used
    };
  }).sort((a, b) => b.Month.localeCompare(a.Month));
}

function saveBudget(data) {
  const user = currentUser_();
  const month = String(data.Month || '');
  const categoryId = String(data.Category_ID || '');
  const limit = num_(data.Limit);

  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(month)) throw new Error('Bulan anggaran tidak valid.');
  if (limit <= 0) throw new Error('Batas anggaran harus lebih dari nol.');

  const cat = all_('CATEGORIES').find(c =>
    String(c.Category_ID) === categoryId && c.Type === 'EXPENSE' && c.Status === 'ACTIVE'
  );
  if (!cat) throw new Error('Pilih kategori pengeluaran yang valid.');

  const existing = owned_('BUDGETS', user.id).find(b =>
    String(b.Month) === month && String(b.Category_ID) === categoryId
  );

  if (existing) {
    update_('BUDGETS', rowBy_('BUDGETS', 'Budget_ID', existing.Budget_ID), { Limit: limit });
    audit_(user.id, 'UPDATE_BUDGET', existing.Budget_ID, month);
  } else {
    const id = uid_('BDG');
    append_('BUDGETS', {
      Budget_ID: id, User_ID: user.id, Month: month,
      Category_ID: categoryId, Limit: limit, Created_At: now_()
    });
    audit_(user.id, 'CREATE_BUDGET', id, month);
  }
  return { success: true };
}

function deleteBudget(budgetId) {
  const user = currentUser_();
  const id = String(budgetId || '');
  const row = rowBy_('BUDGETS', 'Budget_ID', id);
  const item = owned_('BUDGETS', user.id).find(b => String(b.Budget_ID) === id);
  if (row < 0 || !item) throw new Error('Anggaran tidak ditemukan atau akses ditolak.');
  sh_('BUDGETS').deleteRow(row);
  audit_(user.id, 'DELETE_BUDGET', id, item.Month);
  return { success: true };
}

/* ========================= GOALS ========================= */

function getGoals() {
  const user = currentUser_();
  return owned_('GOALS', user.id)
    .filter(g => g.Status === 'ACTIVE')
    .map(g => ({
      Goal_ID: String(g.Goal_ID),
      Goal_Name: String(g.Goal_Name),
      Target_Amount: num_(g.Target_Amount),
      Current_Amount: num_(g.Current_Amount),
      Target_Date: isoDate_(g.Target_Date),
      Status: String(g.Status)
    }));
}

function saveGoal(data) {
  const user = currentUser_();
  const name = String(data.Goal_Name || '').trim();
  const target = num_(data.Target_Amount);
  const current = num_(data.Current_Amount);
  const targetDate = String(data.Target_Date || '');

  if (!name) throw new Error('Nama target wajib diisi.');
  if (name.length > 100) throw new Error('Nama target maksimal 100 karakter.');
  if (target <= 0) throw new Error('Target harus lebih besar dari nol.');
  if (current < 0) throw new Error('Saldo terkumpul tidak boleh negatif.');
  if (targetDate && !validDate_(targetDate)) throw new Error('Tanggal target tidak valid.');

  const id = uid_('GL');
  append_('GOALS', {
    Goal_ID: id, User_ID: user.id, Goal_Name: name,
    Target_Amount: target, Current_Amount: current,
    Target_Date: targetDate, Status: 'ACTIVE', Created_At: now_()
  });
  audit_(user.id, 'CREATE_GOAL', id, name);
  return { success: true };
}

function updateGoalProgress(goalId, currentAmount) {
  const user = currentUser_();
  const id = String(goalId || '');
  const current = num_(currentAmount);
  const item = owned_('GOALS', user.id).find(g =>
    String(g.Goal_ID) === id && g.Status === 'ACTIVE'
  );
  if (!item) throw new Error('Target tidak ditemukan atau akses ditolak.');
  if (current < 0) throw new Error('Jumlah tabungan tidak boleh negatif.');

  update_('GOALS', rowBy_('GOALS', 'Goal_ID', id), { Current_Amount: current });
  audit_(user.id, 'UPDATE_GOAL_PROGRESS', id, String(current));
  return { success: true };
}

function deleteGoal(goalId) {
  const user = currentUser_();
  const id = String(goalId || '');
  const item = owned_('GOALS', user.id).find(g =>
    String(g.Goal_ID) === id && g.Status === 'ACTIVE'
  );
  if (!item) throw new Error('Target tidak ditemukan atau akses ditolak.');
  update_('GOALS', rowBy_('GOALS', 'Goal_ID', id), { Status: 'DELETED' });
  audit_(user.id, 'DELETE_GOAL', id, item.Goal_Name);
  return { success: true };
}

/* ========================= TABUNGAN ========================= */
function getSavings() {
  const user = currentUser_();
  const goals = owned_('GOALS', user.id);
  return owned_('SAVINGS', user.id).filter(x => x.Status === 'ACTIVE').map(x => {
    const goal = goals.find(g => String(g.Goal_ID) === String(x.Goal_ID));
    return { Saving_ID:String(x.Saving_ID), Goal_ID:String(x.Goal_ID || ''), Goal_Name:goal ? String(goal.Goal_Name) : 'Tabungan umum', Date:isoDate_(x.Date), Amount:num_(x.Amount), Note:String(x.Note || '') };
  }).sort((a,b)=>b.Date.localeCompare(a.Date));
}
function saveSaving(data) {
  const user=currentUser_(); const amount=num_(data && data.Amount); const date=String(data && data.Date || '');
  const goalId=String(data && data.Goal_ID || ''); const note=String(data && data.Note || '').trim();
  if(amount<=0) throw new Error('Nominal setoran harus lebih dari nol.');
  if(!validDate_(date)) throw new Error('Tanggal setoran tidak valid.');
  let goal=null;
  if(goalId){goal=owned_('GOALS',user.id).find(g=>String(g.Goal_ID)===goalId && g.Status==='ACTIVE');if(!goal)throw new Error('Target tabungan tidak ditemukan.');}
  const id=uid_('SAV'); append_('SAVINGS',{Saving_ID:id,User_ID:user.id,Goal_ID:goalId,Date:date,Amount:amount,Note:note,Created_At:now_(),Status:'ACTIVE'});
  if(goal) update_('GOALS',rowBy_('GOALS','Goal_ID',goalId),{Current_Amount:num_(goal.Current_Amount)+amount});
  audit_(user.id,'CREATE_SAVING',id,note || String(amount)); return {success:true,id:id};
}
function deleteSaving(savingId) {
  const user=currentUser_(); const id=String(savingId||''); const item=owned_('SAVINGS',user.id).find(x=>String(x.Saving_ID)===id && x.Status==='ACTIVE');
  if(!item)throw new Error('Setoran tidak ditemukan.');
  update_('SAVINGS',rowBy_('SAVINGS','Saving_ID',id),{Status:'DELETED'});
  if(item.Goal_ID){const goal=owned_('GOALS',user.id).find(g=>String(g.Goal_ID)===String(item.Goal_ID));if(goal)update_('GOALS',rowBy_('GOALS','Goal_ID',item.Goal_ID),{Current_Amount:Math.max(0,num_(goal.Current_Amount)-num_(item.Amount))});}
  audit_(user.id,'DELETE_SAVING',id,String(item.Amount));return {success:true};
}

/* ========================= RENCANA MENABUNG OTOMATIS ========================= */
function getAutoSavings(){
 const u=currentUser_();
 return owned_('AUTO_SAVINGS',u.id).filter(x=>x.Status==='ACTIVE').map(x=>({Auto_ID:String(x.Auto_ID),Goal_ID:String(x.Goal_ID||''),Title:String(x.Title||''),Amount:num_(x.Amount),Day_Of_Month:num_(x.Day_Of_Month),Start_Date:isoDate_(x.Start_Date),Last_Run_Month:String(x.Last_Run_Month||'')}));
}
function saveAutoSaving(data){
 const u=currentUser_(); const title=String(data&&data.Title||'').trim(); const amount=num_(data&&data.Amount); const day=num_(data&&data.Day_Of_Month); const start=String(data&&data.Start_Date||''); const goalId=String(data&&data.Goal_ID||'');
 if(!title)throw new Error('Nama rencana wajib diisi.');
 if(amount<=0)throw new Error('Nominal harus lebih dari nol.');
 if(!Number.isInteger(day)||day<1||day>28)throw new Error('Tanggal otomatis harus 1 sampai 28 agar berlaku setiap bulan.');
 if(!validDate_(start))throw new Error('Tanggal mulai tidak valid.');
 if(goalId&&!owned_('GOALS',u.id).some(g=>String(g.Goal_ID)===goalId&&g.Status==='ACTIVE'))throw new Error('Target tabungan tidak ditemukan.');
 const id=uid_('AUT'); append_('AUTO_SAVINGS',{Auto_ID:id,User_ID:u.id,Goal_ID:goalId,Title:title,Amount:amount,Day_Of_Month:day,Start_Date:start,Last_Run_Month:'',Status:'ACTIVE',Created_At:now_()});
 ensureReminderTrigger_(); audit_(u.id,'CREATE_AUTO_SAVING',id,title); return {success:true};
}
function deleteAutoSaving(id){const u=currentUser_();const x=owned_('AUTO_SAVINGS',u.id).find(a=>String(a.Auto_ID)===String(id)&&a.Status==='ACTIVE');if(!x)throw new Error('Rencana menabung tidak ditemukan.');update_('AUTO_SAVINGS',rowBy_('AUTO_SAVINGS','Auto_ID',id),{Status:'DELETED'});audit_(u.id,'DELETE_AUTO_SAVING',id,x.Title);return {success:true};}
function processAutoSavings_(){
 const today=Utilities.formatDate(new Date(),CADU.TZ,'yyyy-MM-dd'); const month=today.slice(0,7); const day=Number(today.slice(8,10));
 all_('AUTO_SAVINGS').filter(a=>a.Status==='ACTIVE'&&Number(a.Day_Of_Month)===day&&isoDate_(a.Start_Date)<=today&&String(a.Last_Run_Month||'')!==month).forEach(a=>{
  const id=uid_('SAV'); append_('SAVINGS',{Saving_ID:id,User_ID:a.User_ID,Goal_ID:a.Goal_ID||'',Date:today,Amount:num_(a.Amount),Note:'Setoran otomatis: '+a.Title,Created_At:now_(),Status:'ACTIVE'});
  if(a.Goal_ID){const goal=all_('GOALS').find(g=>String(g.Goal_ID)===String(a.Goal_ID)&&String(g.User_ID)===String(a.User_ID)&&g.Status==='ACTIVE');if(goal)update_('GOALS',rowBy_('GOALS','Goal_ID',a.Goal_ID),{Current_Amount:num_(goal.Current_Amount)+num_(a.Amount)});}
  update_('AUTO_SAVINGS',rowBy_('AUTO_SAVINGS','Auto_ID',a.Auto_ID),{Last_Run_Month:month});
  audit_(a.User_ID,'AUTO_SAVING_EXECUTED',id,a.Title+' - '+a.Amount);
 });
}

/* ========================= PENGINGAT TABUNGAN ========================= */
function getReminders(){const u=currentUser_();return owned_('REMINDERS',u.id).filter(x=>x.Status==='ACTIVE').map(x=>({Reminder_ID:String(x.Reminder_ID),Title:String(x.Title),Reminder_Date:isoDate_(x.Reminder_Date),Reminder_Time:String(x.Reminder_Time||'08:00'),Frequency:String(x.Frequency||'ONCE'),Email:String(x.Email||u.email)}));}
function saveReminder(data){const u=currentUser_();const title=String(data&&data.Title||'').trim(),date=String(data&&data.Reminder_Date||''),time=String(data&&data.Reminder_Time||'08:00'),freq=String(data&&data.Frequency||'ONCE');
 if(!title)throw new Error('Judul pengingat wajib diisi.');if(!validDate_(date))throw new Error('Tanggal pengingat tidak valid.');if(!/^([01]\d|2[0-3]):[0-5]\d$/.test(time))throw new Error('Jam pengingat tidak valid.');if(!['ONCE','DAILY','WEEKLY','MONTHLY'].includes(freq))throw new Error('Frekuensi tidak valid.');
 const id=uid_('REM');append_('REMINDERS',{Reminder_ID:id,User_ID:u.id,Title:title,Reminder_Date:date,Reminder_Time:time,Frequency:freq,Email:u.email,Last_Sent_Date:'',Status:'ACTIVE',Created_At:now_()});
 audit_(u.id,'CREATE_REMINDER',id,title);return {success:true};}
function deleteReminder(id){const u=currentUser_();const item=owned_('REMINDERS',u.id).find(x=>String(x.Reminder_ID)===String(id)&&x.Status==='ACTIVE');if(!item)throw new Error('Pengingat tidak ditemukan.');update_('REMINDERS',rowBy_('REMINDERS','Reminder_ID',id),{Status:'DELETED'});return {success:true};}
function ensureReminderTrigger_(){const exists=ScriptApp.getProjectTriggers().some(t=>t.getHandlerFunction()==='sendCADUReminders');if(!exists)ScriptApp.newTrigger('sendCADUReminders').timeBased().everyHours(1).create();}
function sendCADUReminders(){
  // Portfolio demo: tidak mengirim email. Proses hanya rencana menabung otomatis.
  processAutoSavings_();
}


/* ========================= SETTINGS ========================= */

function saveSettings(data) {
  const user = currentUser_();
  const name = String(data.Display_Name || '').trim();
  if (!name) throw new Error('Nama pengguna wajib diisi.');
  if (name.length > 80) throw new Error('Nama maksimal 80 karakter.');

  putSetting_('USER:' + user.email + ':DISPLAY_NAME', name);

  const row = rowBy_('USERS', 'User_ID', user.id);
  if (row > 0) update_('USERS', row, { Display_Name: name });

  audit_(user.id, 'UPDATE_SETTINGS', user.id, 'Mengubah nama profil');
  return { success: true };
}

/* ========================= DIAGNOSTICS ========================= */

function checkCADU() {
  const ss = db_();
  const result = {
    app: CADU.NAME,
    version: CADU.VERSION,
    spreadsheet: ss.getName(),
    sheets: {}
  };
  Object.keys(CADU.TABLES).forEach(name => {
    const sheet = ss.getSheetByName(name);
    result.sheets[name] = {
      exists: !!sheet,
      rows: sheet ? Math.max(0, sheet.getLastRow() - 1) : 0
    };
  });
  return result;
}


/* Jalankan SATU KALI setelah seedCADUDemoData() untuk melengkapi contoh demo. */
function seedCADUDemoPortfolioData() {
  const u = currentUser_();
  if (!owned_('TRANSACTIONS', u.id).length) seedCADUDemoData();
  if (owned_('GOALS', u.id).length || owned_('BUDGETS', u.id).length) {
    return { success: true, message: 'Data portfolio tambahan sudah tersedia; tidak ditambahkan ulang.' };
  }
  const cats = all_('CATEGORIES');
  const catId = name => {
    const c = cats.find(x => x.Category_Name === name && x.Type === 'EXPENSE' && x.Status === 'ACTIVE');
    return c ? String(c.Category_ID) : '';
  };
  const today = Utilities.formatDate(new Date(), CADU.TZ, 'yyyy-MM-dd');
  const month = today.slice(0, 7);
  const goal1 = uid_('GL'), goal2 = uid_('GL');
  append_('GOALS', { Goal_ID:goal1, User_ID:u.id, Goal_Name:'Dana Darurat', Target_Amount:15000000, Current_Amount:3500000, Target_Date:'2027-12-31', Status:'ACTIVE', Created_At:now_() });
  append_('GOALS', { Goal_ID:goal2, User_ID:u.id, Goal_Name:'Liburan Keluarga', Target_Amount:8000000, Current_Amount:1750000, Target_Date:'2027-08-31', Status:'ACTIVE', Created_At:now_() });
  append_('SAVINGS', { Saving_ID:uid_('SAV'), User_ID:u.id, Goal_ID:goal1, Date:today, Amount:500000, Note:'Setoran contoh dana darurat', Created_At:now_(), Status:'ACTIVE' });
  append_('SAVINGS', { Saving_ID:uid_('SAV'), User_ID:u.id, Goal_ID:goal2, Date:today, Amount:250000, Note:'Setoran contoh liburan', Created_At:now_(), Status:'ACTIVE' });
  [[catId('Makanan'),1500000],[catId('Transportasi'),600000],[catId('Tagihan'),900000],[catId('Belanja'),1000000]].filter(x=>x[0]).forEach(([category,limit])=>append_('BUDGETS',{Budget_ID:uid_('BDG'),User_ID:u.id,Month:month,Category_ID:category,Limit:limit,Created_At:now_()}));
  append_('REMINDERS', { Reminder_ID:uid_('REM'), User_ID:u.id, Title:'Bayar tagihan listrik', Reminder_Date:today, Reminder_Time:'08:00', Frequency:'MONTHLY', Email:u.email, Last_Sent_Date:'', Status:'ACTIVE', Created_At:now_() });
  const start = today;
  append_('AUTO_SAVINGS', { Auto_ID:uid_('AUT'), User_ID:u.id, Goal_ID:goal1, Title:'Sisihkan gaji untuk dana darurat', Amount:300000, Day_Of_Month:28, Start_Date:start, Last_Run_Month:'', Status:'ACTIVE', Created_At:now_() });
  ensureReminderTrigger_();
  return { success:true, message:'Data portfolio demo berhasil dilengkapi dengan target, setoran, anggaran, pengingat, dan rencana menabung.' };
}

/* ========================= PUBLIC API FOR VERCEL PROXY ========================= */
function doPost(e) {
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const expectedSecret = PropertiesService.getScriptProperties().getProperty('CADU_API_SECRET');
    if (!expectedSecret || String(payload.secret || '') !== expectedSecret) throw new Error('Akses API tidak sah.');
    const result = dispatchCaduApi_(payload.action, payload.args || []);
    return ContentService.createTextOutput(JSON.stringify({ success: true, data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: String(err && err.message || err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function dispatchCaduApi_(action, args) {
  const allowed = {
    getMyProfile: getMyProfile,
    getDashboardData: getDashboardData,
    getTransactions: getTransactions,
    saveTransaction: saveTransaction,
    deleteTransaction: deleteTransaction,
    getCategories: getCategories,
    saveCategory: saveCategory,
    deleteCategory: deleteCategory,
    getAccounts: getAccounts,
    saveAccount: saveAccount,
    deleteAccount: deleteAccount,
    getBudgets: getBudgets,
    saveBudget: saveBudget,
    deleteBudget: deleteBudget,
    getGoals: getGoals,
    saveGoal: saveGoal,
    updateGoalProgress: updateGoalProgress,
    deleteGoal: deleteGoal,
    getSavings: getSavings,
    saveSaving: saveSaving,
    deleteSaving: deleteSaving,
    getAutoSavings: getAutoSavings,
    saveAutoSaving: saveAutoSaving,
    deleteAutoSaving: deleteAutoSaving,
    getReminders: getReminders,
    saveReminder: saveReminder,
    deleteReminder: deleteReminder,
    saveSettings: saveSettings
  };
  if (!Object.prototype.hasOwnProperty.call(allowed, action)) throw new Error('Aksi tidak diizinkan: ' + action);
  if (!Array.isArray(args)) throw new Error('Parameter API tidak valid.');
  if (JSON.stringify(args).length > 20000) throw new Error('Ukuran permintaan terlalu besar.');
  return allowed[action].apply(null, args);
}
