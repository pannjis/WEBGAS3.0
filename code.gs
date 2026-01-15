function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('SiGAS PRO - Sistem Agen Gas')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = [
    {name: 'USERS', header: ['Username', 'Password', 'Role', 'Nama']},
    {name: 'PRODUK', header: ['ID', 'Nama_Produk', 'Harga_Jual', 'Harga_Beli', 'Stok_Isi', 'Stok_Kosong', 'SKU', 'Kode', 'Link_Gambar']},
    {name: 'PELANGGAN', header: ['ID', 'Nama', 'Nama_Perusahan', 'NoHP', 'Alamat']},
    {name: 'SUPPLIER', header: ['ID', 'Nama_Supplier', 'NoHP', 'Alamat']},
    {name: 'TRANSAKSI', header: ['ID_Trans', 'Waktu', 'Pelanggan', 'Produk', 'Qty', 'Total', 'Tipe', 'Kasir', 'Metode_Bayar', 'Jatuh_Tempo', 'Status']},
    {name: 'PEMBELIAN', header: ['ID_Beli', 'Waktu', 'Supplier', 'Produk', 'Qty', 'Total', 'Metode']},
    // [UPDATE] Header Keuangan ditambah kolom 'Akun'
    {name: 'KEUANGAN', header: ['ID', 'Tanggal', 'Jenis', 'Kategori', 'Nominal', 'Keterangan', 'Akun']}, 
    {name: 'KATEGORI', header: ['Nama_Kategori']},
    {name: 'KARYAWAN', header: ['ID', 'Nama', 'NoHP', 'Gaji_Pokok', 'Bonus_Per_Pcs', 'Status']}, 
    {name: 'KASBON', header: ['ID_Kasbon', 'Tanggal', 'Nama_Karyawan', 'Nominal', 'Keterangan', 'Status_Lunas']},
    {name: 'PENGATURAN', header: ['Key', 'Value']},
    // [BARU] Sheet untuk menyimpan Daftar Akun/Dompet
    {name: 'AKUN_KAS', header: ['ID_Akun', 'Nama_Akun', 'No_Rekening', 'Tipe', 'Saldo_Awal']} 
  ];

  sheets.forEach(s => {
    let sheet = ss.getSheetByName(s.name);
    if (!sheet) {
      sheet = ss.insertSheet(s.name);
      sheet.appendRow(s.header);
      // Data Default User
      if(s.name === 'USERS') sheet.appendRow(['admin', 'admin123', 'Admin', 'Super Admin']);
      // [BARU] Data Default Akun Kas
      if(s.name === 'AKUN_KAS') {
         // Urutan: ID, Nama, No_Rekening, Tipe, Saldo
         sheet.appendRow(['ACC-1', 'Kas Tunai (Laci)', '-', 'Tunai', 0]); 
         sheet.appendRow(['ACC-2', 'Bank BCA', '1234567890', 'Bank', 0]);
         sheet.appendRow(['ACC-3', 'Bank BRI', '0987654321', 'Bank', 0]);
      }
    }
  });
}

// 3. Update Baca Data Akun (Sesuaikan Index Kolom)
function getDaftarAkun() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetAkun = ss.getSheetByName('AKUN_KAS');
  const sheetKeu = ss.getSheetByName('KEUANGAN');
  
  if(!sheetAkun || !sheetKeu) return [];

  const dataAkun = sheetAkun.getDataRange().getValues().slice(1);
  const dataKeu = sheetKeu.getDataRange().getValues().slice(1);

  let listAkun = dataAkun.map(r => {
      // [UPDATE INDEX KARENA ADA KOLOM BARU]
      let id = r[0];
      let nama = r[1];
      let norek = r[2]; // Kolom C (Index 2)
      let tipe = r[3];  // Kolom D (Index 3)
      let saldo = Number(r[4]); // Kolom E (Index 4) - Saldo Awal

      // Loop Transaksi Keuangan
      dataKeu.forEach(k => {
          let akunTrx = k[6]; 
          let jenis = k[2];
          let nominal = Number(k[4]);

          if(akunTrx === nama) {
              if(jenis === 'Pemasukan') saldo += nominal;
              if(jenis === 'Pengeluaran') saldo -= nominal;
          }
      });

      // Kembalikan objek lengkap
      return { id: id, nama: nama, norek: norek, tipe: tipe, saldo: saldo };
  });

  return listAkun;
}

// --- UPDATE 2: PERBAIKI LOGIN (Agar me-return Username) ---
function loginUser(username, password) {
  const data = getData('USERS');
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] == username && data[i][1] == password) {
      // Tambahkan username ke return object
      return { status: 'success', role: data[i][2], nama: data[i][3], username: data[i][0] }; 
    }
  }
  return { status: 'failed' };
}

// --- BARU: MANAJEMEN USER ---

function getAllUsers() {
  return getData('USERS');
}

function simpanUserBaru(form) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('USERS');
  const data = sheet.getDataRange().getValues();

  // Cek Duplicate Username (kecuali edit diri sendiri, tapi disini kita asumsikan username key)
  // Mode Edit (Jika password kosong, berarti update data lain saja, tapi disini simple overwrite)
  
  let userExists = false;
  let rowIndex = -1;

  for(let i=1; i<data.length; i++) {
     if(data[i][0] === form.username) {
        userExists = true;
        rowIndex = i + 1;
        break;
     }
  }

  if(form.isEdit && userExists) {
     // Update Data
     // Jika password diisi, update password. Jika tidak, pakai password lama.
     let oldPass = sheet.getRange(rowIndex, 2).getValue();
     let newPass = form.password ? form.password : oldPass;
     
     sheet.getRange(rowIndex, 1, 1, 4).setValues([[form.username, newPass, form.role, form.nama]]);
     return "Data User Berhasil Diupdate";
  } else if (!form.isEdit && userExists) {
     return "Error: Username sudah terpakai!";
  } else {
     // Buat Baru
     sheet.appendRow([form.username, form.password, form.role, form.nama]);
     return "User Baru Berhasil Ditambahkan";
  }
}

function hapusUser(username) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('USERS');
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(data[i][0] == username) {
       sheet.deleteRow(i+1);
       return "User dihapus.";
    }
  }
}

function gantiPasswordSendiri(username, oldPass, newPass) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('USERS');
  const data = sheet.getDataRange().getValues();
  
  for(let i=1; i<data.length; i++) {
    if(data[i][0] == username) {
       if(data[i][1] != oldPass) return "Password Lama Salah!";
       
       sheet.getRange(i+1, 2).setValue(newPass);
       return "Password Berhasil Diganti";
    }
  }
  return "User tidak ditemukan";
}

// --- BARU: PENGATURAN PERUSAHAAN ---

function getProfilPerusahaan() {
  const data = getData('PENGATURAN');
  // Convert Array [Key, Value] menjadi Object {key: value}
  let config = {};
  data.forEach(row => {
     config[row[0]] = row[1];
  });
  return config;
}

// [UPDATE] Fungsi Simpan Profil dengan Fitur Upload Logo
function simpanProfilPerusahaan(form) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('PENGATURAN');
  const data = sheet.getDataRange().getValues();
  
  // Gunakan ID Folder yang sama dengan Produk (atau ganti jika punya folder khusus logo)
  const FOLDER_ID = '15hiLtvusofF2OJpXVq8lJkePbmqVIuPM'; 

  // Helper function update/insert
    const updateOrInsert = (key, val) => {
     let found = false;
     
     // [PERBAIKAN] Paksa jadi String dengan menambahkan tanda petik satu (') di depan
     // Ini trik agar Google Sheet tidak menghapus angka 0 di depan
     let finalVal = val;
     if (key === 'no_perusahaan' || key === 'no_pemilik') {
         finalVal = "'" + val; 
     }

     for(let i=1; i<data.length; i++) {
        if(data[i][0] === key) {
           sheet.getRange(i+1, 2).setValue(finalVal); // Gunakan finalVal
           found = true;
           break;
        }
     }
     if(!found) sheet.appendRow([key, finalVal]); // Gunakan finalVal
  };

  // 1. PROSES UPLOAD LOGO (Jika ada file baru dipilih)
  if (form.logo && form.logo.data) {
    try {
       const decoded = Utilities.base64Decode(form.logo.data);
       const blob = Utilities.newBlob(decoded, form.logo.mimeType, 'LOGO-' + Date.now());
       
       const folder = DriveApp.getFolderById(FOLDER_ID);
       const file = folder.createFile(blob);
       
       // Set Permission agar bisa dilihat publik
       file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
       
       const logoUrl = "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w1000";
       
       // Simpan URL Logo ke Database
       updateOrInsert('logo_perusahaan', logoUrl);

    } catch (e) {
       throw new Error("Gagal Upload Logo: " + e.message);
    }
  }

  // 2. Simpan Data Teks Lainnya
  updateOrInsert('nama_perusahaan', form.nama_perusahaan);
  updateOrInsert('nama_pemilik', form.nama_pemilik);
  updateOrInsert('alamat', form.alamat);
  updateOrInsert('no_perusahaan', form.no_perusahaan);
  updateOrInsert('no_pemilik', form.no_pemilik);

  return "Profil & Logo Berhasil Disimpan!";
}

// GANTI function getData(sheetName) yang lama dengan ini:

function getData(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues().slice(1); // Hapus Header
  
  return data.filter(r => r[0] !== "").map(r => {
      // PERBAIKAN: Paksa konversi Tanggal (Kolom Index 1)
      // Jika format di Excel jadi Teks (Alignment Kiri), kita ubah paksa jadi Date Object
      if (r[1] && !(r[1] instanceof Date)) {
          let d = new Date(r[1]);
          if(!isNaN(d.getTime())) { 
             r[1] = d; // Berhasil dikonversi
          }
      }
      
      // Standarisasi ke ISO String agar Frontend bisa baca
      if (r[1] instanceof Date) {
         // Tambahkan offset agar tidak mundur sehari karena Timezone UTC
         // Kita pakai teknik simpel: Format lokal ID
         let localDate = new Date(r[1].getTime() - (r[1].getTimezoneOffset() * 60000));
         r[1] = localDate.toISOString(); 
      }
      
      return r;
  });
}

// --- LOGIN ---
function loginUser(username, password) {
  const data = getData('USERS');
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] == username && data[i][1] == password) {
      return { status: 'success', role: data[i][2], nama: data[i][3] };
    }
  }
  return { status: 'failed' };
}

// GANTI function getDashboardStats() yang lama dengan ini:

function getDashboardStats() {
  const keu = getData('KEUANGAN');
  let income = 0, expense = 0;
  
  keu.forEach(r => {
    // r[2] adalah Jenis, r[4] adalah Nominal
    // Gunakan String() dan trim() agar aman dari spasi
    let jenis = String(r[2]).trim(); 
    let nominal = Number(r[4]);

    if(jenis === 'Pemasukan') income += nominal;
    if(jenis === 'Pengeluaran') expense += nominal;
  });
  
  return { income, expense, net: income - expense };
}

// [UPDATE] Fungsi Tambah Produk (Versi Debugging)
// [UPDATE] Fungsi Tambah Produk (Upload ke Folder Khusus)
function tambahProduk(form) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('PRODUK');
  
  // ID Folder Google Drive Anda
  const FOLDER_ID = '15hiLtvusofF2OJpXVq8lJkePbmqVIuPM'; 
  
  let imageUrl = '';

  // PROSES UPLOAD
  if (form.gambar && form.gambar.data) {
    try {
      const decoded = Utilities.base64Decode(form.gambar.data);
      const blob = Utilities.newBlob(decoded, form.gambar.mimeType, form.gambar.fileName);
      
      // 1. Ambil Folder Tujuan
      const folder = DriveApp.getFolderById(FOLDER_ID);
      
      // 2. Simpan File di Folder Tersebut
      const file = folder.createFile(blob); 
      
      // 3. Set Permission (Coba Publik -> Domain -> Private)
      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (e1) {
        try {
           file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
        } catch (e2) {
           console.log("Gagal set permission: " + e1.message); 
        }
      }

      // 4. Ambil Link
      // Ganti format link jadi Thumbnail (agar tidak crash/broken di browser)
      imageUrl = "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w1000";

    } catch (e) {
      // Tampilkan error detail jika gagal
      throw new Error("Gagal Upload: " + e.message); 
    }
  } else {
    // Jika manual link
    imageUrl = (typeof form.gambar === 'string') ? form.gambar : '';
  }

  // Simpan ke Spreadsheet
  sheet.appendRow([
    'P-' + Date.now(), 
    form.nama, 
    form.hargaJual, 
    form.hargaBeli, 
    form.stokIsi, 
    form.stokKosong,
    form.sku,     
    form.kode,    
    imageUrl 
  ]);
}

// [BARU] Fungsi Update Produk (Edit Mode)
function updateProduk(form) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('PRODUK');
  const data = sheet.getDataRange().getValues();
  
  // ID Folder Google Drive (Sama seperti tambah produk)
  const FOLDER_ID = '15hiLtvusofF2OJpXVq8lJkePbmqVIuPM'; 

  let rowTarget = -1;
  let oldImage = '';

  // 1. Cari Baris Produk Berdasarkan ID (Kolom A / Index 0)
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == form.id) {
      rowTarget = i + 1;
      oldImage = data[i][8]; // Simpan gambar lama
      break;
    }
  }

  if (rowTarget === -1) throw new Error("Produk tidak ditemukan/ID salah.");

  // 2. Cek Apakah Ada Gambar Baru Diupload?
  let finalImageUrl = oldImage;

  if (form.gambar && form.gambar.data) {
    try {
      const decoded = Utilities.base64Decode(form.gambar.data);
      const blob = Utilities.newBlob(decoded, form.gambar.mimeType, 'UPD-' + form.gambar.fileName);
      const folder = DriveApp.getFolderById(FOLDER_ID);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      finalImageUrl = "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w1000";
    } catch (e) {
      // Jika gagal upload, tetap lanjut simpan data teks, gambar pakai yang lama
      console.log("Gagal update gambar: " + e.message);
    }
  } else if (typeof form.gambar === 'string' && form.gambar !== '') {
      // Jika user memasukkan link manual baru
      finalImageUrl = form.gambar;
  }

  // 3. Update Baris (KECUALI STOK ISI & KOSONG)
  // Urutan Kolom: [0]ID, [1]Nama, [2]Jual, [3]Beli, [4]Isi(SKIP), [5]Kosong(SKIP), [6]SKU, [7]Kode, [8]Gambar
  
  sheet.getRange(rowTarget, 2).setValue(form.nama);       // Update Nama
  sheet.getRange(rowTarget, 3).setValue(form.hargaJual);  // Update Harga Jual
  sheet.getRange(rowTarget, 4).setValue(form.hargaBeli);  // Update Harga Beli
  // Kolom 5 & 6 (Stok) TIDAK DISENTUH
  sheet.getRange(rowTarget, 7).setValue(form.sku);        // Update SKU
  sheet.getRange(rowTarget, 8).setValue(form.kode);       // Update Kode Barcode
  sheet.getRange(rowTarget, 9).setValue(finalImageUrl);   // Update Gambar

  return "Produk Berhasil Diupdate!";
}

function hapusProduk(nama) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PRODUK');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] == nama) { sheet.deleteRow(i + 1); break; }
  }
}

// --- MODIFIKASI: TRANSAKSI & KASIR ---

// GANTI function simpanTransaksiBulk(dataTransaksi) dengan ini:

function simpanTransaksiBulk(dataTransaksi) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prodSheet = ss.getSheetByName('PRODUK');
  const trxSheet = ss.getSheetByName('TRANSAKSI');
  const keuSheet = ss.getSheetByName('KEUANGAN');
  
  const prodData = prodSheet.getDataRange().getValues();
  const idTrxMaster = 'KBA-' + Date.now();
  const waktu = new Date();
  
  let totalBelanja = 0;
  let summaryProduk = [];
  
  // Status Transaksi
  let statusTrx = (dataTransaksi.metode === 'Hutang') ? 'Belum Lunas' : 'Lunas';

  // 1. LOOP BARANG (Stok)
  dataTransaksi.items.forEach(item => {
    let itemFound = false;
    for (let i = 1; i < prodData.length; i++) {
      if (prodData[i][1] == item.produkNama) {
        let curIsi = Number(prodData[i][4]);
        let curKosong = Number(prodData[i][5]);
        
        if (curIsi < item.qty) throw new Error(`Stok ${item.produkNama} Habis! Sisa: ${curIsi}`);

        let newIsi = curIsi - item.qty;
        let newKosong = curKosong;
        if (item.tipe === 'Tukar (Refill)') {
           newKosong = curKosong + Number(item.qty); 
        }
        
        prodSheet.getRange(i + 1, 5).setValue(newIsi);
        prodSheet.getRange(i + 1, 6).setValue(newKosong);
        itemFound = true;
        break;
      }
    }
    
    if(!itemFound) throw new Error(`Produk ${item.produkNama} tidak ditemukan.`);

    // Catat Transaksi
    trxSheet.appendRow([
      idTrxMaster, waktu, dataTransaksi.pelanggan, item.produkNama, item.qty, 
      item.total, item.tipe, dataTransaksi.kasir, dataTransaksi.metode, 
      dataTransaksi.jatuhTempo, statusTrx 
    ]);

    totalBelanja += Number(item.total);
    summaryProduk.push(`${item.produkNama} (${item.qty})`);
  });

  // LOGIKA KEUANGAN
  if (dataTransaksi.metode !== 'Hutang') {
      keuSheet.appendRow([
        'FIN-' + idTrxMaster, 
        waktu, 
        'Pemasukan', 
        'Penjualan Gas', 
        totalBelanja, 
        `Penjualan: ${summaryProduk.join(', ')}`,
        dataTransaksi.metode 
      ]);
  }
  
  // [TAMBAHAN WAJIB] Paksa simpan detik ini juga
  SpreadsheetApp.flush(); 
  
  return "Transaksi Berhasil Disimpan!";
}

function getDataPiutang() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TRANSAKSI');
  if (!sheet) return [];
  const allData = sheet.getDataRange().getValues();
  if (allData.length < 2) return [];

  // Index Kolom (Sesuai Header):
  // 0:ID, 1:Waktu, 2:Pelanggan, ... 8:Metode_Bayar, 9:Jatuh_Tempo, 10:Status
  const idxMetode = 8;
  const idxJatuhTempo = 9;
  const idxStatus = 10;

  let grouped = {};

  for (let i = 1; i < allData.length; i++) {
    let row = allData[i];
    
    // 1. Cek Metode Bayar (Ambil semua yg 'Hutang', mau lunas atau belum)
    let metode = String(row[idxMetode]).trim();
    
    if (metode === 'Hutang') {
       let id = row[0];
       let status = String(row[idxStatus]).trim(); // Ambil status (Lunas/Belum)

       if(!grouped[id]) {
          let tglWaktu = (row[1] instanceof Date) ? row[1].toISOString() : String(row[1]);
          let tglTempo = (row[idxJatuhTempo] instanceof Date) ? row[idxJatuhTempo].toISOString() : String(row[idxJatuhTempo]);
          
          grouped[id] = {
             id: id,
             waktu: tglWaktu,      
             pelanggan: row[2],
             total: 0,
             jatuhTempo: tglTempo,
             status: status // Simpan statusnya
          };
       }
       grouped[id].total += Number(row[5]);
    }
  }
  
  // Return Array: [0]ID, [1]Waktu, [2]Pelanggan, [3]Total, [4]JatuhTempo, [5]Status
  // Kita urutkan: Yang "Belum Lunas" di atas, baru yang "Lunas" di bawah
  return Object.values(grouped).sort((a, b) => {
      if (a.status === b.status) {
          return new Date(b.waktu) - new Date(a.waktu); // Urut tanggal desc
      }
      return a.status === 'Belum Lunas' ? -1 : 1; // Prioritaskan Belum Lunas
  }).map(x => [x.id, x.waktu, x.pelanggan, x.total, x.jatuhTempo, x.status]);
}

// 2. Proses Pelunasan
function lunasiHutang(idTrx, totalBayar, namaPelanggan) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetTrx = ss.getSheetByName('TRANSAKSI');
  const sheetKeu = ss.getSheetByName('KEUANGAN');
  
  const dataTrx = sheetTrx.getDataRange().getValues();
  
  // A. Update Status di TRANSAKSI jadi 'Lunas'
  for(let i=1; i<dataTrx.length; i++) {
     if(dataTrx[i][0] == idTrx) {
        // Kolom K (Index 11, karena start dari 1 di sheet) -> Kolom ke-11
        sheetTrx.getRange(i+1, 11).setValue('Lunas'); 
     }
  }

  // B. Masukkan Uang ke KEUANGAN (Karena baru terima duit sekarang)
  sheetKeu.appendRow([
      'LUNAS-' + Date.now(), 
      new Date(), 
      'Pemasukan', 
      'Pelunasan Piutang', 
      totalBayar, 
      `Pelunasan Bon: ${namaPelanggan} (${idTrx})`
  ]);

  return "Hutang Berhasil Dilunasi & Masuk Kas!";
}

function getJumlahJatuhTempo() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TRANSAKSI');
  const data = sheet.getDataRange().getValues();
  const today = new Date();
  let count = 0;
  let uniqueIDs = []; // Supaya tidak double hitung item dalam 1 struk

  // Loop data transaksi
  for (let i = 1; i < data.length; i++) {
    let idTrx = data[i][0];
    let status = data[i][10]; // Kolom K (Status)
    let tglTempo = new Date(data[i][9]); // Kolom J (Jatuh Tempo)

    // Logika: Status Belum Lunas DAN Tanggal Tempo < Hari Ini (Sudah lewat)
    if (status === 'Belum Lunas' && tglTempo <= today && !uniqueIDs.includes(idTrx)) {
       count++;
       uniqueIDs.push(idTrx);
    }
  }
  return count;
}

function getRiwayatTransaksi() {
  const data = getData('TRANSAKSI');
  
  let grouped = {};
  data.forEach(row => {
    let id = row[0];
    let waktuStr = row[1] instanceof Date ? row[1].toISOString() : row[1];

    if (!grouped[id]) {
      grouped[id] = {
        id: id,
        waktu: waktuStr,
        pelanggan: row[2],
        kasir: row[7],
        // [PERBAIKAN DISINI] Tambahkan pembacaan kolom Metode & Jatuh Tempo
        metode: row[8],        // Kolom I (Index 8) -> Metode Bayar
        jatuhTempo: row[9],    // Kolom J (Index 9) -> Jatuh Tempo
        totalBayar: 0,  
        items: []       
      };
    }
    
    // ... (kode bawahnya tetap sama) ...
    grouped[id].items.push({
      produk: row[3],
      qty: row[4],
      hargaTotal: row[5],
      tipe: row[6],
      status: row[10]
    });

    grouped[id].totalBayar += Number(row[5]);
  });
  
  return Object.values(grouped).sort((a, b) => new Date(b.waktu) - new Date(a.waktu)).slice(0, 50);
}

// --- Code.gs ---

// 1. GET RIWAYAT PEMBELIAN (Grouping per ID)
function getRiwayatPembelian() {
  const data = getData('PEMBELIAN');
  let grouped = {};

  data.forEach(row => {
    let id = row[0];
    let waktuStr = row[1] instanceof Date ? row[1].toISOString() : row[1];

    if (!grouped[id]) {
      grouped[id] = {
        id: id,
        waktu: waktuStr,
        pelanggan: row[2], // Di sheet PEMBELIAN kolom ini adalah Supplier
        totalBayar: 0,
        items: []
      };
    }

    // Sheet PEMBELIAN: ID, Waktu, Supplier, Produk, Qty, Total, Metode
    grouped[id].items.push({
      produk: row[3],
      qty: row[4],
      hargaTotal: row[5],
      tipe: 'Stok Masuk', // Default tipe
      status: 'Sukses' 
    });
    
    grouped[id].totalBayar += Number(row[5]);
  });

  return Object.values(grouped).sort((a, b) => new Date(b.waktu) - new Date(a.waktu)).slice(0, 50);
}

// 2. FUNGSI RETUR BARU (Support Partial & Jenis Transaksi)
function prosesReturBaru(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prodSheet = ss.getSheetByName('PRODUK');
  const keuSheet = ss.getSheetByName('KEUANGAN');
  
  // Tentukan Sheet Target berdasarkan jenis
  const targetSheetName = payload.jenis === 'JUAL' ? 'TRANSAKSI' : 'PEMBELIAN';
  const trxSheet = ss.getSheetByName(targetSheetName);
  const trxData = trxSheet.getDataRange().getValues();
  const prodData = prodSheet.getDataRange().getValues();

  let totalRefund = 0;
  let logItem = [];

  // Loop item yang diretur
  payload.items.forEach(returItem => {
    if(returItem.qtyRetur > 0) {
      
      // A. UPDATE STOK PRODUK
      for (let i = 1; i < prodData.length; i++) {
        if (prodData[i][1] == returItem.produk) {
           let curIsi = Number(prodData[i][4]);
           let curKosong = Number(prodData[i][5]);
           
           if(payload.jenis === 'JUAL') {
              // Retur Penjualan: Stok Isi KEMBALI (+), Stok Kosong BERKURANG (karena sebelumnya tukar)
              prodSheet.getRange(i+1, 5).setValue(curIsi + returItem.qtyRetur);
              // Cek jika itu refill, tabung kosong dikembalikan ke pelanggan (stok kita berkurang)
              if(returItem.tipe && returItem.tipe.includes('Refill')) {
                 prodSheet.getRange(i+1, 6).setValue(curKosong - returItem.qtyRetur);
              }
           } else {
              // Retur Pembelian: Stok Isi BERKURANG (-) (Balikin ke supplier)
              prodSheet.getRange(i+1, 5).setValue(curIsi - returItem.qtyRetur);
              // Jika beli tukar tabung, stok kosong kita bertambah lagi (dibalikin supplier)
               // (Sederhananya kita kurangi stok isi saja dulu untuk keamanan)
           }
           break;
        }
      }

      // B. UPDATE STATUS TRANSAKSI (Tandai Retur)
      // Cari baris transaksi spesifik
      for(let i=1; i<trxData.length; i++) {
         if(trxData[i][0] == payload.idTrx && trxData[i][3] == returItem.produk) {
             // Opsional: Bisa update kolom qty atau tambah catatan "Retur Partial"
             // Disini kita biarkan record asli, tapi catat di Keuangan sebagai pengurang
         }
      }
      
      totalRefund += (returItem.hargaSatuan * returItem.qtyRetur);
      logItem.push(`${returItem.produk} (x${returItem.qtyRetur})`);
    }
  });

  // C. CATAT DI KEUANGAN (Balance)
  if(totalRefund > 0) {
     if(payload.jenis === 'JUAL') {
        // Retur Jual = Uang Keluar (Refund ke Pelanggan)
        keuSheet.appendRow(['RET-' + Date.now(), new Date(), 'Pengeluaran', 'Retur Penjualan', totalRefund, `Retur TRX: ${payload.idTrx}. ${payload.alasan}`]);
     } else {
        // Retur Beli = Uang Masuk (Refund dari Supplier)
        keuSheet.appendRow(['RET-' + Date.now(), new Date(), 'Pemasukan', 'Retur Pembelian', totalRefund, `Retur BELI: ${payload.idTrx}. ${payload.alasan}`]);
     }
  }

  return "Retur Berhasil Diproses!";
}

function simpanPelanggan(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PELANGGAN');
  
  // EDIT MODE
  if(form.id) { 
    const data = sheet.getDataRange().getValues();
    for(let i=1; i<data.length; i++) {
      if(data[i][0] == form.id) {
        // Update: Nama, Perusahaan, HP, Alamat
        sheet.getRange(i+1, 2, 1, 4).setValues([[form.nama, form.pt, form.hp, form.alamat]]);
        return "Data Pelanggan Diupdate";
      }
    }
  }
  
  // BARU MODE
  sheet.appendRow(['CUST-' + Date.now(), form.nama, form.pt, form.hp, form.alamat]);
  return "Pelanggan Baru Disimpan";
}

function hapusPelanggan(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PELANGGAN');
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(data[i][0] == id) { 
      sheet.deleteRow(i+1); 
      return "Pelanggan Dihapus";
    }
  }
}

// Fungsi bantu untuk mengambil List Pelanggan di Kasir
function getListPelanggan() {
  return getData('PELANGGAN'); // <--- WAJIB ADA 'return'
}

// 3. Hapus / Retur Transaksi
function prosesRetur(idTrx, produkNama, qty, tipe, mode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prodSheet = ss.getSheetByName('PRODUK');
  const trxSheet = ss.getSheetByName('TRANSAKSI');
  const keuSheet = ss.getSheetByName('KEUANGAN');
  
  // A. KEMBALIKAN STOK
  const prodData = prodSheet.getDataRange().getValues();
  for (let i = 1; i < prodData.length; i++) {
    if (prodData[i][1] == produkNama) {
       let curIsi = Number(prodData[i][4]);
       let curKosong = Number(prodData[i][5]);
       
       // Logic Retur: Kembalikan Stok Isi, Kurangi Stok Kosong (jika refill)
       prodSheet.getRange(i + 1, 5).setValue(curIsi + Number(qty));
       
       if(tipe === 'Tukar (Refill)') {
          prodSheet.getRange(i + 1, 6).setValue(curKosong - Number(qty));
       }
       break;
    }
  }

  // B. UPDATE STATUS TRANSAKSI & KEUANGAN
  // Cari baris transaksi
  const trxData = trxSheet.getDataRange().getValues();
  let nominalRefund = 0;

  for(let i=1; i<trxData.length; i++) {
    // Mencocokkan ID, Produk, dan memastikan belum diretur
    if(trxData[i][0] == idTrx && trxData[i][3] == produkNama && trxData[i][8] != 'Retur') {
       if(mode === 'FULL') {
         trxSheet.deleteRow(i+1); // Hapus baris permanen jika mau bersih
         // Atau tandai: trxSheet.getRange(i+1, 9).setValue('Retur');
       } else {
         trxSheet.getRange(i+1, 9).setValue('Retur Item');
       }
       nominalRefund = trxData[i][5]; // Ambil total harga item tsb
       break;
    }
  }

  // C. CATAT PENGELUARAN REFUND DI KEUANGAN (Agar Balance)
  keuSheet.appendRow([
      'REFUND-' + Date.now(), new Date(), 
      'Pengeluaran', 'Retur Penjualan', 
      nominalRefund, `Retur: ${produkNama} (${idTrx})`
  ]);

  return "Berhasil Retur/Hapus";
}

// --- TAMBAHAN: SIMPAN PEMBELIAN BULK (KERANJANG) ---
function simpanPembelianBulk(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetBeli = ss.getSheetByName('PEMBELIAN');
  const sheetProd = ss.getSheetByName('PRODUK');
  const sheetKeu = ss.getSheetByName('KEUANGAN');
  
  const idBeliMaster = 'BELI-' + Date.now();
  const waktu = new Date();
  const prodData = sheetProd.getDataRange().getValues();
  
  let summaryItem = [];

  // Loop setiap item di keranjang beli
  data.items.forEach(item => {
    // 1. Catat di Sheet PEMBELIAN
    // Format: ID, Waktu, Supplier, Produk, Qty, Total, Metode
    sheetBeli.appendRow([
      idBeliMaster, 
      waktu, 
      data.supplier, 
      item.produk, 
      item.qty, 
      item.total, 
      'Tunai'
    ]);

    // 2. Update Stok di Sheet PRODUK
    for (let i = 1; i < prodData.length; i++) {
      if (prodData[i][1] == item.produk) {
        let curIsi = Number(prodData[i][4]);
        let curKosong = Number(prodData[i][5]);
        
        // Stok Isi Bertambah (+)
        sheetProd.getRange(i + 1, 5).setValue(curIsi + Number(item.qty));
        
        // Jika Tukar Tabung, Stok Kosong Berkurang (-)
        if(item.isTukar) {
           sheetProd.getRange(i + 1, 6).setValue(curKosong - Number(item.qty));
        }
        break;
      }
    }
    summaryItem.push(`${item.produk} (x${item.qty})`);
  });

  // 3. Catat di KEUANGAN (Satu baris total pengeluaran)
  sheetKeu.appendRow([
    'OUT-' + Date.now(), 
    waktu, 
    'Pengeluaran', 
    'Pembelian Stok', 
    data.grandTotal, 
    `Beli Stok: ${summaryItem.join(', ')}`
  ]);

  return "Stok Berhasil Ditambahkan!";
}

// --- PEMBELIAN (BELI) ---
function tambahSupplier(form) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SUPPLIER').appendRow(['SUP-' + Date.now(), form.nama, form.hp, form.alamat]);
}

function simpanPembelian(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prodSheet = ss.getSheetByName('PRODUK');
  
  // 1. Catat Beli
  ss.getSheetByName('PEMBELIAN').appendRow(['BELI-' + Date.now(), new Date(), data.supplier, data.produk, data.qty, data.total, data.metode]);
  
  // 2. Update Stok
  const prodData = prodSheet.getDataRange().getValues();
  for (let i = 1; i < prodData.length; i++) {
    if (prodData[i][1] == data.produk) {
      let curIsi = Number(prodData[i][4]);
      let curKosong = Number(prodData[i][5]);
      
      prodSheet.getRange(i + 1, 5).setValue(curIsi + Number(data.qty)); // Stok Isi Nambah
      if(data.isTukar) {
        prodSheet.getRange(i + 1, 6).setValue(curKosong - Number(data.qty)); // Stok Kosong Berkurang
      }
      break;
    }
  }
  
  // 3. Catat Pengeluaran
  ss.getSheetByName('KEUANGAN').appendRow(['OUT-' + Date.now(), new Date(), 'Pengeluaran', 'Pembelian Stok', data.total, `Beli ${data.produk}`]);
}

// --- KEUANGAN ---
function getKategori() {
  return getData('KATEGORI').map(r => r[0]);
}

function tambahKategori(nama) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KATEGORI').appendRow([nama]);
}

// --- [UPDATE] Simpan Keuangan dengan Kolom AKUN ---
function simpanKeuangan(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KEUANGAN');
  const tglInput = new Date(form.tanggal);

  // Tambahkan Header jika belum ada (Update Header Lama)
  if(sheet.getLastColumn() < 7) {
     sheet.getRange(1, 7).setValue('Akun');
  }

  // LOGIKA EDIT
  if (form.id && !form.id.includes('MANUAL')) { 
     // ... (Kode edit lama disesuaikan jika perlu, disini kita fokus Input Baru dulu)
  }

  // LOGIKA BARU
  const newId = 'MANUAL-' + Date.now();
  sheet.appendRow([
      newId, 
      tglInput, 
      form.jenis, 
      form.kategori, 
      form.nominal, 
      form.keterangan,
      form.akun // [BARU] Simpan Nama Akun
  ]);
  
  return { status: 'success', data: { id: newId, ...form } };
}

// 2. Update Simpan Akun Baru (Tambah parameter norek)
function simpanAkunBaru(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AKUN_KAS');
  const id = 'ACC-' + Date.now();
  
  // [UPDATE] Urutan simpan: ID, Nama, NoRek, Tipe, Saldo
  sheet.appendRow([id, form.nama, "'" + form.norek, form.tipe, form.saldo]); 
  // Note: Ditambah tanda petik satu (') di depan norek agar angka 0 tidak hilang
  
  return "Akun Berhasil Ditambahkan!";
}

function hapusAkun(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AKUN_KAS');
  const data = sheet.getDataRange().getValues();
  
  // Mulai loop dari 1 (skip header)
  for(let i=1; i<data.length; i++) {
     if(data[i][0] == id) {
        sheet.deleteRow(i+1);
        return "Akun Dihapus.";
     }
  }
}

// --- [BARU] Fitur Transfer Saldo Antar Akun ---
function prosesTransferSaldo(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KEUANGAN');
  const waktu = new Date();
  const idTrx = 'TRF-' + Date.now();

  // Konsep Transfer:
  // 1. Catat PENGELUARAN di Akun Asal
  sheet.appendRow([
     idTrx + '-OUT', waktu, 'Pengeluaran', 'Transfer Keluar', form.nominal, 
     `Transfer ke ${form.akunTujuan} (${form.ket})`, form.akunAsal
  ]);

  // 2. Catat PEMASUKAN di Akun Tujuan
  sheet.appendRow([
     idTrx + '-IN', waktu, 'Pemasukan', 'Transfer Masuk', form.nominal, 
     `Terima dari ${form.akunAsal} (${form.ket})`, form.akunTujuan
  ]);

  return "Transfer Berhasil!";
}

// --- BARU: HAPUS KEUANGAN ---
function hapusKeuangan(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KEUANGAN');
  const data = sheet.getDataRange().getValues();
  
  // Safety: Cek lagi di server, hanya boleh hapus yang MANUAL
  if(!String(id).includes('MANUAL')) {
     throw new Error("Data sistem (Otomatis) tidak boleh dihapus dari sini!");
  }

  for(let i = 1; i < data.length; i++) {
    if(data[i][0] == id) {
       sheet.deleteRow(i+1);
       return "Data Dihapus";
    }
  }
  throw new Error("ID tidak ditemukan");
}

// --- SDM: KARYAWAN ---
function simpanKaryawan(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KARYAWAN');
  
  if(form.id) { // Edit Mode
    const data = sheet.getDataRange().getValues();
    for(let i=1; i<data.length; i++) {
      if(data[i][0] == form.id) {
        sheet.getRange(i+1, 2, 1, 4).setValues([[form.nama, form.hp, form.gaji, form.bonus]]);
        return "Data Updated";
      }
    }
  } 
  // New Mode
  sheet.appendRow(['KRY-' + Date.now(), form.nama, form.hp, form.gaji, form.bonus, 'Aktif']);
  return "Karyawan Baru Disimpan";
}

function hapusKaryawan(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KARYAWAN');
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(data[i][0] == id) { sheet.deleteRow(i+1); return; }
  }
}

// --- SDM: KASBON ---
function simpanKasbon(form) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KASBON')
    .appendRow(['KSB-' + Date.now(), new Date(), form.nama, form.nominal, form.ket, 'Belum Lunas']);
  return "Kasbon Dicatat";
}

// --- SDM: PAYROLL LOGIC ---
function getDataPayroll() {
  const karyawan = getData('KARYAWAN');
  const kasbonData = getData('KASBON');
  
  let result = karyawan.map(k => {
    let nama = k[1];
    let gaji = Number(k[3]);
    let bonusSet = Number(k[4]);
    
    // Hitung Kasbon Belum Lunas
    let totalKasbon = 0;
    kasbonData.forEach(ksb => {
      if(ksb[2] === nama && ksb[5] === 'Belum Lunas') {
        totalKasbon += Number(ksb[3]);
      }
    });
    
    // Bonus Sementara (Dummy: 0), nanti bisa dikembangkan hitung jumlah penjualan kasir
    let totalBonus = 0; 

    return {
      id: k[0],
      nama: nama,
      gaji: gaji,
      bonus: totalBonus,
      kasbon: totalKasbon,
      total: gaji + totalBonus - totalKasbon
    };
  });
  return result;
}

function prosesPayrollFinal(listGaji) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const keuSheet = ss.getSheetByName('KEUANGAN');
  const kasbonSheet = ss.getSheetByName('KASBON');
  const kasbonData = kasbonSheet.getDataRange().getValues();
  
  let totalKeluar = 0;
  
  listGaji.forEach(g => {
    totalKeluar += Number(g.total);
    // Lunaskan Kasbon
    if(g.kasbon > 0) {
      for(let i=1; i<kasbonData.length; i++) {
        if(kasbonData[i][2] == g.nama && kasbonData[i][5] == 'Belum Lunas') {
          kasbonSheet.getRange(i+1, 6).setValue('Lunas (Potong Gaji)');
        }
      }
    }
  });
  
  keuSheet.appendRow(['PAY-' + Date.now(), new Date(), 'Pengeluaran', 'Gaji Karyawan', totalKeluar, 'Payroll Periode Ini']);
  return "Gaji Dicairkan & Kasbon Terpotong.";
}

function TES_BIKIN_FILE() {
  // ID Folder Anda
  const id = '15hiLtvusofF2OJpXVq8lJkePbmqVIuPM'; 
  
  const folder = DriveApp.getFolderById(id);
  
  // Kita coba bikin file teks kosong beneran untuk mancing izin "Write"
  folder.createFile('Tes_Izin.txt', 'Halo, ini tes izin upload.');
  
  console.log("Sukses! Izin Upload sudah aktif.");
}
