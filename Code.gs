/**
 * ============================================================
 * KOPERASI GATRA MARANATHA - GOOGLE APPS SCRIPT
 * Aplikasi Manajemen Koperasi Digital
 * ============================================================
 */

// ==================== KONFIGURASI GLOBAL ====================
const SPREADSHEET_ID = '1RaYcQCBAhqhIV-0u2f5GDhKcV8B7CL2FHNKpGxEaglI';
const SHEET_NAMES = {
  BARANG: 'Barang',
  ANGGOTA: 'Anggota',
  PENJUALAN: 'Penjualan',
  STOK_LOG: 'Stok_Log',
  PENGURUS: 'Pengurus'
};

// ==================== SHEET UTILITIES ====================

/**
 * Mendapatkan sheet berdasarkan nama
 * @param {string} name - Nama sheet
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} Sheet object
 */
function getSheetByName(name) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss.getSheetByName(name);
}

/**
 * Mendapatkan spreadsheet aktif
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} Spreadsheet object
 */
function getActiveSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// ==================== WEB APP HANDLER ====================

/**
 * Handler untuk request GET (Web App)
 * @returns {GoogleAppsScript.HTML.HtmlOutput} Halaman HTML aplikasi
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Aplikasi Koperasi')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==================== ANGGOTA FUNCTIONS ====================

/**
 * Validasi anggota berdasarkan NIK
 * @param {string} nik - Nomor Induk Kependudukan
 * @returns {Object} Status validasi anggota
 */
function validasiAnggota(nik) {
  try {
    var sheet = getSheetByName(SHEET_NAMES.ANGGOTA);
    if (!sheet) {
      return { status: false, pesan: 'Sheet Anggota tidak ditemukan!' };
    }
    
    var data = sheet.getDataRange().getValues();
    var nikInput = nik.toString().trim();
    
    // Loop dari baris 1 (skip header)
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      
      var nikDb = data[i][0].toString().trim();
      var statusAnggota = data[i][4] ? data[i][4].toString().trim() : "";
      
      if (nikDb === nikInput) {
        if (statusAnggota.toLowerCase() === 'aktif') {
          return { 
            status: true, 
            nama: data[i][1].toString(),
            noHp: data[i][3] ? data[i][3].toString().trim() : "" 
          };
        } else {
          return { 
            status: false, 
            pesan: 'Anggota ditemukan, tapi status tidak Aktif' 
          };
        }
      }
    }
    
    return { 
      status: false, 
      pesan: 'NIK tidak terdaftar sebagai Anggota Anda Ditolak' 
    };
  } catch(e) {
    return { 
      status: false, 
      pesan: 'Error sistem: ' + e.message 
    };
  }
}

// ==================== BARANG FUNCTIONS ====================

/**
 * Mendapatkan daftar semua barang dari sheet Barang
 * @returns {Array} Array object barang
 */
function getDaftarBarang() {
  try {
    var sheet = getSheetByName(SHEET_NAMES.BARANG);
    if (!sheet) {
      throw new Error("Sheet dengan nama 'Barang' tidak ditemukan. Pastikan nama tab sudah benar!");
    }
    
    var data = sheet.getDataRange().getValues();
    var listBarang = [];
    
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      
      // Parse harga dan stok
      var hargaMentah = data[i][3] != null ? data[i][3].toString() : "0";
      var stokMentah = data[i][4] != null ? data[i][4].toString() : "0";
      
      var hargaBersih = parseFloat(hargaMentah.replace(/[^0-9.-]+/g, "")) || 0;
      var stokBersih = parseInt(stokMentah.replace(/[^0-9.-]+/g, "")) || 0;
      
      // Ambil data Barcode dari Kolom G (indeks ke-6)
      var barcodePabrik = "";
      if (data[i][6] !== null && data[i][6] !== undefined) {
        if (typeof data[i][6] === 'number') {
          barcodePabrik = data[i][6].toFixed(0).toString().trim();
        } else {
          barcodePabrik = data[i][6].toString().trim();
        }
      }
      
      listBarang.push({
        kode: data[i][0].toString(),
        nama: data[i][1] ? data[i][1].toString() : "Tanpa Nama",
        kategori: data[i][2] ? data[i][2].toString() : "-",
        harga: hargaBersih,
        stok: stokBersih,
        gambar: data[i][5] ? data[i][5].toString().trim() : "",
        barcode: barcodePabrik
      });
    }
    
    return listBarang;
  } catch (e) {
    throw new Error(e.message);
  }
}

// ==================== TRANSAKSI FUNCTIONS ====================

/**
 * Memproses transaksi pembelian
 * @param {Array} dataKeranjang - Array item keranjang
 * @param {string} nikPembeli - NIK pembeli
 * @param {string} namaPembeli - Nama pembeli
 * @returns {Object} Status transaksi
 */
function prosesTransaksi(dataKeranjang, nikPembeli, namaPembeli) {
  try {
    var sheetPenjualan = getSheetByName(SHEET_NAMES.PENJUALAN);
    var sheetBarang = getSheetByName(SHEET_NAMES.BARANG);
    var nota = 'NOTA-' + new Date().getTime();
    var tanggal = new Date();
    
    var barangData = sheetBarang.getDataRange().getValues();
    
    dataKeranjang.forEach(function(item) {
      // Simpan ke sheet Penjualan
      sheetPenjualan.appendRow([
        nota,                      // Kolom A: No Nota
        tanggal,                   // Kolom B: Tanggal
        nikPembeli,                // Kolom C: NIK Pembeli
        namaPembeli,               // Kolom D: Nama Pembeli
        item.kode,                 // Kolom E: Kode Barang
        item.qty,                  // Kolom F: Qty
        item.harga * item.qty,     // Kolom G: Total
        "Mandiri"                  // Kolom H: Tipe
      ]);
      
      // Update stok barang
      for (var i = 1; i < barangData.length; i++) {
        if (barangData[i][0] == item.kode) {
          var stokLama = barangData[i][4];
          var stokBaru = stokLama - item.qty;
          sheetBarang.getRange(i + 1, 5).setValue(stokBaru);
          break;
        }
      }
    });
    
    return { status: 'Sukses', nota: nota };
  } catch(e) {
    return { status: 'Gagal', pesan: e.message };
  }
}

/**
 * Menyimpan nota transaksi ke spreadsheet
 * @param {Object} dataTransaksi - Data transaksi lengkap
 * @returns {Object} Status penyimpanan
 */
function simpanNotaKeSpreadsheet(dataTransaksi) {
  try {
    var ss = getActiveSpreadsheet();
    
    var sheetPenjualan = ss.getSheetByName(SHEET_NAMES.PENJUALAN);
    if (!sheetPenjualan) {
      return { success: false, message: 'Sheet Penjualan tidak ditemukan!' };
    }
    
    var waktu = new Date();
    var noNota = "GATRA-" + waktu.getTime();
    
    var namaPembeli = dataTransaksi.nama ? dataTransaksi.nama.toString().trim() : "Pelanggan";
    
    // Format detail barang
    var detailBarang = dataTransaksi.keranjang.map(function(item) {
      return item.nama + " (" + item.qty + "x)";
    }).join(", ");
    
    // Hitung total qty
    var totalQty = 0;
    dataTransaksi.keranjang.forEach(function(item) {
      totalQty += item.qty;
    });
    
    // Simpan ke sheet Penjualan
    sheetPenjualan.appendRow([
      noNota,                   // No Nota
      waktu,                    // Waktu
      dataTransaksi.nik,        // NIK
      namaPembeli,              // Nama Pembeli
      detailBarang,             // Detail Barang
      totalQty,                 // Total Qty
      dataTransaksi.total,      // Total Harga
      dataTransaksi.pengurus    // Pengurus
    ]);
    
    // Update stok barang
    var sheetBarang = ss.getSheetByName(SHEET_NAMES.BARANG);
    if (sheetBarang) {
      var dataBarang = sheetBarang.getDataRange().getValues();
      
      dataTransaksi.keranjang.forEach(function(itemKeranjang) {
        for (var i = 1; i < dataBarang.length; i++) {
          var kodeBarangDb = dataBarang[i][0].toString().trim();
          
          if (kodeBarangDb === itemKeranjang.kode.toString().trim()) {
            var stokSekarang = parseInt(dataBarang[i][4]) || 0;
            var stokBaru = stokSekarang - itemKeranjang.qty;
            
            sheetBarang.getRange(i + 1, 5).setValue(stokBaru);
            break;
          }
        }
      });
    }
    
    return { success: true, noNota: noNota };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ==================== STOK LOG FUNCTIONS ====================

/**
 * Input log perubahan stok
 * @param {string} kodeBarang - Kode barang
 * @param {string} jenis - Jenis perubahan (Masuk/Keluar)
 * @param {number} jumlah - Jumlah perubahan
 * @param {string} keterangan - Keterangan perubahan
 * @returns {string} Status operasi
 */
function inputStokLog(kodeBarang, jenis, jumlah, keterangan) {
  var sheetLog = getSheetByName(SHEET_NAMES.STOK_LOG);
  var sheetBarang = getSheetByName(SHEET_NAMES.BARANG);
  var idLog = 'LOG-' + new Date().getTime();
  
  // Simpan log
  sheetLog.appendRow([idLog, new Date(), kodeBarang, jenis, jumlah, keterangan]);
  
  // Update stok barang
  var barangData = sheetBarang.getDataRange().getValues();
  for (var i = 1; i < barangData.length; i++) {
    if (barangData[i][0] == kodeBarang) {
      var stokLama = parseInt(barangData[i][4]);
      var perubahan = parseInt(jumlah);
      var stokBaru = (jenis === 'Masuk') ? (stokLama + perubahan) : (stokLama - perubahan);
      sheetBarang.getRange(i + 1, 5).setValue(stokBaru);
      break;
    }
  }
  
  return 'Stok berhasil diperbarui';
}

// ==================== PENGURUS FUNCTIONS ====================

/**
 * Validasi PIN pengurus
 * @param {string} pinInput - PIN yang diinput
 * @returns {Object} Status validasi
 */
function validasiPinPengurusSistemDirect(pinInput) {
  try {
    var sheet = getSheetByName(SHEET_NAMES.PENGURUS);
    if (!sheet) {
      return { success: false, message: "Sheet 'Pengurus' tidak ditemukan!" };
    }
    
    var data = sheet.getDataRange().getValues();
    var pinDicari = pinInput.toString().trim();
    
    for (var i = 1; i < data.length; i++) {
      var pinDb = data[i][2] ? data[i][2].toString().trim() : "";
      
      if (pinDb === pinDicari) {
        var namaDb = data[i][0] ? data[i][0].toString().trim() : "Pengurus";
        var jabatanDb = data[i][1] ? data[i][1].toString().trim() : "";
        return { 
          success: true, 
          message: "PIN Valid!", 
          namaLengkap: namaDb + " (" + jabatanDb + ")" 
        };
      }
    }
    
    return { 
      success: false, 
      message: "PIN yang Anda masukkan salah atau tidak terdaftar Anda Ditolak!" 
    };
  } catch (e) {
    return { 
      success: false, 
      message: "Terjadi kesalahan sistem: " + e.toString() 
    };
  }
}

/**
 * Mendapatkan daftar lengkap pengurus
 * @returns {Array} Array object pengurus
 */
function getDaftarPengurusLengkap() {
  try {
    var sheet = getSheetByName(SHEET_NAMES.PENGURUS);
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    var hasil = [];
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        hasil.push({
          nama: data[i][0].toString().trim(),
          jabatan: data[i][1] ? data[i][1].toString().trim() : "",
          pin: data[i][2] ? data[i][2].toString().trim() : ""
        });
      }
    }
    return hasil;
  } catch (e) {
    return [];
  }
}

// ==================== LAPORAN FUNCTIONS ====================

/**
 * Mendapatkan laporan koperasi
 * @returns {Object} Data laporan
 */
function getLaporanKoperasi() {
  try {
    var sheetPenjualan = getSheetByName(SHEET_NAMES.PENJUALAN);
    var sheetBarang = getSheetByName(SHEET_NAMES.BARANG);
    
    var dataPenjualan = sheetPenjualan.getDataRange().getValues();
    var dataBarang = sheetBarang.getDataRange().getValues();
    
    var totalPendapatan = 0;
    var totalTransaksi = 0;
    var totalItemTerjual = 0;
    var notaUnik = {};
    var barangMenipis = [];
    
    // Hitung dari data penjualan
    for (var i = 1; i < dataPenjualan.length; i++) {
      if (!dataPenjualan[i][0]) continue;
      
      totalPendapatan += parseFloat(dataPenjualan[i][6]) || 0;
      totalItemTerjual += parseInt(dataPenjualan[i][5]) || 0;
      
      // Hitung nota unik
      if (!notaUnik[dataPenjualan[i][0]]) {
        notaUnik[dataPenjualan[i][0]] = true;
        totalTransaksi++;
      }
    }
    
    // Cek stok menipis
    for (var j = 1; j < dataBarang.length; j++) {
      if (!dataBarang[j][0]) continue;
      var stok = parseInt(dataBarang[j][4]) || 0;
      if (stok <= 10) {
        barangMenipis.push({
          nama: dataBarang[j][1].toString(),
          stok: stok
        });
      }
    }
    
    return {
      totalPendapatan: totalPendapatan,
      totalTransaksi: totalTransaksi,
      totalItemTerjual: totalItemTerjual,
      barangMenipis: barangMenipis
    };
  } catch (e) {
    throw new Error(e.message);
  }
}

// ==================== RETUR FUNCTIONS ====================

/**
 * Memproses retur barang
 * @param {string} kodeBarang - Kode barang
 * @param {number} qtyRetur - Jumlah retur
 * @param {string} alasanRetur - Alasan retur
 * @param {string} jenisRetur - Jenis retur (Masuk/Keluar)
 * @returns {Object} Status retur
 */
function prosesReturBarang(kodeBarang, qtyRetur, alasanRetur, jenisRetur) {
  try {
    var sheetLog = getSheetByName(SHEET_NAMES.STOK_LOG);
    var sheetBarang = getSheetByName(SHEET_NAMES.BARANG);
    var tanggal = new Date();
    var idRetur = 'RTR-' + tanggal.getTime();
    
    // Ambil data barang untuk update stok
    var barangData = sheetBarang.getDataRange().getValues();
    var namaBarang = "";
    var stokDitemukan = false;
    
    for (var i = 1; i < barangData.length; i++) {
      if (barangData[i][0] == kodeBarang) {
        namaBarang = barangData[i][1];
        var stokLama = Number(barangData[i][4]);
        var qty = Number(qtyRetur);
        
        var stokBaru = (jenisRetur === 'Masuk') ? (stokLama + qty) : (stokLama - qty);
        sheetBarang.getRange(i + 1, 5).setValue(stokBaru);
        stokDitemukan = true;
        break;
      }
    }
    
    if (!stokDitemukan) {
      return { status: 'Gagal', pesan: 'Kode barang tidak ditemukan di database!' };
    }
    
    // Simpan log retur
    sheetLog.appendRow([
      idRetur,        // ID Log
      tanggal,        // Tanggal
      kodeBarang,     // Kode Barang
      namaBarang,     // Nama Barang
      jenisRetur,     // Jenis Retur
      qtyRetur,       // Qty
      alasanRetur     // Alasan
    ]);
    
    return { 
      status: 'Sukses', 
      pesan: 'Retur ' + namaBarang + ' berhasil dicatat.' 
    };
  } catch(e) {
    return { 
      status: 'Gagal', 
      pesan: 'Error sistem: ' + e.message 
    };
  }
}

// ==================== OPNAME FUNCTIONS ====================

/**
 * Update stok opname berdasarkan barcode
 * @param {string} kodeBarang - Kode/barcode barang
 * @param {number} stokBaru - Stok baru yang diinput
 * @returns {Object} Status update
 */
function updateStokOpnameDirect(kodeBarang, stokBaru) {
  try {
    var sheet = getSheetByName(SHEET_NAMES.BARANG);
    if (!sheet) {
      return { success: false, message: "Sheet 'Barang' tidak ditemukan!" };
    }
    
    var data = sheet.getDataRange().getValues();
    var kodeDicari = kodeBarang.toString().trim();
    var barisDitemukan = -1;
    
    for (var i = 1; i < data.length; i++) {
      var nilaiDb = data[i][6]; // Kolom G (Barcode)
      var kodeDb = "";
      
      if (nilaiDb !== null && nilaiDb !== undefined) {
        if (typeof nilaiDb === 'number') {
          kodeDb = nilaiDb.toFixed(0).toString().trim();
        } else {
          kodeDb = nilaiDb.toString().trim();
        }
      }
      
      if (kodeDb === kodeDicari && kodeDb !== "") {
        barisDitemukan = i + 1;
        var namaBarang = data[i][1] ? data[i][1].toString().trim() : "Barang Tanpa Nama";
        
        // Ambil stok lama dan jumlahkan
        var stokLama = data[i][4] ? Number(data[i][4]) : 0;
        var totalStokTerbaru = stokLama + Number(stokBaru);
        
        // Simpan hasil penjumlahan ke Kolom E
        sheet.getRange(barisDitemukan, 5).setValue(totalStokTerbaru);
        
        return { 
          success: true, 
          message: "Stok " + namaBarang + " berhasil ditambah! (" + stokLama + " + " + stokBaru + " = " + totalStokTerbaru + " pcs)" 
        };
      }
    }
    
    return { 
      success: false, 
      message: "Kode barcode '" + kodeDicari + "' belum terdaftar!" 
    };
  } catch (e) {
    return { 
      success: false, 
      message: "Gagal update stok: " + e.toString() 
    };
  }
}

// ==================== PDF GENERATOR ====================

/**
 * Membuat nota PDF
 * @param {string} notaId - ID nota
 * @param {string} nikPembeli - NIK pembeli
 * @param {string} namaPembeli - Nama pembeli
 * @param {Array} dataKeranjang - Data keranjang
 * @returns {string|null} URL file PDF atau null jika gagal
 */
function buatNotaPDF(notaId, nikPembeli, namaPembeli, dataKeranjang) {
  try {
    var tanggal = new Date().toLocaleDateString('id-ID', { 
      year: 'numeric', 
      month: 'long', 
      day: 'numeric', 
      hour: '2-digit', 
      minute: '2-digit' 
    });
    
    var htmlContent = `
      <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; margin: 20px; color: #333; }
          .header { text-align: center; border-bottom: 2px solid #28a745; padding-bottom: 10px; }
          .info { margin: 20px 0; font-size: 14px; }
          table { width: 100%; border-collapse: collapse; margin-top: 10px; }
          th, td { border: 1px solid #ddd; padding: 10px; text-align: left; }
          th { background-color: #28a745; color: white; }
          .total { text-align: right; font-weight: bold; font-size: 16px; margin-top: 15px; }
          .footer { margin-top: 40px; text-align: center; font-size: 12px; color: #777; }
        </style>
      </head>
      <body>
        <div class="header">
          <h2>KOPERASI GATRA MARANATHA</h2>
          <p>Nota Transaksi Digital Belanja Mandiri</p>
        </div>
        <div class="info">
          <b>No. Nota:</b> ${notaId}<br>
          <b>Tanggal:</b> ${tanggal}<br>
          <b>Pembeli:</b> ${namaPembeli} (${nikPembeli})
        </div>
        <table>
          <thead>
            <tr>
              <th>Nama Barang</th>
              <th>Harga Satuan</th>
              <th>Qty</th>
              <th>Total</th>
            </tr>
          </thead>
          <tbody>
    `;
    
    var totalBelanja = 0;
    dataKeranjang.forEach(function(item) {
      var subtotal = item.harga * item.qty;
      totalBelanja += subtotal;
      htmlContent += `
        <tr>
          <td>${item.nama}</td>
          <td>Rp ${item.harga.toLocaleString('id-ID')}</td>
          <td>${item.qty}</td>
          <td>Rp ${subtotal.toLocaleString('id-ID')}</td>
        </tr>
      `;
    });
    
    htmlContent += `
          </tbody>
        </table>
        <div class="total">Total Bayar: Rp ${totalBelanja.toLocaleString('id-ID')}</div>
        <div class="footer">
          <p>Terima kasih telah berbelanja di Koperasi GATRA MARANATHA.</p>
          <p>Simpan nota ini sebagai bukti pembayaran yang sah.</p>
        </div>
      </body>
      </html>
    `;
    
    var blob = Utilities.newBlob(htmlContent, 'text/html', 'Nota_' + notaId + '.html');
    var fileHtml = DriveApp.createFile(blob);
    var pdfBlob = fileHtml.getAs('application/pdf');
    var filePdf = DriveApp.createFile(pdfBlob).setName('Nota_' + notaId + '.pdf');
    
    DriveApp.getFileById(fileHtml.getId()).setTrashed(true);
    filePdf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return filePdf.getUrl();
  } catch (e) {
    return null;
  }
}


