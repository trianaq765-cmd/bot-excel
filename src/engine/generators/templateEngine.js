// ═══════════════════════════════════════════════════════════════════════════
// TEMPLATE ENGINE - Generate Excel templates
// ═══════════════════════════════════════════════════════════════════════════

const ExcelJS = require('exceljs');
const { FORMATTING, INDONESIA } = require('../../utils/constants');
const helpers = require('../../utils/helpers');

class TemplateEngine {
  constructor() {
    this.templates = {
      'invoice': this.createInvoiceTemplate.bind(this),
      'payroll': this.createPayrollTemplate.bind(this),
      'inventory': this.createInventoryTemplate.bind(this),
      'sales-report': this.createSalesReportTemplate.bind(this),
      'budget': this.createBudgetTemplate.bind(this),
      'attendance': this.createAttendanceTemplate.bind(this),
      'expense': this.createExpenseTemplate.bind(this),
    };
  }

  // ─────────────────────────────────────────────────────────────────────────
  // MAIN TEMPLATE GENERATION
  // ─────────────────────────────────────────────────────────────────────────

  /**
   * Generate template by name
   */
  async generate(templateName, options = {}) {
    const templateFn = this.templates[templateName.toLowerCase()];
    
    if (!templateFn) {
      throw new Error(`Template "${templateName}" not found. Available: ${Object.keys(this.templates).join(', ')}`);
    }

    return await templateFn(options);
  }

  /**
   * List available templates
   */
  listTemplates() {
    return Object.keys(this.templates).map(name => ({
      name,
      description: this._getTemplateDescription(name),
    }));
  }

  // ─────────────────────────────────────────────────────────────────────────
  // INVOICE TEMPLATE
  // ─────────────────────────────────────────────────────────────────────────

  async createInvoiceTemplate(options = {}) {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Excel Intelligent Bot';
    workbook.created = new Date();

    const ws = workbook.addWorksheet('Invoice', {
      properties: { tabColor: { argb: '2B579A' } },
      pageSetup: { paperSize: 9, orientation: 'portrait' },
    });

    // Set column widths
    ws.columns = [
      { width: 5 },   // A - No
      { width: 35 },  // B - Description
      { width: 10 },  // C - Qty
      { width: 18 },  // D - Unit Price
      { width: 18 },  // E - Total
    ];

    let row = 1;

    // ─────────────────────────────────────────────────────────────────────
    // Header Section
    // ─────────────────────────────────────────────────────────────────────
    
    // Company name
    ws.mergeCells(`A${row}:E${row}`);
    const companyCell = ws.getCell(`A${row}`);
    companyCell.value = options.companyName || '[NAMA PERUSAHAAN]';
    companyCell.font = { size: 18, bold: true, color: { argb: 'FF2B579A' } };
    row++;

    // Company address
    ws.mergeCells(`A${row}:E${row}`);
    ws.getCell(`A${row}`).value = options.companyAddress || '[Alamat Perusahaan]';
    row++;

    // Company contact
    ws.mergeCells(`A${row}:E${row}`);
    ws.getCell(`A${row}`).value = options.companyPhone || 'Telp: [Nomor Telepon]';
    row += 2;

    // Invoice title
    ws.mergeCells(`A${row}:E${row}`);
    const titleCell = ws.getCell(`A${row}`);
    titleCell.value = 'INVOICE';
    titleCell.font = { size: 16, bold: true };
    titleCell.alignment = { horizontal: 'center' };
    titleCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF2B579A' },
    };
    titleCell.font = { size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
    row += 2;

    // Invoice details
    ws.getCell(`A${row}`).value = 'Invoice No:';
    ws.getCell(`B${row}`).value = options.invoiceNo || `INV-${Date.now()}`;
    ws.getCell(`D${row}`).value = 'Tanggal:';
    ws.getCell(`E${row}`).value = helpers.formatDate(new Date());
    row++;

    ws.getCell(`A${row}`).value = 'Due Date:';
    ws.getCell(`B${row}`).value = options.dueDate || helpers.formatDate(new Date(Date.now() + 30 * 24 * 60 * 60 * 1000));
    row += 2;

    // Bill To section
    ws.getCell(`A${row}`).value = 'BILL TO:';
    ws.getCell(`A${row}`).font = { bold: true };
    row++;
    ws.getCell(`A${row}`).value = options.clientName || '[Nama Client]';
    row++;
    ws.getCell(`A${row}`).value = options.clientAddress || '[Alamat Client]';
    row++;
    ws.getCell(`A${row}`).value = `NPWP: ${options.clientNPWP || '[NPWP Client]'}`;
    row += 2;

    // ─────────────────────────────────────────────────────────────────────
    // Items Table
    // ─────────────────────────────────────────────────────────────────────
    
    const tableStartRow = row;
    
    // Table header
    const headers = ['No', 'Deskripsi', 'Qty', 'Harga Satuan', 'Total'];
    const headerRow = ws.addRow(headers);
    headerRow.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF2B579A' },
      };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
    });
    headerRow.height = 25;
    row++;

    // Sample items (empty rows for filling)
    const itemCount = options.itemCount || 5;
    for (let i = 1; i <= itemCount; i++) {
      const itemRow = ws.addRow([
        i,
        '',
        '',
        '',
        { formula: `C${row}*D${row}` },
      ]);
      
      itemRow.eachCell((cell, colNumber) => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
        
        if (colNumber === 4 || colNumber === 5) {
          cell.numFmt = '"Rp "#,##0';
        }
      });
      
      row++;
    }

    row++;

    // ─────────────────────────────────────────────────────────────────────
    // Totals Section
    // ─────────────────────────────────────────────────────────────────────
    
    const itemEndRow = row - 2;
    
    // Subtotal
    ws.getCell(`D${row}`).value = 'Subtotal:';
    ws.getCell(`D${row}`).font = { bold: true };
    ws.getCell(`D${row}`).alignment = { horizontal: 'right' };
    ws.getCell(`E${row}`).value = { formula: `SUM(E${tableStartRow + 1}:E${itemEndRow})` };
    ws.getCell(`E${row}`).numFmt = '"Rp "#,##0';
    ws.getCell(`E${row}`).font = { bold: true };
    row++;

    // PPN
    ws.getCell(`D${row}`).value = 'PPN (11%):';
    ws.getCell(`D${row}`).alignment = { horizontal: 'right' };
    ws.getCell(`E${row}`).value = { formula: `E${row - 1}*0.11` };
    ws.getCell(`E${row}`).numFmt = '"Rp "#,##0';
    row++;

    // Grand Total
    ws.getCell(`D${row}`).value = 'GRAND TOTAL:';
    ws.getCell(`D${row}`).font = { bold: true, size: 12 };
    ws.getCell(`D${row}`).alignment = { horizontal: 'right' };
    ws.getCell(`E${row}`).value = { formula: `E${row - 2}+E${row - 1}` };
    ws.getCell(`E${row}`).numFmt = '"Rp "#,##0';
    ws.getCell(`E${row}`).font = { bold: true, size: 12 };
    ws.getCell(`E${row}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFE2EFDA' },
    };
    row += 2;

    // ─────────────────────────────────────────────────────────────────────
    // Payment Info
    // ─────────────────────────────────────────────────────────────────────
    
    ws.mergeCells(`A${row}:E${row}`);
    ws.getCell(`A${row}`).value = 'Informasi Pembayaran:';
    ws.getCell(`A${row}`).font = { bold: true };
    row++;

    ws.mergeCells(`A${row}:E${row}`);
    ws.getCell(`A${row}`).value = `Bank: ${options.bankName || '[Nama Bank]'} | No. Rek: ${options.bankAccount || '[No. Rekening]'} | a.n. ${options.accountName || '[Nama Pemilik]'}`;
    row += 2;

    // Footer
    ws.mergeCells(`A${row}:E${row}`);
    ws.getCell(`A${row}`).value = 'Terima kasih atas kepercayaan Anda.';
    ws.getCell(`A${row}`).alignment = { horizontal: 'center' };
    ws.getCell(`A${row}`).font = { italic: true };

    return workbook;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // PAYROLL TEMPLATE
  // ─────────────────────────────────────────────────────────────────────────

  async createPayrollTemplate(options = {}) {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Excel Intelligent Bot';

    const ws = workbook.addWorksheet('Payroll', {
      properties: { tabColor: { argb: '00B050' } },
    });

    // Column widths
    ws.columns = [
      { width: 5 },   // No
      { width: 20 },  // Nama
      { width: 18 },  // NIK
      { width: 15 },  // Jabatan
      { width: 15 },  // Gaji Pokok
      { width: 12 },  // Tunjangan
      { width: 12 },  // Lembur
      { width: 12 },  // BPJS Kes
      { width: 12 },  // BPJS TK
      { width: 12 },  // PPh 21
      { width: 15 },  // Gaji Bersih
    ];

    let row = 1;

    // Title
    ws.mergeCells(`A${row}:K${row}`);
    ws.getCell(`A${row}`).value = options.companyName || 'SLIP GAJI KARYAWAN';
    ws.getCell(`A${row}`).font = { size: 16, bold: true };
    ws.getCell(`A${row}`).alignment = { horizontal: 'center' };
    row++;

    // Period
    ws.mergeCells(`A${row}:K${row}`);
    ws.getCell(`A${row}`).value = `Periode: ${options.period || new Date().toLocaleString('id-ID', { month: 'long', year: 'numeric' })}`;
    ws.getCell(`A${row}`).alignment = { horizontal: 'center' };
    row += 2;

    // Headers
    const headers = [
      'No', 'Nama Karyawan', 'NIK', 'Jabatan', 
      'Gaji Pokok', 'Tunjangan', 'Lembur',
      'BPJS Kes', 'BPJS TK', 'PPh 21', 'Gaji Bersih'
    ];

    const headerRow = ws.addRow(headers);
    headerRow.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF00B050' },
      };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
    });
    headerRow.height = 30;
    row++;

    // Sample data rows
    const employeeCount = options.employeeCount || 10;
    for (let i = 1; i <= employeeCount; i++) {
      const currentRow = row;
      const dataRow = ws.addRow([
        i,
        '',  // Nama
        '',  // NIK
        '',  // Jabatan
        '',  // Gaji Pokok
        '',  // Tunjangan
        '',  // Lembur
        { formula: `(E${currentRow}+F${currentRow})*0.01` },  // BPJS Kes 1%
        { formula: `(E${currentRow}+F${currentRow})*0.02` },  // BPJS TK 2%
        '',  // PPh 21
        { formula: `E${currentRow}+F${currentRow}+G${currentRow}-H${currentRow}-I${currentRow}-J${currentRow}` },
      ]);

      dataRow.eachCell((cell, colNumber) => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
        
        if (colNumber >= 5) {
          cell.numFmt = '"Rp "#,##0';
        }
      });
      
      row++;
    }

    row++;

    // Totals row
    const totalRow = ws.addRow([
      '', 'TOTAL', '', '',
      { formula: `SUM(E5:E${row - 2})` },
      { formula: `SUM(F5:F${row - 2})` },
      { formula: `SUM(G5:G${row - 2})` },
      { formula: `SUM(H5:H${row - 2})` },
      { formula: `SUM(I5:I${row - 2})` },
      { formula: `SUM(J5:J${row - 2})` },
      { formula: `SUM(K5:K${row - 2})` },
    ]);

    totalRow.eachCell((cell, colNumber) => {
      cell.font = { bold: true };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE2EFDA' },
      };
      if (colNumber >= 5) {
        cell.numFmt = '"Rp "#,##0';
      }
    });

    return workbook;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // INVENTORY TEMPLATE
  // ─────────────────────────────────────────────────────────────────────────

  async createInventoryTemplate(options = {}) {
    const workbook = new ExcelJS.Workbook();
    
    const ws = workbook.addWorksheet('Inventory', {
      properties: { tabColor: { argb: 'FFC000' } },
      views: [{ state: 'frozen', ySplit: 1 }],
    });

    ws.columns = [
      { width: 12 },  // Kode
      { width: 30 },  // Nama Barang
      { width: 15 },  // Kategori
      { width: 10 },  // Satuan
      { width: 10 },  // Stok
      { width: 12 },  // Min Stok
      { width: 15 },  // Harga Beli
      { width: 15 },  // Harga Jual
      { width: 18 },  // Nilai Stok
      { width: 15 },  // Lokasi
      { width: 12 },  // Status
    ];

    const headers = [
      'Kode', 'Nama Barang', 'Kategori', 'Satuan', 'Stok',
      'Min Stok', 'Harga Beli', 'Harga Jual', 'Nilai Stok', 'Lokasi', 'Status'
    ];

    const headerRow = ws.addRow(headers);
    headerRow.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFC000' },
      };
      cell.font = { bold: true };
      cell.alignment = { horizontal: 'center' };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
    });

    // Sample rows with formulas
    for (let i = 2; i <= 21; i++) {
      const row = ws.addRow([
        `ITM${String(i - 1).padStart(3, '0')}`,
        '', '', 'Pcs', '', '',
        '', '',
        { formula: `E${i}*G${i}` },  // Nilai Stok = Stok × Harga Beli
        '',
        { formula: `IF(E${i}<F${i},"Reorder","OK")` },  // Status
      ]);

      row.eachCell((cell, colNumber) => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
        if ([7, 8, 9].includes(colNumber)) {
          cell.numFmt = '"Rp "#,##0';
        }
      });
    }

    return workbook;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // SALES REPORT TEMPLATE
  // ─────────────────────────────────────────────────────────────────────────

  async createSalesReportTemplate(options = {}) {
    const workbook = new ExcelJS.Workbook();
    
    const ws = workbook.addWorksheet('Sales Report', {
      properties: { tabColor: { argb: '7030A0' } },
    });

    ws.columns = [
      { width: 12 },  // Tanggal
      { width: 15 },  // No Invoice
      { width: 25 },  // Customer
      { width: 25 },  // Produk
      { width: 10 },  // Qty
      { width: 15 },  // Harga
      { width: 18 },  // Total
      { width: 15 },  // Sales Rep
      { width: 12 },  // Status
    ];

    let row = 1;

    // Title
    ws.mergeCells(`A${row}:I${row}`);
    ws.getCell(`A${row}`).value = 'LAPORAN PENJUALAN';
    ws.getCell(`A${row}`).font = { size: 16, bold: true };
    ws.getCell(`A${row}`).alignment = { horizontal: 'center' };
    row++;

    ws.mergeCells(`A${row}:I${row}`);
    ws.getCell(`A${row}`).value = `Periode: ${options.period || new Date().toLocaleString('id-ID', { month: 'long', year: 'numeric' })}`;
    ws.getCell(`A${row}`).alignment = { horizontal: 'center' };
    row += 2;

    // Headers
    const headers = ['Tanggal', 'No Invoice', 'Customer', 'Produk', 'Qty', 'Harga', 'Total', 'Sales Rep', 'Status'];
    const headerRow = ws.addRow(headers);
    headerRow.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF7030A0' },
      };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.alignment = { horizontal: 'center' };
    });
    row++;

    // Sample rows
    for (let i = 0; i < 20; i++) {
      const currentRow = row + i;
      const dataRow = ws.addRow([
        '', '', '', '', '', '',
        { formula: `E${currentRow}*F${currentRow}` },
        '', ''
      ]);
      
      dataRow.eachCell((cell, colNumber) => {
        if ([6, 7].includes(colNumber)) {
          cell.numFmt = '"Rp "#,##0';
        }
      });
    }

    return workbook;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // BUDGET TEMPLATE
  // ─────────────────────────────────────────────────────────────────────────

  async createBudgetTemplate(options = {}) {
    const workbook = new ExcelJS.Workbook();
    
    const ws = workbook.addWorksheet('Budget', {
      properties: { tabColor: { argb: '0070C0' } },
    });

    // Similar structure as above with budget-specific columns
    ws.columns = [
      { width: 5 },
      { width: 25 },
      { width: 18 },
      { width: 18 },
      { width: 18 },
      { width: 18 },
      { width: 12 },
    ];

    const headers = ['No', 'Item/Kategori', 'Budget', 'Actual', 'Variance', '% Used', 'Status'];
    const headerRow = ws.addRow(headers);
    headerRow.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0070C0' } };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.alignment = { horizontal: 'center' };
    });

    for (let i = 2; i <= 15; i++) {
      ws.addRow([
        i - 1, '', '', '',
        { formula: `C${i}-D${i}` },
        { formula: `IF(C${i}>0,D${i}/C${i},0)` },
        { formula: `IF(D${i}>C${i},"Over","OK")` },
      ]);
      
      ws.getCell(`C${i}`).numFmt = '"Rp "#,##0';
      ws.getCell(`D${i}`).numFmt = '"Rp "#,##0';
      ws.getCell(`E${i}`).numFmt = '"Rp "#,##0';
      ws.getCell(`F${i}`).numFmt = '0.0%';
    }

    return workbook;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // ATTENDANCE TEMPLATE
  // ─────────────────────────────────────────────────────────────────────────

  async createAttendanceTemplate(options = {}) {
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet('Attendance');

    // Create monthly attendance grid
    const month = options.month || new Date().getMonth();
    const year = options.year || new Date().getFullYear();
    const daysInMonth = new Date(year, month + 1, 0).getDate();

    // Build headers: Name + 1-31 days
    const headers = ['No', 'Nama'];
    for (let d = 1; d <= daysInMonth; d++) {
      headers.push(d.toString());
    }
    headers.push('Hadir', 'Izin', 'Sakit', 'Alpa');

    ws.addRow(headers);
    
    // Add employee rows
    for (let i = 2; i <= 21; i++) {
      const row = ['', ''];
      for (let d = 1; d <= daysInMonth; d++) row.push('');
      row.push(
        { formula: `COUNTIF(C${i}:${String.fromCharCode(66 + daysInMonth)}${i},"H")` },
        { formula: `COUNTIF(C${i}:${String.fromCharCode(66 + daysInMonth)}${i},"I")` },
        { formula: `COUNTIF(C${i}:${String.fromCharCode(66 + daysInMonth)}${i},"S")` },
        { formula: `COUNTIF(C${i}:${String.fromCharCode(66 + daysInMonth)}${i},"A")` },
      );
      ws.addRow(row);
    }

    return workbook;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // EXPENSE TEMPLATE
  // ─────────────────────────────────────────────────────────────────────────

  async createExpenseTemplate(options = {}) {
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet('Expenses');

    ws.columns = [
      { width: 12 },
      { width: 30 },
      { width: 15 },
      { width: 18 },
      { width: 20 },
      { width: 15 },
    ];

    const headers = ['Tanggal', 'Deskripsi', 'Kategori', 'Jumlah', 'Metode Bayar', 'Bukti'];
    ws.addRow(headers);

    // Style and add sample rows...
    return workbook;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // HELPER METHODS
  // ─────────────────────────────────────────────────────────────────────────

  _getTemplateDescription(name) {
    const descriptions = {
      'invoice': 'Invoice/Faktur dengan PPN 11%',
      'payroll': 'Slip gaji dengan BPJS & PPh 21',
      'inventory': 'Tracking stok dengan reorder alert',
      'sales-report': 'Laporan penjualan dengan summary',
      'budget': 'Budget vs Actual dengan variance',
      'attendance': 'Absensi bulanan karyawan',
      'expense': 'Tracking pengeluaran/reimbursement',
    };
    return descriptions[name] || name;
  }
}

module.exports = new TemplateEngine();
