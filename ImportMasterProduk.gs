/**
 * Impor Massal 13 Produk + 58 Atribut + 71 Variasi Eksisting ke Google Sheets (V4.2)
 * Mengisi MASTER_PRODUK, ATRIBUT_PRODUK, VARIASI_PRODUK, STOK_VARIASI, dan OPERATION_LOG secara sinkron
 * dengan mempertahankan ID_Variasi spesifik (A-R01-T01 .. VAR-ECF3C81AE29B) serta status aktif/nonaktif-nya.
 *
 * Cara pakai:
 * 1. Tempelkan file ini di editor Google Apps Script (bersama Code.gs).
 * 2. Simpan (Ctrl+S), pilih fungsi "importMasterProdukBatch" di toolbar atas, lalu klik "Run".
 */
function importMasterProdukBatch() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    throw new Error('Server sedang sibuk memproses transaksi lain. Silakan coba lagi.');
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (typeof setup === 'function') {
      setup();
    }

    var catalog = [
      {
        barcode: 'A-R01',
        nama: 'IT WILL GET BETTER',
        aktif: true,
        createdAt: '2026-11-06 19:44:12',
        updatedAt: '2026-11-06 19:44:12',
        attributes: [
          {
            name: 'Warna',
            type: 'color',
            values: [
              { value: 'COKLAT', hex: '#8B4513' },
              { value: 'HITAM', hex: '#000000' },
              { value: 'PUTIH', hex: '#FFFFFF' }
            ]
          },
          {
            name: 'Ukuran',
            type: 'text',
            values: [{ value: 'ALLSIZE' }, { value: 'XL' }]
          }
        ]
      },
      {
        barcode: 'A-R02',
        nama: 'ANGEL',
        aktif: true,
        createdAt: '2026-11-06 19:44:12',
        updatedAt: '2026-11-06 19:44:12',
        attributes: [
          {
            name: 'Warna',
            type: 'color',
            values: [
              { value: 'COKLAT', hex: '#8B4513' },
              { value: 'HITAM', hex: '#000000' },
              { value: 'PINK MUDA', hex: '#FFB6C1' },
              { value: 'PUTIH', hex: '#FFFFFF' }
            ]
          },
          {
            name: 'Ukuran',
            type: 'text',
            values: [{ value: 'ALLSIZE' }, { value: 'XL' }]
          }
        ]
      },
      {
        barcode: 'A-R03',
        nama: 'BRAZIL',
        aktif: true,
        createdAt: '2026-11-06 19:44:12',
        updatedAt: '2026-11-06 19:44:12',
        attributes: [
          {
            name: 'Warna',
            type: 'color',
            values: [
              { value: 'HIJAU', hex: '#22C55E' },
              { value: 'HITAM', hex: '#000000' },
              { value: 'KUNING', hex: '#EAB308' },
              { value: 'PUTIH', hex: '#FFFFFF' }
            ]
          },
          {
            name: 'Ukuran',
            type: 'text',
            values: [{ value: 'ALLSIZE' }, { value: 'XL' }]
          }
        ]
      },
      {
        barcode: 'A-R04',
        nama: 'AVEZED SEVENFOLD A7X',
        aktif: true,
        createdAt: '2026-11-06 19:44:12',
        updatedAt: '2026-11-06 19:44:12',
        attributes: [
          {
            name: 'Warna',
            type: 'color',
            values: [
              { value: 'COKLAT', hex: '#8B4513' },
              { value: 'HITAM', hex: '#000000' },
              { value: 'PINK', hex: '#FF69B4' },
              { value: 'PINK MUDA', hex: '#FFB6C1' },
              { value: 'PUTIH', hex: '#FFFFFF' }
            ]
          },
          {
            name: 'Ukuran',
            type: 'text',
            values: [{ value: 'ALLSIZE' }, { value: 'XL' }]
          }
        ]
      },
      {
        barcode: 'A-R05',
        nama: 'POLKADOT LOVE',
        aktif: true,
        createdAt: '2026-11-06 19:44:12',
        updatedAt: '2026-11-06 19:44:12',
        attributes: [
          {
            name: 'Warna',
            type: 'color',
            values: [
              { value: 'COKLAT', hex: '#8B4513' },
              { value: 'HITAM', hex: '#000000' },
              { value: 'PUTIH', hex: '#FFFFFF' }
            ]
          },
          {
            name: 'Ukuran',
            type: 'text',
            values: [{ value: 'ALLSIZE' }, { value: 'XL' }]
          }
        ]
      },
      {
        barcode: 'A-R06',
        nama: 'POLKADOT LOVE PREMIUM',
        aktif: true,
        createdAt: '2026-11-06 19:44:12',
        updatedAt: '2026-11-06 19:44:12',
        attributes: [
          {
            name: 'Warna',
            type: 'color',
            values: [
              { value: 'COKLAT', hex: '#8B4513' },
              { value: 'HITAM', hex: '#000000' },
              { value: 'PUTIH', hex: '#FFFFFF' }
            ]
          },
          {
            name: 'Ukuran',
            type: 'text',
            values: [{ value: 'ALLSIZE' }, { value: 'XL' }]
          }
        ]
      },
      {
        barcode: 'A-R09',
        nama: 'SPIDERMAN NEW',
        aktif: true,
        createdAt: '2026-11-06 19:44:12',
        updatedAt: '2026-11-06 19:44:12',
        attributes: [
          {
            name: 'Warna',
            type: 'color',
            values: [
              { value: 'COKLAT', hex: '#8B4513' },
              { value: 'HITAM', hex: '#000000' },
              { value: 'MERAH', hex: '#EF4444' }
            ]
          },
          {
            name: 'Ukuran',
            type: 'text',
            values: [{ value: 'ALLSIZE' }, { value: 'XL' }]
          }
        ]
      },
      {
        barcode: 'A-R10',
        nama: 'KUDA',
        aktif: true,
        createdAt: '2026-11-06 19:44:12',
        updatedAt: '2026-11-06 19:44:12',
        attributes: [
          {
            name: 'Warna',
            type: 'color',
            values: [
              { value: 'COKLAT', hex: '#8B4513' },
              { value: 'HITAM', hex: '#000000' },
              { value: 'PINK', hex: '#FF69B4' },
              { value: 'PUTIH', hex: '#FFFFFF' }
            ]
          },
          {
            name: 'Ukuran',
            type: 'text',
            values: [{ value: 'ALLSIZE' }, { value: 'XL' }]
          }
        ]
      },
      {
        barcode: 'A-R11',
        nama: 'ONE SHOULDER',
        aktif: true,
        createdAt: '2026-11-06 19:44:12',
        updatedAt: '2026-11-06 19:44:12',
        attributes: [
          {
            name: 'Ukuran',
            type: 'text',
            values: [{ value: 'ALLSIZE' }, { value: 'XL' }]
          }
        ]
      },
      {
        barcode: 'A-R12',
        nama: 'SABRINA',
        aktif: true,
        createdAt: '2026-11-06 19:44:12',
        updatedAt: '2026-11-06 19:44:12',
        attributes: [
          {
            name: 'Ukuran',
            type: 'text',
            values: [{ value: 'ALLSIZE' }, { value: 'XL' }]
          }
        ]
      },
      {
        barcode: 'A-R13',
        nama: 'SARDEN RINGGER',
        aktif: true,
        createdAt: '2026-11-06 19:44:12',
        updatedAt: '2026-11-06 19:44:12',
        attributes: [
          {
            name: 'Ukuran',
            type: 'text',
            values: [{ value: 'ALLSIZE' }, { value: 'XL' }]
          }
        ]
      },
      {
        barcode: 'A-R14',
        nama: 'MC QUEEN',
        aktif: true,
        createdAt: '2026-11-06 19:44:12',
        updatedAt: '2026-11-06 19:44:12',
        attributes: [
          {
            name: 'Warna',
            type: 'color',
            values: [
              { value: 'HITAM', hex: '#000000' },
              { value: 'MERAH', hex: '#EF4444' },
              { value: 'PUTIH', hex: '#FFFFFF' }
            ]
          },
          {
            name: 'Ukuran',
            type: 'text',
            values: [{ value: 'ALLSIZE' }, { value: 'XL' }]
          }
        ]
      },
      {
        barcode: 'A-R15',
        nama: 'OVERSIZE SARDEN',
        aktif: true,
        createdAt: '2026-11-06 19:44:12',
        updatedAt: '2026-11-06 19:44:12',
        attributes: [
          {
            name: 'Ukuran',
            type: 'text',
            values: [{ value: 'ALLSIZE' }, { value: 'XL' }]
          }
        ]
      }
    ];

    var exactVariants = [
      ['A-R01-T01', 'A-R01', 'Warna: PUTIH · Ukuran: ALLSIZE', '{"Warna": "PUTIH", "Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R01-T02', 'A-R01', 'Warna: HITAM · Ukuran: ALLSIZE', '{"Warna": "HITAM", "Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R01-T03', 'A-R01', 'Warna: COKLAT · Ukuran: ALLSIZE', '{"Warna": "COKLAT", "Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R01-T04', 'A-R01', 'Warna: PUTIH · Ukuran: XL', '{"Warna": "PUTIH", "Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R01-T05', 'A-R01', 'Warna: HITAM · Ukuran: XL', '{"Warna": "HITAM", "Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R01-T06', 'A-R01', 'Warna: COKLAT · Ukuran: XL', '{"Warna": "COKLAT", "Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R02-T01', 'A-R02', 'Warna: PUTIH · Ukuran: ALLSIZE', '{"Warna": "PUTIH", "Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R02-T02', 'A-R02', 'Warna: HITAM · Ukuran: ALLSIZE', '{"Warna": "HITAM", "Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R02-T03', 'A-R02', 'Warna: PINK MUDA · Ukuran: ALLSIZE', '{"Warna": "PINK MUDA", "Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R02-T04', 'A-R02', 'Warna: COKLAT · Ukuran: ALLSIZE', '{"Warna": "COKLAT", "Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R02-T05', 'A-R02', 'Warna: PUTIH · Ukuran: XL', '{"Warna": "PUTIH", "Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R02-T06', 'A-R02', 'Warna: HITAM · Ukuran: XL', '{"Warna": "HITAM", "Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R02-T07', 'A-R02', 'Warna: PINK MUDA · Ukuran: XL', '{"Warna": "PINK MUDA", "Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R02-T08', 'A-R02', 'Warna: COKLAT · Ukuran: XL', '{"Warna": "COKLAT", "Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R03-T01', 'A-R03', 'Warna: KUNING · Ukuran: ALLSIZE', '{"Warna": "KUNING", "Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R03-T02', 'A-R03', 'Warna: HITAM · Ukuran: ALLSIZE', '{"Warna": "HITAM", "Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R03-T03', 'A-R03', 'Warna: PUTIH · Ukuran: ALLSIZE', '{"Warna": "PUTIH", "Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R03-T04', 'A-R03', 'Warna: HIJAU · Ukuran: ALLSIZE', '{"Warna": "HIJAU", "Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R03-T05', 'A-R03', 'Warna: KUNING · Ukuran: XL', '{"Warna": "KUNING", "Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R03-T06', 'A-R03', 'Warna: HITAM · Ukuran: XL', '{"Warna": "HITAM", "Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R03-T07', 'A-R03', 'Warna: PUTIH · Ukuran: XL', '{"Warna": "PUTIH", "Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R03-T08', 'A-R03', 'Warna: HIJAU · Ukuran: XL', '{"Warna": "HIJAU", "Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R04-T01', 'A-R04', 'Warna: HITAM · Ukuran: ALLSIZE', '{"Warna": "HITAM", "Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R04-T02', 'A-R04', 'Warna: PUTIH · Ukuran: ALLSIZE', '{"Warna": "PUTIH", "Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R04-T03', 'A-R04', 'Warna: PINK · Ukuran: ALLSIZE', '{"Warna": "PINK", "Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R04-T04', 'A-R04', 'Warna: COKLAT · Ukuran: ALLSIZE', '{"Warna": "COKLAT", "Ukuran": "ALLSIZE"}', false, '2026-11-06 19:44:12', '2026-11-08 0:07:23'],
      ['A-R04-T05', 'A-R04', 'Warna: HITAM · Ukuran: XL', '{"Warna": "HITAM", "Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R04-T06', 'A-R04', 'Warna: PUTIH · Ukuran: XL', '{"Warna": "PUTIH", "Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R04-T07', 'A-R04', 'Warna: PINK MUDA · Ukuran: XL', '{"Warna": "PINK MUDA", "Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R04-T08', 'A-R04', 'Warna: COKLAT · Ukuran: XL', '{"Warna": "COKLAT", "Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R05-T01', 'A-R05', 'Warna: PUTIH · Ukuran: ALLSIZE', '{"Warna": "PUTIH", "Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R05-T02', 'A-R05', 'Warna: HITAM · Ukuran: ALLSIZE', '{"Warna": "HITAM", "Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R05-T03', 'A-R05', 'Warna: COKLAT · Ukuran: ALLSIZE', '{"Warna": "COKLAT", "Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R05-T04', 'A-R05', 'Warna: PUTIH · Ukuran: XL', '{"Warna": "PUTIH", "Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R05-T05', 'A-R05', 'Warna: HITAM · Ukuran: XL', '{"Warna": "HITAM", "Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R05-T06', 'A-R05', 'Warna: COKLAT · Ukuran: XL', '{"Warna": "COKLAT", "Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R06-T01', 'A-R06', 'Warna: PUTIH · Ukuran: ALLSIZE', '{"Warna": "PUTIH", "Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R07-T01', 'A-R06', 'Warna: HITAM · Ukuran: ALLSIZE', '{"Warna": "HITAM", "Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R08-T01', 'A-R06', 'Warna: COKLAT · Ukuran: ALLSIZE', '{"Warna": "COKLAT", "Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R06-T02', 'A-R06', 'Warna: PUTIH · Ukuran: XL', '{"Warna": "PUTIH", "Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R07-T02', 'A-R06', 'Warna: HITAM · Ukuran: XL', '{"Warna": "HITAM", "Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R08-T02', 'A-R06', 'Warna: COKLAT · Ukuran: XL', '{"Warna": "COKLAT", "Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R09-T01', 'A-R09', 'Warna: MERAH · Ukuran: ALLSIZE', '{"Warna": "MERAH", "Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R09-T02', 'A-R09', 'Warna: HITAM · Ukuran: ALLSIZE', '{"Warna": "HITAM", "Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R09-T03', 'A-R09', 'Warna: COKLAT · Ukuran: ALLSIZE', '{"Warna": "COKLAT", "Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R09-T04', 'A-R09', 'Warna: MERAH · Ukuran: XL', '{"Warna": "MERAH", "Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R09-T05', 'A-R09', 'Warna: HITAM · Ukuran: XL', '{"Warna": "HITAM", "Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R09-T06', 'A-R09', 'Warna: COKLAT · Ukuran: XL', '{"Warna": "COKLAT", "Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R10-T01', 'A-R10', 'Warna: PUTIH · Ukuran: ALLSIZE', '{"Warna": "PUTIH", "Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R10-T02', 'A-R10', 'Warna: HITAM · Ukuran: ALLSIZE', '{"Warna": "HITAM", "Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R10-T03', 'A-R10', 'Warna: COKLAT · Ukuran: ALLSIZE', '{"Warna": "COKLAT", "Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R10-T04', 'A-R10', 'Warna: PINK · Ukuran: ALLSIZE', '{"Warna": "PINK", "Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R10-T05', 'A-R10', 'Warna: PUTIH · Ukuran: XL', '{"Warna": "PUTIH", "Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R10-T06', 'A-R10', 'Warna: HITAM · Ukuran: XL', '{"Warna": "HITAM", "Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R10-T07', 'A-R10', 'Warna: COKLAT · Ukuran: XL', '{"Warna": "COKLAT", "Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R10-T08', 'A-R10', 'Warna: PINK · Ukuran: XL', '{"Warna": "PINK", "Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R11-T01', 'A-R11', 'Ukuran: ALLSIZE', '{"Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R11-T02', 'A-R11', 'Ukuran: XL', '{"Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R12-T01', 'A-R12', 'Ukuran: XL', '{"Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R12-T02', 'A-R12', 'Ukuran: ALLSIZE', '{"Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R13-T01', 'A-R13', 'Ukuran: ALLSIZE', '{"Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R13-T02', 'A-R13', 'Ukuran: XL', '{"Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R14-T01', 'A-R14', 'Warna: MERAH · Ukuran: ALLSIZE', '{"Warna": "MERAH", "Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R14-T02', 'A-R14', 'Warna: HITAM · Ukuran: ALLSIZE', '{"Warna": "HITAM", "Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R14-T03', 'A-R14', 'Warna: PUTIH · Ukuran: ALLSIZE', '{"Warna": "PUTIH", "Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R14-T04', 'A-R14', 'Warna: MERAH · Ukuran: XL', '{"Warna": "MERAH", "Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R14-T05', 'A-R14', 'Warna: HITAM · Ukuran: XL', '{"Warna": "HITAM", "Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R14-T06', 'A-R14', 'Warna: PUTIH · Ukuran: XL', '{"Warna": "PUTIH", "Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R15-T01', 'A-R15', 'Ukuran: ALLSIZE', '{"Ukuran": "ALLSIZE"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['A-R15-T02', 'A-R15', 'Ukuran: XL', '{"Ukuran": "XL"}', true, '2026-11-06 19:44:12', '2026-11-06 19:44:12'],
      ['VAR-ECF3C81AE29B', 'A-R04', 'Warna: PINK MUDA · Ukuran: ALLSIZE', '{"Warna":"PINK MUDA","Ukuran":"ALLSIZE"}', true, '2026-11-08 0:06:32', '2026-11-08 0:06:32']
    ];

    var shMaster = ss.getSheetByName(IBS.SHEET_MASTER);
    var shAttrs = ss.getSheetByName(IBS.SHEET_ATTRIBUTES);
    var shVariants = ss.getSheetByName(IBS.SHEET_VARIANTS);
    var shStock = ss.getSheetByName(IBS.SHEET_STOCK);
    var shLog = ss.getSheetByName(IBS.SHEET_LOG);
    var activeLocs = getActiveLocationsV2_(ss);
    var now = new Date();

    // Indeks data yang sudah ada agar idempoten
    var existingMasterMap = {};
    readObjectsV2_(shMaster).forEach(function (r) {
      var bc = upperV2_(r.Barcode);
      if (bc) existingMasterMap[bc] = true;
    });

    var existingAttrKeys = {};
    readObjectsV2_(shAttrs).forEach(function (r) {
      var key = upperV2_(r.Barcode_Induk) + '|' + upperV2_(r.Nama_Atribut) + '|' + upperV2_(r.Nilai);
      existingAttrKeys[key] = true;
    });

    var existingVariantIds = {};
    var placeholderVariantIdsToClean = {};
    readObjectsV2_(shVariants).forEach(function (r) {
      var vid = textV2_(r.ID_Variasi);
      if (vid) existingVariantIds[vid] = true;
      if (textV2_(r.Atribut_JSON) === '{}' || textV2_(r.Label_Variasi) === 'Standar') {
        placeholderVariantIdsToClean[vid] = r._row;
      }
    });

    var existingStockKeys = {};
    readObjectsV2_(shStock).forEach(function (r) {
      var sk = textV2_(r.ID_Variasi) + '|' + textV2_(r.Lokasi_Rak);
      existingStockKeys[sk] = true;
    });

    var masterRows = [];
    var attrRows = [];
    var variantRows = [];
    var stockRows = [];
    var logRows = [];

    catalog.forEach(function (item) {
      var bcUpper = upperV2_(item.barcode);
      var cDate = parseIsoLocalDate_(item.createdAt);
      var uDate = parseIsoLocalDate_(item.updatedAt);

      if (!existingMasterMap[bcUpper]) {
        masterRows.push([item.barcode, item.nama, item.aktif, cDate, uDate]);
        existingMasterMap[bcUpper] = true;
        logRows.push([
          now,
          'IMP-' + item.barcode + '-' + Utilities.getUuid().slice(0, 6).toUpperCase(),
          'ProdukBaru',
          item.barcode,
          '',
          '',
          0,
          'SUCCESS',
          '',
          item.nama,
          '',
          '',
          item.nama,
          'Batch Import'
        ]);
      }

      item.attributes.forEach(function (attr, ai) {
        attr.values.forEach(function (val, vi) {
          var aKey = bcUpper + '|' + upperV2_(attr.name) + '|' + upperV2_(val.value);
          if (!existingAttrKeys[aKey]) {
            attrRows.push([
              item.barcode,
              attr.name,
              attr.type,
              val.value,
              val.hex || '',
              ai + 1,
              vi + 1,
              true
            ]);
            existingAttrKeys[aKey] = true;
          }
        });
      });
    });

    exactVariants.forEach(function (v) {
      var varId = v[0];
      var bc = v[1];
      var label = v[2];
      var attrJson = v[3];
      var isActive = v[4];
      var cDate = parseIsoLocalDate_(v[5]);
      var uDate = parseIsoLocalDate_(v[6]);

      if (!existingVariantIds[varId]) {
        variantRows.push([varId, bc, label, attrJson, isActive, cDate, uDate]);
        existingVariantIds[varId] = true;
      }

      activeLocs.forEach(function (loc) {
        var sk = varId + '|' + loc.value;
        if (!existingStockKeys[sk]) {
          stockRows.push([varId, bc, loc.value, 0, now]);
          existingStockKeys[sk] = true;
        }
      });
    });

    // Bersihkan variasi placeholder "Standar" (hanya jika saldo stoknya 0)
    var placeholderIds = Object.keys(placeholderVariantIdsToClean);
    if (placeholderIds.length > 0) {
      var stockObjs = readObjectsV2_(shStock);
      var nonZeroIds = {};
      stockObjs.forEach(function (s) {
        var vid = textV2_(s.ID_Variasi);
        if (placeholderVariantIdsToClean[vid] && (Number(s.Kuantitas) || 0) !== 0) {
          nonZeroIds[vid] = true;
        }
      });

      var stockRowsToDelete = [];
      stockObjs.forEach(function (s) {
        var vid = textV2_(s.ID_Variasi);
        if (placeholderVariantIdsToClean[vid] && !nonZeroIds[vid]) {
          stockRowsToDelete.push(s._row);
        }
      });
      stockRowsToDelete.sort(function (a, b) { return b - a; }).forEach(function (r) {
        shStock.deleteRow(r);
      });

      var varRowsToDelete = [];
      placeholderIds.forEach(function (vid) {
        var rowNum = placeholderVariantIdsToClean[vid];
        if (nonZeroIds[vid]) {
          shVariants.getRange(rowNum, 5).setValue(false);
          shVariants.getRange(rowNum, 7).setValue(now);
        } else {
          varRowsToDelete.push(rowNum);
        }
      });
      varRowsToDelete.sort(function (a, b) { return b - a; }).forEach(function (r) {
        shVariants.deleteRow(r);
      });
    }

    // Tulis seluruh data baru secara batch
    if (masterRows.length > 0) {
      var mStart = shMaster.getLastRow() + 1;
      shMaster.getRange(mStart, 1, masterRows.length, 5).setValues(masterRows);
      applyCheckboxValidationV2_(shMaster, 3, mStart, masterRows.length);
    }

    if (attrRows.length > 0) {
      var aStart = shAttrs.getLastRow() + 1;
      shAttrs.getRange(aStart, 1, attrRows.length, 8).setValues(attrRows);
      applyCheckboxValidationV2_(shAttrs, 8, aStart, attrRows.length);
    }

    if (variantRows.length > 0) {
      var vStart = shVariants.getLastRow() + 1;
      shVariants.getRange(vStart, 1, variantRows.length, 7).setValues(variantRows);
      applyCheckboxValidationV2_(shVariants, 5, vStart, variantRows.length);
    }

    if (stockRows.length > 0) {
      var sStart = shStock.getLastRow() + 1;
      shStock.getRange(sStart, 1, stockRows.length, 5).setValues(stockRows);
    }

    if (logRows.length > 0) {
      var lStart = shLog.getLastRow() + 1;
      shLog.getRange(lStart, 1, logRows.length, 14).setValues(logRows);
    }

    SpreadsheetApp.flush();
    var summary = 'Impor selesai: ' + masterRows.length + ' produk baru, ' +
      attrRows.length + ' atribut baru, ' + variantRows.length + ' variasi spesifik ditambahkan (' +
      stockRows.length + ' baris saldo awal 0).';
    Logger.log(summary);
    return summary;
  } finally {
    lock.releaseLock();
  }
}

function parseIsoLocalDate_(str) {
  var parts = String(str).trim().split(' ');
  var d = parts[0].split('-');
  var t = (parts[1] || '00:00:00').split(':');
  return new Date(Number(d[0]), Number(d[1]) - 1, Number(d[2]), Number(t[0] || 0), Number(t[1] || 0), Number(t[2] || 0));
}