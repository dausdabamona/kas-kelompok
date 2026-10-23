#!/usr/bin/env node
// ══════════════════════════════════════════════════════
// Server MCP Kas Kelompok — BACA SAJA.
//
// Meneruskan permintaan ke endpoint JSON aplikasi (doPost di Auth.gs), bukan
// membaca spreadsheet mentah. Bedanya penting: angka yang keluar sudah melewati
// logika aplikasi — baris Dibatalkan & Draft dikecualikan, hak kelompok/desa/
// daerah sudah dipisah, Buku IR yang belum dirinci tidak diakui milik siapa pun.
//
// Tidak ada satu pun alat di sini yang menulis. Sesuai aturan proyek:
// "Sistem melaporkan; manusia memutuskan."
//
// Konfigurasi lewat environment:
//   KAS_EXEC_URL  URL /exec aplikasi
//   KAS_TOKEN     isi Script Property MCP_TOKEN
// ══════════════════════════════════════════════════════
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { CallToolRequestSchema, ListToolsRequestSchema } from '@modelcontextprotocol/sdk/types.js';

const EXEC_URL = process.env.KAS_EXEC_URL || '';
const TOKEN = process.env.KAS_TOKEN || '';

// Tiap alat dipetakan ke satu endpoint baca. Daftar ini harus tetap sama
// dengan MCP_ALLOW_ di Auth.gs — server juga menolak yang tidak terdaftar.
const TOOLS = [
  {
    name: 'kas_ringkasan',
    fn: 'getDashboardData',
    description: 'Ringkasan kas periode berjalan: saldo tunai & bank, total kas, kas di tangan penerobos, Buku IR yang belum dirinci, pengeluaran menunggu persetujuan.',
    args: () => []
  },
  {
    name: 'kas_kepemilikan_saldo',
    fn: 'getKepemilikanSaldo',
    description: 'Pembagian saldo: berapa hak kelompok (boleh dipakai), titipan desa, titipan daerah, Buku IR belum dirinci, dan titipan jamaah (mis. Qurban). Termasuk ringkasan kewajiban pengeluaran rutin bulan ini.',
    args: () => []
  },
  {
    name: 'kas_pengeluaran_rutin',
    fn: 'getPengeluaranRutin',
    description: 'Perkiraan vs realisasi pengeluaran rutin bulan ini, beserta sisa pokok cicilan. periodeId opsional; kosong = periode aktif.',
    schema: { periodeId: { type: 'string', description: 'ID periode. Kosongkan untuk periode aktif.' } },
    args: (a) => [a.periodeId || null]
  },
  {
    name: 'kas_pantau_ir',
    fn: 'getPantauIR',
    description: 'Pantauan setoran Buku IR tiap jamaah beberapa bulan terakhir, termasuk siapa yang belum setor bulan ini.',
    schema: { jumlahBulan: { type: 'number', description: 'Jumlah bulan yang dipantau (default 6).' } },
    args: (a) => [a.jumlahBulan || 6]
  },
  {
    name: 'kas_rekap_setoran',
    fn: 'getRekapSetoran',
    description: 'Kewajiban setor ke desa/daerah per pos: target, realisasi, sisa atau kelebihan, beserta dasar perhitungannya.',
    args: () => []
  },
  {
    name: 'kas_rincian_ir',
    fn: 'getRincianIR',
    description: 'Rincian Buku IR periode tertentu per jamaah, dipecah menurut komponennya (IR, 1/10 IR, cicilan, infak daerah, index, kesehatan, sewa kos, qurban, lain-lain).',
    schema: { periodeId: { type: 'string', description: 'ID periode. Kosongkan untuk periode aktif.' } },
    args: (a) => [a.periodeId || null]
  },
  {
    name: 'kas_buku_ir',
    fn: 'getBukuIRData',
    description: 'Daftar transaksi Buku IR yang sudah dan belum dirincikan.',
    args: () => []
  },
  {
    name: 'kas_transaksi',
    fn: 'getSemuaTransaksi',
    description: 'Seluruh transaksi pemasukan & pengeluaran periode berjalan (yang dibatalkan tidak termasuk).',
    schema: { periodeId: { type: 'string', description: 'ID periode. Kosongkan untuk periode aktif.' } },
    args: (a) => [a.periodeId || null]
  },
  {
    name: 'kas_rekapitulasi',
    fn: 'getRekapitulasiData',
    description: 'Rekapitulasi periode: pemasukan dan pengeluaran per jenis, mutasi tunai/bank, saldo awal & akhir.',
    schema: { periodeId: { type: 'string', description: 'ID periode. Kosongkan untuk periode aktif.' } },
    args: (a) => [a.periodeId || null]
  },
  {
    name: 'kas_periode',
    fn: 'getAllPeriode',
    description: 'Daftar seluruh periode beserta statusnya (OPEN/CLOSED) dan saldo awal/akhir.',
    args: () => []
  },
  {
    name: 'kas_jamaah',
    fn: 'getAnggota',
    description: 'Daftar jamaah beserta nilai acuan setoran bulanannya.',
    args: () => []
  }
];

function cek() {
  const kurang = [];
  if (!EXEC_URL) kurang.push('KAS_EXEC_URL');
  if (!TOKEN) kurang.push('KAS_TOKEN');
  return kurang;
}

async function panggil(fn, args) {
  const res = await fetch(EXEC_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },  // hindari preflight GAS
    body: JSON.stringify({ token: TOKEN, fn, args }),
    redirect: 'follow'
  });
  const teks = await res.text();
  if (!res.ok) throw new Error(`HTTP ${res.status}: ${teks.slice(0, 300)}`);
  try {
    return JSON.parse(teks);
  } catch {
    // GAS mengembalikan halaman HTML bila URL salah atau deployment tak dapat diakses.
    throw new Error('Balasan bukan JSON. Periksa KAS_EXEC_URL dan akses deployment. Cuplikan: ' + teks.slice(0, 200));
  }
}

const server = new Server(
  { name: 'kas-kelompok', version: '1.0.0' },
  { capabilities: { tools: {} } }
);

server.setRequestHandler(ListToolsRequestSchema, async () => ({
  tools: TOOLS.map((t) => ({
    name: t.name,
    description: t.description + ' (baca-saja)',
    inputSchema: {
      type: 'object',
      properties: t.schema || {},
      required: []
    }
  }))
}));

server.setRequestHandler(CallToolRequestSchema, async (req) => {
  const alat = TOOLS.find((t) => t.name === req.params.name);
  if (!alat) {
    return { isError: true, content: [{ type: 'text', text: `Alat tidak dikenal: ${req.params.name}` }] };
  }
  const kurang = cek();
  if (kurang.length) {
    return { isError: true, content: [{ type: 'text', text: `Belum disetel: ${kurang.join(', ')}. Lihat mcp/README.md.` }] };
  }
  try {
    const hasil = await panggil(alat.fn, alat.args(req.params.arguments || {}));
    if (hasil && hasil.success === false) {
      return { isError: true, content: [{ type: 'text', text: `Ditolak aplikasi: ${hasil.message}` }] };
    }
    return { content: [{ type: 'text', text: JSON.stringify(hasil, null, 2) }] };
  } catch (err) {
    return { isError: true, content: [{ type: 'text', text: `Gagal: ${err.message}` }] };
  }
});

await server.connect(new StdioServerTransport());
