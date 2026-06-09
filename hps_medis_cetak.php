<?php
include 'koneksi/konek.php';

$kode = isset($_REQUEST['kode']) ? $_REQUEST['kode'] : '';
if (empty($kode)) {
    die('Parameter kode tidak ditemukan.');
}

// Query header dokumen
$kode_esc = mysql_real_escape_string($kode);

$s_po = "SELECT
            nd.no_dinas,
            nd.tgl,
            nd.sifat,
            nd.ttd_dari_set,
            kp.jabatan AS kepada,
            dr.jabatan AS dari,
            dr.file,
            dr.nama,
            dr.jenis_nip,
            dr.nip,
            br.namabarang,
            sm.keterangan AS sumber_dana,
            nd.kode_program,
            br.kode_anggaran,
            nd.kode_rup,
            nd.jml_pagu_anggaran,
            nd.jangka_waktu,
            pn.nama AS metode_pemilihan
        FROM
            as_nota_dinas_obat nd
        INNER JOIN as_ms_penanda_tangan kp ON kp.id = nd.kepada
        INNER JOIN as_ms_penanda_tangan dr ON dr.id = nd.dari
        INNER JOIN as_ms_barang br ON br.idbarang = nd.kategori_barang_id
        INNER JOIN as_ms_penyedia pn ON pn.id = nd.metode_pemilihan
        INNER JOIN as_ms_sumberdana sm ON sm.idsumberdana = nd.sumberdana_id
        WHERE nd.kode_nd = '$kode_esc'";

$q_po = mysql_query($s_po);
if (!$q_po) {
    die('Query header gagal: ' . mysql_error());
}
$d_po = mysql_fetch_array($q_po);
if (!$d_po) {
    die('Data tidak ditemukan untuk kode: ' . htmlspecialchars($kode));
}

// Query detail item dari as_po_ext
$s_items = "SELECT
                pe.id,
                ab.obat_nama AS namabarang,
                pe.satuan,
                pe.qty_satuan as jml,
                pe.harga_satuan,
                pe.ongkir_satuan,
                pe.ppn,
                pe.nilai_ppn,
                pe.status_ppn,
                pe.deskripsi
            FROM as_po_obat_ext pe
            LEFT JOIN rsmn_apotek.a_obat ab ON ab.obat_id = pe.barang_id 
            WHERE pe.kode_nd = '$kode_esc'
            ORDER BY pe.id ASC";

$q_items = mysql_query($s_items);
if (!$q_items) {
    die('Query items gagal: ' . mysql_error());
}

// ----------------------------------------------------------------
// Proses setiap item: hitung harga_netto, pajak_label, total
// berdasarkan status_ppn
// ----------------------------------------------------------------
// status_ppn:
//   0 = tidak kena PPN → harga netto = harga_awal + ongkir, pajak = 0%
//   1 = kena PPN, harga BELUM termasuk PPN → netto = harga_awal + ongkir, total = netto × (1 + ppn/100) × jml
//   2 = kena PPN, harga SUDAH termasuk PPN → netto (DPP) = (harga_awal + ongkir) / (1 + ppn/100), total = (harga_awal + ongkir) × jml
// ----------------------------------------------------------------
$items      = array();
$grand_total = 0;
$no          = 1;

while ($row = mysql_fetch_array($q_items, MYSQL_ASSOC)) {
    $harga_raw  = floatval($row['harga_satuan']) + floatval($row['ongkir_satuan']);
    $ppn        = intval($row['nilai_ppn']);
    $status_ppn = intval($row['status_ppn']);
    $jml        = floatval($row['jml']);

    switch ($status_ppn) {
        case 0:
            $harga_netto = $harga_raw;
            $pajak_label = 0;
            $total       = $harga_netto * $jml;
            break;

        case 1:
        case 2:
          $harga_netto = $harga_raw;
          $pajak_label = $ppn;
          $total       = $harga_netto * (1 + $ppn / 100) * $jml;
          break;

        default:
            $harga_netto = $harga_raw;
            $pajak_label = 0;
            $total       = $harga_netto * $jml;
    }

    $grand_total += $total;

    $items[] = array(
        'no'          => $no++,
        'nama'        => $row['namabarang'] ? $row['namabarang'] : $row['deskripsi'],
        'satuan'      => $row['satuan'],
        'jml'         => $jml,
        'harga_netto' => $harga_netto,
        'pajak'       => $pajak_label,
        'total'       => $total,
        'status_ppn'  => $status_ppn,
    );
}

// Info header untuk tampilan
$info = array(
    'instansi'      => 'PEMERINTAH PROVINSI JAWA TIMUR',
    'unit'          => 'RSUD MOHAMMAD NOER',
    'alamat'        => 'JL. BONOROGO NO. 17',
    'kota'          => 'PAMEKASAN',
    'sumber_dana'   => $d_po['sumber_dana'],
    'prog_keg'      => $d_po['kode_program'],
    'kode_rekening' => $d_po['kode_anggaran'],
    'paket_belanja' => $d_po['namabarang'],
    'judul_rincian' => 'RINCIAN KEBUTUHAN (' . strtoupper($d_po['namabarang']) . ')',
);

// ================================================================
// Handle download Excel
// ================================================================
if (isset($_GET['download']) && $_GET['download'] == 'excel') {
    require_once('include/PHPExcel/Classes/PHPExcel.php');
    require_once('include/PHPExcel/Classes/PHPExcel/IOFactory.php');

    $objPHPExcel = new PHPExcel();
    $objPHPExcel->setActiveSheetIndex(0);
    $objPHPExcel->getDefaultStyle()
      ->getFont()
      ->setName('Arial')
      ->setSize(10);
    $ws = $objPHPExcel->getActiveSheet();
    $ws->setTitle('hps rl pz');

    $thin      = array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('argb' => 'FF000000'));
    $borderAll = array('borders' => array('allborders' => $thin));

    // --- Header baris 1-7 ---
    $ws->setCellValue('A1', $info['instansi'] . ' HARGA PERKIRAAN SENDIRI');
    $ws->setCellValue('A2', $info['unit']     . ' SUMBER DANA : ' . $info['sumber_dana']);
    $ws->setCellValue('A3', $info['alamat']   . ' PROG/KEG/SUB KEGIATAN : ' . $info['prog_keg']);
    $ws->setCellValue('A4', $info['kota']     . ' KODE REKENING : ' . $info['kode_rekening']);
    $ws->setCellValue('A5', 'PAKET BELANJA : ' . $info['paket_belanja']);
    $ws->setCellValue('A7', $info['judul_rincian']);

    foreach (range(1, 8) as $r) {
        $ws->getStyle('A' . $r)->getFont()->setName('Arial')->setSize(10);
        $ws->getStyle('A' . $r)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $ws->getRowDimension($r)->setRowHeight(15.75);
    }

    // --- Header tabel (baris 9) ---
    $headers = array(
        'A9' => 'No', 'B9' => 'Jenis Barang/Jasa', 'C9' => 'Satuan',
        'D9' => 'Vol', 'E9' => 'Harga', 'F9' => 'Pajak (%)', 'G9' => 'Total', 'H9' => 'Keterangan',
    );
    foreach ($headers as $cell => $val) {
        $ws->setCellValue($cell, $val);
    }
    $ws->getStyle('A9:H9')->applyFromArray(array_merge($borderAll, array(
        'font'      => array('name' => 'Arial', 'size' => 10, 'bold' => true),
        'alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
            'vertical'   => PHPExcel_Style_Alignment::VERTICAL_CENTER,
        ),
    )));
    $ws->getRowDimension(9)->setRowHeight(15.75);

    // --- Baris data ---
    $startRow = 10;
    foreach ($items as $idx => $item) {
        $row   = $startRow + $idx;
        $pajak = $item['pajak'];

        $ws->setCellValue('A' . $row, $item['no']);
        $ws->setCellValue('B' . $row, $item['nama']);
        $ws->setCellValue('C' . $row, $item['satuan']);
        $ws->setCellValue('D' . $row, $item['jml']);
        $ws->setCellValue('E' . $row, $item['harga_netto']);
        $ws->setCellValue('F' . $row, $pajak > 0 ? $pajak : '-');
        $ws->setCellValue('G' . $row, $item['total']);
        $ws->setCellValue('H' . $row, '');

        $ws->getStyle('E' . $row)->getNumberFormat()->setFormatCode('#,##0.00');
        $ws->getStyle('G' . $row)->getNumberFormat()->setFormatCode('#,##0.00');
        $ws->getStyle('D' . $row)->getNumberFormat()->setFormatCode('#,##0');
        $ws->getStyle('A' . $row . ':H' . $row)->applyFromArray($borderAll);
        $ws->getStyle('A' . $row)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $ws->getStyle('F' . $row)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $ws->getStyle('A' . $row)->getFont()->setName('Arial')->setSize(10);
        $ws->getRowDimension($row)->setRowHeight($idx === 0 ? 41.4 : 26.4);
    }

    // Baris TOTAL
    $totalRow  = $startRow + count($items);
    $dataStart = $startRow;
    $dataEnd   = $startRow + count($items) - 1;
    $ws->setCellValue('A' . $totalRow, count($items) + 1);
    $ws->setCellValue('B' . $totalRow, 'TOTAL');
    $ws->mergeCells('B' . $totalRow . ':F' . $totalRow);
    $ws->setCellValue('G' . $totalRow, '=SUM(G' . $dataStart . ':G' . $dataEnd . ')');
    $ws->getStyle('G' . $totalRow)->getNumberFormat()->setFormatCode('#,##0.00');

    $ws->getStyle('A' . $totalRow)->applyFromArray($borderAll);
    $ws->getStyle('B' . $totalRow)->applyFromArray($borderAll);
    $ws->getStyle('C' . $totalRow . ':E' . $totalRow)->applyFromArray(array(
        'borders' => array('top' => $thin, 'bottom' => $thin)));
    $ws->getStyle('F' . $totalRow)->applyFromArray(array(
        'borders' => array('top' => $thin, 'bottom' => $thin, 'right' => $thin)));
    $ws->getStyle('G' . $totalRow)->applyFromArray($borderAll);
    $ws->getStyle('H' . $totalRow)->applyFromArray($borderAll);
    $ws->getStyle('A' . $totalRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $ws->getStyle('B' . $totalRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $ws->getRowDimension($totalRow)->setRowHeight(15.75);

    // Lebar kolom
    $ws->getColumnDimension('A')->setWidth(5.89);
    $ws->getColumnDimension('B')->setWidth(18.66);
    $ws->getColumnDimension('C')->setWidth(8.0);
    $ws->getColumnDimension('D')->setWidth(10.0);
    $ws->getColumnDimension('E')->setWidth(17.78);
    $ws->getColumnDimension('F')->setWidth(20.11);
    $ws->getColumnDimension('G')->setWidth(25.33);
    $ws->getColumnDimension('H')->setWidth(22.11);

    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="Rincian_HPS_' . $kode_esc . '.xlsx"');
    header('Cache-Control: max-age=0');

    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
    $objWriter->save('php://output');
    exit;
}
?>

<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Preview HPS - <?php echo htmlspecialchars($info['paket_belanja']); ?></title>
<style>
  body {
    font-family: Arial, sans-serif;
    font-size: 10pt;
    background: #e8ecef;
    margin: 0;
    padding: 20px;
  }
  .toolbar {
    max-width: 900px;
    margin: 0 auto 14px auto;
    display: flex;
    gap: 10px;
    align-items: center;
  }
  .btn {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 8px 18px;
    border: none;
    border-radius: 4px;
    cursor: pointer;
    font-size: 13px;
    font-family: Arial, sans-serif;
    font-weight: bold;
    text-decoration: none;
  }
  .btn-print { background: #1a73e8; color: #fff; }
  .btn-print:hover { background: #1558b0; }
  .btn-excel { background: #217346; color: #fff; }
  .btn-excel:hover { background: #155230; }

  .paper {
    background: #fff;
    max-width: 900px;
    margin: 0 auto;
    padding: 24px 28px;
    box-shadow: 0 2px 10px rgba(0,0,0,0.15);
  }
  .doc-header {
    width: 100%;
    border-collapse: collapse;
    margin-bottom: 10px;
    font-size: 10pt;
  }
  .doc-header td { padding: 1px 0; vertical-align: middle; }
  .doc-header .kol-kiri  { width: 45%; }
  .doc-header .kol-kanan { width: 55%; }
  .doc-header .separator { width: 10px; }

  .judul-rincian { font-size: 10pt; margin: 8px 0 6px 0; }

  table.hps { width: 100%; border-collapse: collapse; font-size: 10pt; }
  table.hps th, table.hps td {
    border: 1px solid #000;
    padding: 3px 5px;
    vertical-align: middle;
  }
  table.hps thead th { background: #fff; text-align: center; font-weight: bold; }
  table.hps td.center { text-align: center; }
  table.hps td.right  { text-align: right; }

  .total-label { text-align: center; font-weight: bold; }
  .no-left-border  { border-left: none !important; }
  .no-right-border { border-right: none !important; }

  .ppn-badge {
    display: inline-block;
    font-size: 8pt;
    padding: 0 3px;
    border-radius: 2px;
    margin-left: 3px;
    vertical-align: middle;
    color: #fff;
  }
  .ppn-0 { background: #888; }
  .ppn-1 { background: #c0392b; }
  .ppn-2 { background: #27ae60; }

  @media print {
    body     { background: none; padding: 0; }
    .toolbar { display: none; }
    .paper   { box-shadow: none; padding: 0; max-width: 100%; }
    .ppn-badge { display: none; }
    @page    { margin: 1.5cm; }
  }
</style>
</head>
<body>

<div class="toolbar">
  <button class="btn btn-print" onclick="window.print()">Cetak</button>
  <a class="btn btn-excel" href="?kode=<?php echo urlencode($kode); ?>&download=excel">Download Excel</a>
</div>

<div class="paper">

  <table class="doc-header">
    <tr>
      <td class="kol-kiri"><?php echo htmlspecialchars($info['instansi']); ?></td>
      <td class="separator"></td>
      <td class="kol-kanan">HARGA PERKIRAAN SENDIRI</td>
    </tr>
    <tr>
      <td><?php echo htmlspecialchars($info['unit']); ?></td>
      <td></td>
      <td>SUMBER DANA : <?php echo htmlspecialchars($info['sumber_dana']); ?></td>
    </tr>
    <tr>
      <td><?php echo htmlspecialchars($info['alamat']); ?></td>
      <td></td>
      <td>PROG/KEG/SUB KEGIATAN : <?php echo htmlspecialchars($info['prog_keg']); ?></td>
    </tr>
    <tr>
      <td><?php echo htmlspecialchars($info['kota']); ?></td>
      <td></td>
      <td>KODE REKENING : <?php echo htmlspecialchars($info['kode_rekening']); ?></td>
    </tr>
    <tr>
      <td>PAKET BELANJA : <?php echo htmlspecialchars($info['paket_belanja']); ?></td>
      <td></td>
      <td></td>
    </tr>
  </table>

  <p class="judul-rincian"><?php echo htmlspecialchars($info['judul_rincian']); ?></p>

  <table class="hps">
    <thead>
      <tr>
        <th style="width:4%">No</th>
        <th style="width:24%">Jenis Barang/Jasa</th>
        <th style="width:7%">Satuan</th>
        <th style="width:8%">Vol</th>
        <th style="width:14%">Harga</th>
        <th style="width:10%">Pajak (%)</th>
        <th style="width:15%">Total</th>
        <th style="width:18%">Keterangan</th>
      </tr>
    </thead>
    <tbody>
      <?php if (empty($items)): ?>
      <tr>
        <td colspan="8" style="text-align:center; color:#888;">Tidak ada data item.</td>
      </tr>
      <?php else: ?>
      <?php foreach ($items as $item): ?>
      <tr>
        <td class="center"><?php echo $item['no']; ?></td>
        <td><?php echo htmlspecialchars($item['nama']); ?></td>
        <td class="center"><?php echo htmlspecialchars($item['satuan']); ?></td>
        <td class="right"><?php echo number_format($item['jml'], 0, ',', '.'); ?></td>
        <td class="right"><?php echo number_format($item['harga_netto'], 2, ',', '.'); ?></td>
        <td class="center">
          <?php if ($item['pajak'] > 0): ?>
            <?php echo $item['pajak']; ?>%
            <?php /* badge status_ppn hanya tampil di layar, tersembunyi saat cetak */ ?>
            <?php if ($item['status_ppn'] == 1): ?>
              <span class="ppn-badge ppn-1" title="Harga belum termasuk PPN">+PPN</span>
            <?php elseif ($item['status_ppn'] == 2): ?>
              <span class="ppn-badge ppn-2" title="Harga sudah termasuk PPN">incl.</span>
            <?php endif; ?>
          <?php else: ?>
            -
          <?php endif; ?>
        </td>
        <td class="right"><?php echo number_format($item['total'], 2, ',', '.'); ?></td>
        <td></td>
      </tr>
      <?php endforeach; ?>
      <?php endif; ?>

      <!-- Baris TOTAL -->
      <tr>
        <td class="center"><?php echo count($items) + 1; ?></td>
        <td class="total-label no-right-border">TOTAL</td>
        <td class="no-left-border no-right-border"></td>
        <td class="no-left-border no-right-border"></td>
        <td class="no-left-border no-right-border"></td>
        <td class="no-left-border"></td>
        <td class="right"><strong><?php echo number_format($grand_total, 2, ',', '.'); ?></strong></td>
        <td></td>
      </tr>
    </tbody>
  </table>

  </div>

</body>
</html>