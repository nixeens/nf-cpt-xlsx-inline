<?php
/**
 * Plugin Name: NF CPT → XLSX Inline Export
 * Description: Exports CPT or form data to Excel (.xlsx) using bundled PhpSpreadsheet library (composer-free).
 * Version: 1.0.5
 * Author: Your Name
 * License: GPL2
 */

if (!defined('ABSPATH')) exit;

/* ------------------------------------------------------------
 * 1️⃣ AUTOLOAD LOCAL LIBRARIES
 * ------------------------------------------------------------ */
function nf_xlsx_load_local_lib() {
    $base = plugin_dir_path(__FILE__) . 'lib/';

    // PhpSpreadsheet
    spl_autoload_register(function ($class) use ($base) {
        $prefix = 'PhpOffice\\PhpSpreadsheet\\';
        $len = strlen($prefix);
        if (strncmp($prefix, $class, $len) !== 0) return;
        $relative_class = substr($class, $len);
        $file = $base . 'PhpOffice/PhpSpreadsheet/' . str_replace('\\', '/', $relative_class) . '.php';
        if (file_exists($file)) require_once $file;
    });

    // ✅ PSR SimpleCache polyfill
    spl_autoload_register(function ($class) use ($base) {
        if ($class === 'Psr\\SimpleCache\\CacheInterface') {
            $file = $base . 'Psr/SimpleCache/CacheInterface.php';
            if (file_exists($file)) require_once $file;
        }
    });

    // ✅ Composer\Pcre polyfill
    spl_autoload_register(function ($class) use ($base) {
        if ($class === 'Composer\\Pcre\\Preg') {
            $file = $base . 'Composer/Pcre/Preg.php';
            if (file_exists($file)) require_once $file;
        }
    });
}
nf_xlsx_load_local_lib();

/* ------------------------------------------------------------
 * 2️⃣ IMPORT CLASSES
 * ------------------------------------------------------------ */
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

/* ------------------------------------------------------------
 * 3️⃣ ADMIN MENU
 * ------------------------------------------------------------ */
add_action('admin_menu', function () {
    add_menu_page('NF CPT → XLSX Export', 'NF → XLSX', 'manage_options', 'nf-cpt-xlsx-inline', 'nf_cpt_xlsx_inline_admin_page');
});

/* ------------------------------------------------------------
 * 4️⃣ PAGE
 * ------------------------------------------------------------ */
function nf_cpt_xlsx_inline_admin_page() { ?>
    <div class="wrap">
        <h1>NF CPT → XLSX Export</h1>
        <p>Click below to generate a sample Excel export.</p>
        <a href="<?php echo esc_url(admin_url('admin-post.php?action=nf_cpt_export_xlsx')); ?>" class="button button-primary">Export Test XLSX</a>
    </div>
<?php }

/* ------------------------------------------------------------
 * 5️⃣ HELPER: Convert col/row to address
 * ------------------------------------------------------------ */
function nf_xlsx_colrow_to_address($col, $row) {
    $letter = '';
    while ($col > 0) {
        $mod = ($col - 1) % 26;
        $letter = chr(65 + $mod) . $letter;
        $col = (int)(($col - $mod) / 26);
    }
    return $letter . $row;
}

/* ------------------------------------------------------------
 * 6️⃣ EXPORT HANDLER
 * ------------------------------------------------------------ */
add_action('admin_post_nf_cpt_export_xlsx', 'nf_cpt_xlsx_inline_export');

function nf_cpt_xlsx_inline_export() {
    try {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setTitle('Data Export');

        // Header
        $headers = ['ID', 'Name', 'Email', 'Image'];
        $row = 1; $col = 1;
        foreach ($headers as $header) {
            $cell = nf_xlsx_colrow_to_address($col, $row);
            $sheet->setCellValue($cell, $header);
            $sheet->getStyle($cell)->getFont()->setBold(true);
            $col++;
        }

        // Rows
        $rows = [
            [1, 'John Doe', 'john@example.com'],
            [2, 'Jane Roe', 'jane@example.com'],
        ];
        $r = 2;
        foreach ($rows as $data) {
            $c = 1;
            foreach ($data as $val) {
                $cell = nf_xlsx_colrow_to_address($c, $r);
                $sheet->setCellValue($cell, $val);
                $c++;
            }
            $r++;
        }

        // Image
        $image = WP_CONTENT_DIR . '/uploads/sample.jpg';
        if (file_exists($image)) {
            $drawing = new Drawing();
            $drawing->setPath($image);
            $drawing->setCoordinates('D2');
            $drawing->setHeight(80);
            $drawing->setWorksheet($sheet);
        }

        // Column widths
        foreach (['A','B','C','D'] as $l) {
            $sheet->getColumnDimension($l)->setAutoSize(true);
        }

        // Borders
        $lastRow = count($rows) + 1;
        if (class_exists('PhpOffice\\PhpSpreadsheet\\Style\\Border')) {
            $sheet->getStyle("A1:D{$lastRow}")
                ->getBorders()
                ->getAllBorders()
                ->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
        }

        // Output
        while (ob_get_level()) ob_end_clean();
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="nf-export.xlsx"');
        header('Cache-Control: max-age=0, no-cache, no-store, must-revalidate');

        $writer = new Xlsx($spreadsheet);
        $writer->save('php://output');

        $spreadsheet->disconnectWorksheets();
        exit;
    } catch (Throwable $e) {
        error_log('NF XLSX EXPORT ERROR: '.$e->getMessage());
        wp_die('<strong>NF XLSX Export Error:</strong><br>'.esc_html($e->getMessage()));
    }
}

