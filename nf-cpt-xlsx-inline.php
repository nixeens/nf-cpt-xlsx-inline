<?php
/**
 * Plugin Name: NF CPT → XLSX Inline Export
 * Description: Exports CPT or form data to Excel (.xlsx) using bundled PhpSpreadsheet library (composer-free).
 * Version: 1.0.6
 * Author: Your Name
 * License: GPL2
 */

if (!defined('ABSPATH')) exit;

define('NF_CPT_XLSX_INLINE_VERSION', '1.0.6');

/* ------------------------------------------------------------
 * 1️⃣ AUTOLOAD LOCAL LIBRARIES
 * ------------------------------------------------------------ */
function nf_xlsx_load_local_lib() {
    if (class_exists('PhpOffice\\PhpSpreadsheet\\Spreadsheet', false)) {
        return;
    }

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
        <a href="<?php echo esc_url(wp_nonce_url(admin_url('admin-post.php?action=nf_cpt_export_xlsx'), 'nf_cpt_xlsx_inline_export')); ?>" class="button button-primary">Export Test XLSX</a>
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
    if (!current_user_can('manage_options')) {
        wp_die(__('You do not have permission to export this data.', 'nf-cpt-xlsx-inline'));
    }

    check_admin_referer('nf_cpt_xlsx_inline_export');

    try {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setTitle('Data Export');

        $data = nf_cpt_xlsx_inline_get_export_data();

        // Header
        $headers = isset($data['headers']) && is_array($data['headers']) ? $data['headers'] : [];
        $row = 1; $col = 1;
        foreach ($headers as $header) {
            $cell = nf_xlsx_colrow_to_address($col, $row);
            $sheet->setCellValue($cell, $header);
            $sheet->getStyle($cell)->getFont()->setBold(true);
            $col++;
        }

        // Rows
        $rows = isset($data['rows']) && is_array($data['rows']) ? $data['rows'] : [];
        $r = 2;
        foreach ($rows as $dataRow) {
            $c = 1;
            foreach ($dataRow as $val) {
                $cell = nf_xlsx_colrow_to_address($c, $r);
                $sheet->setCellValueExplicit($cell, $val, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
                $c++;
            }
            $r++;
        }

        // Attach images if provided
        if (!empty($data['images']) && is_array($data['images'])) {
            foreach ($data['images'] as $imageCell => $imagePath) {
                $realPath = apply_filters('nf_cpt_xlsx_inline_image_path', $imagePath, $imageCell);
                if ($realPath && file_exists($realPath)) {
                    $drawing = new Drawing();
                    $drawing->setPath($realPath);
                    $drawing->setCoordinates($imageCell);
                    $drawing->setHeight(80);
                    $drawing->setWorksheet($sheet);
                }
            }
        }

        // Column widths
        $columnCount = max(count($headers), 1);
        for ($i = 1; $i <= $columnCount; $i++) {
            $l = nf_xlsx_colrow_to_address($i, 1);
            $l = preg_replace('/\d+$/', '', $l);
            $sheet->getColumnDimension($l)->setAutoSize(true);
        }

        // Borders
        $lastRow = count($rows) + 1;
        if ($columnCount > 0 && class_exists('PhpOffice\\PhpSpreadsheet\\Style\\Border')) {
            $endColumn = nf_xlsx_colrow_to_address($columnCount, 1);
            $endColumn = preg_replace('/\d+$/', '', $endColumn);
            $sheet->getStyle("A1:{$endColumn}{$lastRow}")
                ->getBorders()
                ->getAllBorders()
                ->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
        }

        // Output
        nocache_headers();
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

function nf_cpt_xlsx_inline_get_export_data() {
    $headers = ['ID', 'Title', 'Author', 'Published'];
    $rows = [];

    if (function_exists('get_posts')) {
        $posts = get_posts([
            'post_type'      => 'any',
            'posts_per_page' => 10,
            'orderby'        => 'date',
            'order'          => 'DESC',
        ]);

        foreach ($posts as $post) {
            $rows[] = [
                (string) $post->ID,
                get_the_title($post),
                get_the_author_meta('display_name', $post->post_author),
                get_the_date('', $post),
            ];
        }
    }

    if (empty($rows)) {
        $rows = [
            ['1', 'John Doe', 'john@example.com', ''],
            ['2', 'Jane Roe', 'jane@example.com', ''],
        ];
    }

    return apply_filters('nf_cpt_xlsx_inline_export_data', [
        'headers' => $headers,
        'rows'    => $rows,
        'images'  => [],
    ]);
}

