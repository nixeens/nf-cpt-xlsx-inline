<?php
/**
 * Plugin Name: NF CPT → XLSX Inline Export
 * Description: Export any public CPT to Excel (.xlsx) with PhpSpreadsheet (bundled vendor libraries, no Composer required).
 * Version: 1.1.0
 * Author: Your Name
 * License: GPL2
 */

if (!defined('ABSPATH')) exit;

define('NF_CPT_XLSX_INLINE_VERSION', '1.1.0');

/* ------------------------------------------------------------
 * 1️⃣ AUTOLOAD LOCAL LIBRARIES
 * ------------------------------------------------------------ */
function nf_xlsx_load_local_lib() {
    static $registered = false;
    if ($registered) {
        return;
    }

    $base = plugin_dir_path(__FILE__) . 'lib/';

    // PhpSpreadsheet
    spl_autoload_register(function ($class) use ($base) {
        $prefix = 'PhpOffice\\PhpSpreadsheet\\';
        $len = strlen($prefix);
        if (strncmp($prefix, $class, $len) !== 0) {
            return;
        }

        $relative_class = substr($class, $len);
        $file = $base . 'PhpOffice/PhpSpreadsheet/' . str_replace('\\', '/', $relative_class) . '.php';
        if (file_exists($file)) {
            require_once $file;
        }
    });

    // ✅ PSR SimpleCache polyfill
    spl_autoload_register(function ($class) use ($base) {
        if ($class === 'Psr\\SimpleCache\\CacheInterface') {
            $file = $base . 'Psr/SimpleCache/CacheInterface.php';
            if (file_exists($file)) {
                require_once $file;
            }
        }
    });

    // ✅ Composer\Pcre polyfill
    spl_autoload_register(function ($class) use ($base) {
        if ($class === 'Composer\\Pcre\\Preg') {
            $file = $base . 'Composer/Pcre/Preg.php';
            if (file_exists($file)) {
                require_once $file;
            }
        }
    });

    $registered = true;
}

add_action('plugins_loaded', 'nf_xlsx_load_local_lib', 1);

/* ------------------------------------------------------------
 * 2️⃣ IMPORT CLASSES
 * ------------------------------------------------------------ */
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/* ------------------------------------------------------------
 * 3️⃣ ADMIN MENU
 * ------------------------------------------------------------ */
add_action('admin_menu', function () {
    add_menu_page(
        __('NF CPT → XLSX Export', 'nf-cpt-xlsx-inline'),
        __('NF → XLSX', 'nf-cpt-xlsx-inline'),
        'manage_options',
        'nf-cpt-xlsx-inline',
        'nf_cpt_xlsx_inline_admin_page',
        'dashicons-media-spreadsheet'
    );
});

/* ------------------------------------------------------------
 * 4️⃣ PAGE
 * ------------------------------------------------------------ */
function nf_cpt_xlsx_inline_admin_page() {
    if (!current_user_can('manage_options')) {
        wp_die(__('You are not allowed to access this page.', 'nf-cpt-xlsx-inline'));
    }

    $post_types = nf_cpt_xlsx_inline_get_exportable_post_types();
    $current = isset($_GET['nf_cpt']) ? sanitize_key(wp_unslash($_GET['nf_cpt'])) : 'post';
    if (!isset($post_types[$current])) {
        $current = 'post';
    }

    ?>
    <div class="wrap">
        <h1><?php esc_html_e('NF CPT → XLSX Export', 'nf-cpt-xlsx-inline'); ?></h1>
        <p><?php esc_html_e('Generate an Excel spreadsheet with entries from your chosen custom post type. The export is generated inline with bundled PhpSpreadsheet libraries (no Composer required).', 'nf-cpt-xlsx-inline'); ?></p>

        <form method="post" action="<?php echo esc_url(admin_url('admin-post.php')); ?>">
            <?php wp_nonce_field('nf_cpt_xlsx_inline_export', '_nf_cpt_nonce'); ?>
            <input type="hidden" name="action" value="nf_cpt_export_xlsx" />

            <table class="form-table" role="presentation">
                <tbody>
                    <tr>
                        <th scope="row"><label for="nf-cpt-xlsx-post-type"><?php esc_html_e('Post Type', 'nf-cpt-xlsx-inline'); ?></label></th>
                        <td>
                            <select id="nf-cpt-xlsx-post-type" name="post_type">
                                <?php foreach ($post_types as $slug => $label) : ?>
                                    <option value="<?php echo esc_attr($slug); ?>" <?php selected($slug, $current); ?>><?php echo esc_html($label); ?></option>
                                <?php endforeach; ?>
                            </select>
                            <p class="description"><?php esc_html_e('Only public post types are listed. Use the filter "nf_cpt_xlsx_inline_post_types" to modify.', 'nf-cpt-xlsx-inline'); ?></p>
                        </td>
                    </tr>
                    <tr>
                        <th scope="row"><?php esc_html_e('Fields', 'nf-cpt-xlsx-inline'); ?></th>
                        <td>
                            <p class="description"><?php esc_html_e('The export contains ID, Title, Status, Author, Date and Permalink columns by default. Hook into "nf_cpt_xlsx_inline_headers" or "nf_cpt_xlsx_inline_row" to customize.', 'nf-cpt-xlsx-inline'); ?></p>
                        </td>
                    </tr>
                </tbody>
            </table>

            <?php submit_button(__('Export to XLSX', 'nf-cpt-xlsx-inline')); ?>
        </form>
    </div>
    <?php
}

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
        wp_die(__('You are not allowed to export data.', 'nf-cpt-xlsx-inline'));
    }

    check_admin_referer('nf_cpt_xlsx_inline_export', '_nf_cpt_nonce');

    $post_type = isset($_REQUEST['post_type']) ? sanitize_key(wp_unslash($_REQUEST['post_type'])) : 'post';
    $post_types = nf_cpt_xlsx_inline_get_exportable_post_types();
    if (!isset($post_types[$post_type])) {
        $post_type = 'post';
    }

    try {
        [$headers, $rows] = nf_cpt_xlsx_inline_dataset($post_type);

        if (empty($headers)) {
            $headers = ['Message'];
            $rows = [['No data available for export.']];
        }

        $spreadsheet = new Spreadsheet();
        $spreadsheet->getProperties()
            ->setCreator(get_bloginfo('name'))
            ->setTitle(sprintf(__('Export: %s', 'nf-cpt-xlsx-inline'), $post_types[$post_type]))
            ->setDescription(__('Generated with NF CPT → XLSX Inline Export', 'nf-cpt-xlsx-inline'));

        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setTitle(apply_filters('nf_cpt_xlsx_inline_sheet_title', 'Data Export', $post_type));

        // Header row.
        $rowIndex = 1;
        foreach (array_values($headers) as $colIndex => $header) {
            $sheet->setCellValueByColumnAndRow($colIndex + 1, $rowIndex, $header);
            $sheet->getStyleByColumnAndRow($colIndex + 1, $rowIndex)->getFont()->setBold(true);
        }

        // Data rows.
        $rowIndex = 2;
        foreach ($rows as $row) {
            $colIndex = 1;
            foreach ($headers as $key => $label) {
                $value = isset($row[$key]) ? $row[$key] : '';
                $sheet->setCellValueByColumnAndRow($colIndex, $rowIndex, $value);
                ++$colIndex;
            }
            ++$rowIndex;
        }

        // Autosize columns for readability.
        foreach (range(1, count($headers)) as $column) {
            $sheet->getColumnDimensionByColumn($column)->setAutoSize(true);
        }

        // Freeze top row (headers).
        $sheet->freezePane('A2');

        // Styling (optional border for data table).
        if (!empty($rows) && class_exists('PhpOffice\\PhpSpreadsheet\\Style\\Border')) {
            $lastColumnLetter = nf_xlsx_colrow_to_address(count($headers), 1);
            $lastColumnLetter = preg_replace('/\d+/', '', $lastColumnLetter);
            $range = sprintf('A1:%s%d', $lastColumnLetter, count($rows) + 1);
            $sheet->getStyle($range)
                ->getBorders()
                ->getAllBorders()
                ->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
        }

        while (ob_get_level()) {
            ob_end_clean();
        }

        nocache_headers();
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="' . sanitize_file_name($post_type . '-export.xlsx') . '"');
        header('Cache-Control: max-age=0, no-cache, no-store, must-revalidate');

        $writer = new Xlsx($spreadsheet);
        $writer->save('php://output');

        $spreadsheet->disconnectWorksheets();
        exit;
    } catch (Throwable $e) {
        error_log('NF XLSX EXPORT ERROR: ' . $e->getMessage());
        wp_die('<strong>' . esc_html__('NF XLSX Export Error:', 'nf-cpt-xlsx-inline') . '</strong><br>' . esc_html($e->getMessage()));
    }
}

/* ------------------------------------------------------------
 * 7️⃣ DATASET HELPERS
 * ------------------------------------------------------------ */

function nf_cpt_xlsx_inline_get_exportable_post_types() {
    $post_types = get_post_types(
        [
            'public'  => true,
            'show_ui' => true,
        ],
        'objects'
    );

    $choices = [];
    foreach ($post_types as $type) {
        $choices[$type->name] = $type->labels->name;
    }

    /**
     * Filter the list of post types available for export.
     *
     * @param array $choices Key => label pairs.
     */
    return apply_filters('nf_cpt_xlsx_inline_post_types', $choices);
}

function nf_cpt_xlsx_inline_dataset($post_type) {
    $args = [
        'post_type'      => $post_type,
        'post_status'    => 'any',
        'posts_per_page' => apply_filters('nf_cpt_xlsx_inline_posts_per_page', -1, $post_type),
        'orderby'        => 'date',
        'order'          => 'DESC',
    ];

    $args = apply_filters('nf_cpt_xlsx_inline_query_args', $args, $post_type);

    $posts = get_posts($args);

    $headers = [
        'id'        => __('ID', 'nf-cpt-xlsx-inline'),
        'title'     => __('Title', 'nf-cpt-xlsx-inline'),
        'status'    => __('Status', 'nf-cpt-xlsx-inline'),
        'author'    => __('Author', 'nf-cpt-xlsx-inline'),
        'date'      => __('Date', 'nf-cpt-xlsx-inline'),
        'permalink' => __('Permalink', 'nf-cpt-xlsx-inline'),
    ];

    $headers = apply_filters('nf_cpt_xlsx_inline_headers', $headers, $post_type);

    $rows = [];
    foreach ($posts as $post) {
        $row = [
            'id'        => $post->ID,
            'title'     => get_the_title($post),
            'status'    => $post->post_status,
            'author'    => get_the_author_meta('display_name', $post->post_author),
            'date'      => get_the_date('', $post),
            'permalink' => get_permalink($post),
        ];

        $row = apply_filters('nf_cpt_xlsx_inline_row', $row, $post, $post_type);

        // Guarantee all headers exist as keys.
        foreach ($headers as $key => $label) {
            if (!array_key_exists($key, $row)) {
                $row[$key] = '';
            }
        }

        $rows[] = $row;
    }

    return [$headers, $rows];
}

