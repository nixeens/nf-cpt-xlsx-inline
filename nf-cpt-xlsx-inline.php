<?php
/**
 * Plugin Name: NF CPT → XLSX Inline Export
 * Description: Export Ninja Forms submissions to Excel (.xlsx) with bundled PhpSpreadsheet library.
 * Version: 2.0.0
 * Author: Your Name
 * License: GPL2
 */

if (!defined('ABSPATH')) {
    exit;
}

define('NF_CPT_XLSX_INLINE_SLUG', 'nf-cpt-xlsx-inline');
define('NF_CPT_XLSX_INLINE_TEXT_DOMAIN', 'nf-cpt-xlsx-inline');

/* ------------------------------------------------------------
 * 1️⃣ AUTOLOAD LOCAL LIBRARIES
 * ------------------------------------------------------------ */
function nf_xlsx_load_local_lib() {
    static $registered = false;

    if ($registered) {
        return;
    }

    $base = plugin_dir_path(__FILE__) . 'lib/';

    spl_autoload_register(function ($class) use ($base) {
        $prefix = 'PhpOffice\\PhpSpreadsheet\\';
        $len    = strlen($prefix);

        if (strncmp($prefix, $class, $len) !== 0) {
            return;
        }

        $relative_class = substr($class, $len);
        $file           = $base . 'PhpOffice/PhpSpreadsheet/' . str_replace('\\', '/', $relative_class) . '.php';

        if (file_exists($file)) {
            require_once $file;
        }
    });

    spl_autoload_register(function ($class) use ($base) {
        if ($class === 'Psr\\SimpleCache\\CacheInterface') {
            $file = $base . 'Psr/SimpleCache/CacheInterface.php';
            if (file_exists($file)) {
                require_once $file;
            }
        }
    });

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
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

/* ------------------------------------------------------------
 * 3️⃣ ADMIN MENU & NOTICES
 * ------------------------------------------------------------ */
function nf_cpt_xlsx_inline_required_capability() {
    $default_cap = 'manage_options';

    if (function_exists('apply_filters')) {
        $default_cap = apply_filters('ninja_forms_admin_menu_capability', $default_cap);
    }

    return apply_filters('nf_cpt_xlsx_inline_capability', $default_cap);
}

add_action('admin_menu', function () {
    $capability = nf_cpt_xlsx_inline_required_capability();

    add_submenu_page(
        'ninja-forms',
        __('Export to XLSX', NF_CPT_XLSX_INLINE_TEXT_DOMAIN),
        __('Export to XLSX', NF_CPT_XLSX_INLINE_TEXT_DOMAIN),
        $capability,
        NF_CPT_XLSX_INLINE_SLUG,
        'nf_cpt_xlsx_inline_admin_page'
    );
});

add_action('admin_notices', function () {
    if (!isset($_GET['page']) || $_GET['page'] !== NF_CPT_XLSX_INLINE_SLUG) {
        return;
    }

    if (!empty($_GET['nf_xlsx_notice']) && $_GET['nf_xlsx_notice'] === 'success') {
        $fileParam = isset($_GET['nf_xlsx_file']) ? wp_unslash($_GET['nf_xlsx_file']) : '';
        $file = $fileParam ? sanitize_file_name(rawurldecode($fileParam)) : '';
        if ($file) {
            $uploads = wp_upload_dir();
            $url     = trailingslashit($uploads['url']) . $file;

            printf(
                '<div class="notice notice-success"><p>%s <a href="%s">%s</a></p></div>',
                esc_html__('Excel export completed successfully.', NF_CPT_XLSX_INLINE_TEXT_DOMAIN),
                esc_url($url),
                esc_html__('Download file', NF_CPT_XLSX_INLINE_TEXT_DOMAIN)
            );
        } else {
            printf(
                '<div class="notice notice-success"><p>%s</p></div>',
                esc_html__('Excel export completed successfully.', NF_CPT_XLSX_INLINE_TEXT_DOMAIN)
            );
        }
    }

    if (!empty($_GET['nf_xlsx_notice']) && $_GET['nf_xlsx_notice'] === 'error') {
        $messageParam = isset($_GET['nf_xlsx_message']) ? wp_unslash($_GET['nf_xlsx_message']) : '';
        $decodedMessage = $messageParam ? rawurldecode($messageParam) : '';
        $message = $decodedMessage ? wp_strip_all_tags($decodedMessage) : __('Unknown error.', NF_CPT_XLSX_INLINE_TEXT_DOMAIN);
        printf(
            '<div class="notice notice-error"><p>%s</p></div>',
            esc_html($message)
        );
    }
});

/* ------------------------------------------------------------
 * 4️⃣ ADMIN PAGE
 * ------------------------------------------------------------ */
function nf_cpt_xlsx_inline_admin_page() {
    if (!current_user_can(nf_cpt_xlsx_inline_required_capability())) {
        wp_die(__('You are not allowed to access this page.', NF_CPT_XLSX_INLINE_TEXT_DOMAIN));
    }

    $forms = nf_cpt_xlsx_inline_get_forms();
    $current_form = isset($_GET['form_id']) ? absint($_GET['form_id']) : 0;

    ?>
    <div class="wrap">
        <h1><?php esc_html_e('Ninja Forms → Export to XLSX', NF_CPT_XLSX_INLINE_TEXT_DOMAIN); ?></h1>
        <p><?php esc_html_e('Generate an Excel workbook of Ninja Forms submissions. Uploaded images are embedded directly in the sheet.', NF_CPT_XLSX_INLINE_TEXT_DOMAIN); ?></p>

        <?php if (empty($forms)) : ?>
            <div class="notice notice-warning"><p><?php esc_html_e('No Ninja Forms found. Create a form to enable exports.', NF_CPT_XLSX_INLINE_TEXT_DOMAIN); ?></p></div>
        <?php else : ?>
            <form method="post" action="<?php echo esc_url(admin_url('admin-post.php')); ?>">
                <?php wp_nonce_field('nf_cpt_xlsx_inline_export', '_nf_cpt_nonce'); ?>
                <input type="hidden" name="action" value="nf_cpt_export_to_xlsx" />
                <table class="form-table" role="presentation">
                    <tbody>
                        <tr>
                            <th scope="row"><label for="nf-cpt-xlsx-form-id"><?php esc_html_e('Select form', NF_CPT_XLSX_INLINE_TEXT_DOMAIN); ?></label></th>
                            <td>
                                <select id="nf-cpt-xlsx-form-id" name="form_id">
                                    <?php foreach ($forms as $form_id => $form_title) : ?>
                                        <option value="<?php echo esc_attr($form_id); ?>" <?php selected($current_form, $form_id); ?>><?php echo esc_html($form_title); ?></option>
                                    <?php endforeach; ?>
                                </select>
                                <p class="description"><?php esc_html_e('Exports include all submission fields and uploaded images.', NF_CPT_XLSX_INLINE_TEXT_DOMAIN); ?></p>
                            </td>
                        </tr>
                    </tbody>
                </table>
                <?php submit_button(__('Export submissions', NF_CPT_XLSX_INLINE_TEXT_DOMAIN)); ?>
            </form>
        <?php endif; ?>
    </div>
    <?php
}

/* ------------------------------------------------------------
 * 5️⃣ EXPORT HANDLER
 * ------------------------------------------------------------ */
add_action('admin_post_nf_cpt_export_to_xlsx', 'nf_cpt_xlsx_inline_export');

function nf_cpt_xlsx_inline_export() {
    if (!current_user_can(nf_cpt_xlsx_inline_required_capability())) {
        wp_die(__('You are not allowed to export data.', NF_CPT_XLSX_INLINE_TEXT_DOMAIN));
    }

    check_admin_referer('nf_cpt_xlsx_inline_export', '_nf_cpt_nonce');

    $form_id = isset($_POST['form_id']) ? absint($_POST['form_id']) : 0;

    if (!$form_id) {
        nf_cpt_xlsx_inline_redirect_error(__('Invalid form selection.', NF_CPT_XLSX_INLINE_TEXT_DOMAIN));
    }

    try {
        $form = nf_cpt_xlsx_inline_get_form($form_id);
        if (!$form) {
            nf_cpt_xlsx_inline_redirect_error(__('Selected form does not exist.', NF_CPT_XLSX_INLINE_TEXT_DOMAIN));
        }

        $spreadsheet = nf_cpt_xlsx_inline_build_spreadsheet($form_id, $form['title']);

        $filename = nf_cpt_xlsx_inline_generate_filename($form_id);
        $saved    = nf_cpt_xlsx_inline_save_spreadsheet($spreadsheet, $filename);

        $redirect_url = nf_cpt_xlsx_inline_build_success_redirect($form_id, $saved['filename']);

        wp_safe_redirect($redirect_url);
        exit;
    } catch (Throwable $exception) {
        error_log('NF XLSX Export Error: ' . $exception->getMessage());
        nf_cpt_xlsx_inline_redirect_error($exception->getMessage());
    }
}

function nf_cpt_xlsx_inline_redirect_error($message) {
    $redirect_url = add_query_arg(
        [
            'page'             => NF_CPT_XLSX_INLINE_SLUG,
            'nf_xlsx_notice'   => 'error',
            'nf_xlsx_message'  => rawurlencode($message),
        ],
        admin_url('admin.php')
    );

    wp_safe_redirect($redirect_url);
    exit;
}

function nf_cpt_xlsx_inline_generate_filename($form_id) {
    return sprintf('nf-export-%d-%s.xlsx', (int) $form_id, gmdate('Y-m-d-H-i'));
}

function nf_cpt_xlsx_inline_save_spreadsheet(Spreadsheet $spreadsheet, $filename) {
    $uploads = wp_upload_dir();

    if (!empty($uploads['error'])) {
        throw new RuntimeException($uploads['error']);
    }

    if (!wp_mkdir_p($uploads['path'])) {
        throw new RuntimeException(__('Unable to create upload directory.', NF_CPT_XLSX_INLINE_TEXT_DOMAIN));
    }

    $filepath = trailingslashit($uploads['path']) . $filename;

    $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
    $writer->setPreCalculateFormulas(false);
    $writer->save($filepath);
    $spreadsheet->disconnectWorksheets();

    return [
        'path'     => $filepath,
        'filename' => basename($filename),
        'url'      => trailingslashit($uploads['url']) . basename($filename),
    ];
}

function nf_cpt_xlsx_inline_build_success_redirect($form_id, $filename) {
    return add_query_arg(
        [
            'page'           => NF_CPT_XLSX_INLINE_SLUG,
            'form_id'        => (int) $form_id,
            'nf_xlsx_notice' => 'success',
            'nf_xlsx_file'   => rawurlencode($filename),
        ],
        admin_url('admin.php')
    );
}

/* ------------------------------------------------------------
 * 6️⃣ DATA HELPERS
 * ------------------------------------------------------------ */
function nf_cpt_xlsx_inline_get_forms() {
    global $wpdb;

    $table = $wpdb->prefix . 'nf3_forms';
    $rows  = $wpdb->get_results("SELECT id, title FROM {$table} ORDER BY title ASC", ARRAY_A);

    $forms = [];

    if ($rows) {
        foreach ($rows as $row) {
            $forms[(int) $row['id']] = $row['title'];
        }
    }

    return $forms;
}

function nf_cpt_xlsx_inline_get_form($form_id) {
    global $wpdb;

    $table = $wpdb->prefix . 'nf3_forms';
    $row   = $wpdb->get_row($wpdb->prepare("SELECT id, title FROM {$table} WHERE id = %d", $form_id), ARRAY_A);

    if (!$row) {
        return null;
    }

    return [
        'id'    => (int) $row['id'],
        'title' => $row['title'],
    ];
}

function nf_cpt_xlsx_inline_get_form_fields($form_id) {
    global $wpdb;

    $table = $wpdb->prefix . 'nf3_fields';
    $rows  = $wpdb->get_results(
        $wpdb->prepare(
            "SELECT id, `key`, label, type FROM {$table} WHERE parent_id = %d ORDER BY `order` ASC, id ASC",
            $form_id
        ),
        ARRAY_A
    );

    $fields = [];

    if ($rows) {
        foreach ($rows as $row) {
            $fields[] = [
                'id'    => (int) $row['id'],
                'key'   => $row['key'],
                'label' => $row['label'] !== '' ? $row['label'] : $row['key'],
                'type'  => $row['type'],
            ];
        }
    }

    return $fields;
}

function nf_cpt_xlsx_inline_get_submissions($form_id) {
    global $wpdb;

    $table = $wpdb->prefix . 'nf3_subs';
    $rows  = $wpdb->get_results(
        $wpdb->prepare("SELECT id, form_id, sub_date, fields FROM {$table} WHERE form_id = %d ORDER BY sub_date ASC", $form_id),
        ARRAY_A
    );

    return $rows ? $rows : [];
}

function nf_cpt_xlsx_inline_build_spreadsheet($form_id, $form_title) {
    $fields      = nf_cpt_xlsx_inline_get_form_fields($form_id);
    $submissions = nf_cpt_xlsx_inline_get_submissions($form_id);

    $spreadsheet = new Spreadsheet();
    $spreadsheet->getProperties()
        ->setCreator(get_bloginfo('name'))
        ->setTitle(sprintf(__('Ninja Forms Export – %s', NF_CPT_XLSX_INLINE_TEXT_DOMAIN), $form_title))
        ->setDescription(__('Generated with NF CPT → XLSX Inline Export.', NF_CPT_XLSX_INLINE_TEXT_DOMAIN));

    $sheet = $spreadsheet->getActiveSheet();
    $sheet->setTitle(__('Submissions', NF_CPT_XLSX_INLINE_TEXT_DOMAIN));

    $headers = [
        'submission_id' => __('Submission ID', NF_CPT_XLSX_INLINE_TEXT_DOMAIN),
        'submitted_at'  => __('Submitted At', NF_CPT_XLSX_INLINE_TEXT_DOMAIN),
    ];

    foreach ($fields as $field) {
        $headers['field_' . $field['id']] = $field['label'];
    }

    $columnIndex = 1;
    foreach ($headers as $header) {
        $coordinate = Coordinate::stringFromColumnIndex($columnIndex) . '1';
        $sheet->setCellValue($coordinate, $header);
        $sheet->getStyle($coordinate)->getFont()->setBold(true);
        $sheet->getStyle($coordinate)->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
        $sheet->getStyle($coordinate)->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
        ++$columnIndex;
    }

    if (empty($submissions)) {
        $sheet->setCellValue('A2', __('No submissions available for this form.', NF_CPT_XLSX_INLINE_TEXT_DOMAIN));
        $sheet->mergeCells('A2:' . Coordinate::stringFromColumnIndex(count($headers)) . '2');
        $sheet->getStyle('A2')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
        $sheet->getRowDimension(2)->setRowHeight(24);
    } else {
        $rowIndex = 2;

        foreach ($submissions as $submission) {
            $normalized = nf_cpt_xlsx_inline_normalize_submission_fields($submission['fields']);
            $sheet->setCellValue('A' . $rowIndex, $submission['id']);
            $sheet->setCellValue('B' . $rowIndex, nf_cpt_xlsx_inline_format_date($submission['sub_date']));

            $columnIndex = 3;
            foreach ($fields as $field) {
                $coordinate = Coordinate::stringFromColumnIndex($columnIndex) . $rowIndex;
                $valuePayload = nf_cpt_xlsx_inline_extract_value($field, $normalized);
                $sheet->setCellValue($coordinate, $valuePayload['text']);
                $sheet->getStyle($coordinate)->getAlignment()->setWrapText(true);
                $sheet->getStyle($coordinate)->getAlignment()->setVertical(Alignment::VERTICAL_TOP);

                if (!empty($valuePayload['images'])) {
                    nf_cpt_xlsx_inline_embed_images($sheet, $coordinate, $valuePayload['images'], $rowIndex);
                }

                ++$columnIndex;
            }

            ++$rowIndex;
        }
    }

    // Styling adjustments.
    $sheet->freezePane('A2');
    $sheet->getColumnDimension('A')->setWidth(14);
    $sheet->getColumnDimension('B')->setWidth(22);

    $totalColumns = count($headers);
    for ($index = 3; $index <= $totalColumns; $index++) {
        $columnLetter = Coordinate::stringFromColumnIndex($index);
        $sheet->getColumnDimension($columnLetter)->setWidth(30);
    }

    $lastColumn = Coordinate::stringFromColumnIndex($totalColumns);
    $lastRow    = max(2, count($submissions) + 1);
    $range      = 'A1:' . $lastColumn . $lastRow;

    $sheet->getStyle($range)->getAlignment()->setWrapText(true);

    return $spreadsheet;
}

function nf_cpt_xlsx_inline_normalize_submission_fields($raw_fields) {
    $fields = maybe_unserialize($raw_fields);

    if (is_string($fields)) {
        $decoded = json_decode($fields, true);
        if (json_last_error() === JSON_ERROR_NONE) {
            $fields = $decoded;
        }
    }

    if (!is_array($fields)) {
        return [];
    }

    $normalized = [];

    foreach ($fields as $field) {
        if (!is_array($field)) {
            continue;
        }

        $valuePayload = nf_cpt_xlsx_inline_prepare_value_payload(isset($field['value']) ? $field['value'] : '');
        $keys = [];

        if (isset($field['id'])) {
            $keys[] = (string) $field['id'];
        }
        if (!empty($field['key'])) {
            $keys[] = (string) $field['key'];
        }

        if (empty($keys)) {
            continue;
        }

        foreach ($keys as $key) {
            $normalized[$key] = $valuePayload;
        }
    }

    return $normalized;
}

function nf_cpt_xlsx_inline_prepare_value_payload($value) {
    $payload = [
        'text'   => '',
        'images' => [],
    ];

    if (is_array($value)) {
        // File uploads or multiple selections.
        if (nf_cpt_xlsx_inline_is_upload_payload($value)) {
            $filePath = nf_cpt_xlsx_inline_locate_file_path($value);
            $label    = nf_cpt_xlsx_inline_guess_file_label($value, $filePath);

            if ($filePath && nf_cpt_xlsx_inline_is_image_path($filePath)) {
                $payload['images'][] = $filePath;
            }

            if ($label !== '') {
                $payload['text'] = $label;
            } elseif (!empty($value['url']) && is_string($value['url'])) {
                $payload['text'] = $value['url'];
            }
        } else {
            $texts  = [];
            $images = [];

            foreach ($value as $item) {
                $itemPayload = nf_cpt_xlsx_inline_prepare_value_payload($item);
                if ($itemPayload['text'] !== '') {
                    $texts[] = $itemPayload['text'];
                }
                if (!empty($itemPayload['images'])) {
                    $images = array_merge($images, $itemPayload['images']);
                }
            }

            if ($texts) {
                $payload['text'] = implode(', ', $texts);
            }
            if ($images) {
                $payload['images'] = array_values(array_unique($images));
            }
        }
    } elseif (is_scalar($value)) {
        $payload['text'] = (string) $value;
    }

    return $payload;
}

function nf_cpt_xlsx_inline_is_upload_payload(array $value) {
    $upload_keys = ['tmp_name', 'file_path', 'file_name', 'url', 'path', 'saved_name'];
    foreach ($upload_keys as $key) {
        if (array_key_exists($key, $value)) {
            return true;
        }
    }

    // Ninja Forms file upload stores entries as arrays containing arrays with these keys.
    if (isset($value[0]) && is_array($value[0])) {
        foreach ($value[0] as $key => $unused) {
            if (in_array($key, $upload_keys, true)) {
                return true;
            }
        }
    }

    return false;
}

function nf_cpt_xlsx_inline_locate_file_path($value) {
    $candidates = [];

    if (is_array($value)) {
        foreach (['file_path', 'path', 'tmp_name', 'value'] as $key) {
            if (!empty($value[$key]) && is_string($value[$key])) {
                $candidates[] = $value[$key];
            }
        }

        if (!empty($value['url']) && is_string($value['url'])) {
            $candidates[] = $value['url'];
        }

        if (isset($value[0]) && is_array($value[0])) {
            $inner = nf_cpt_xlsx_inline_locate_file_path($value[0]);
            if ($inner) {
                $candidates[] = $inner;
            }
        }
    } elseif (is_string($value)) {
        $candidates[] = $value;
    }

    $uploads = wp_upload_dir();
    $baseDir = trailingslashit($uploads['basedir']);
    $baseUrl = trailingslashit($uploads['baseurl']);

    foreach ($candidates as $candidate) {
        $candidate = trim($candidate);
        if (!$candidate) {
            continue;
        }

        if (file_exists($candidate)) {
            return realpath($candidate);
        }

        if (strpos($candidate, $baseUrl) === 0) {
            $maybe = $baseDir . ltrim(substr($candidate, strlen($baseUrl)), '/');
            if (file_exists($maybe)) {
                return realpath($maybe);
            }
        }

        // Allow relative paths from uploads directory.
        $maybe = $baseDir . ltrim($candidate, '/');
        if (file_exists($maybe)) {
            return realpath($maybe);
        }
    }

    return '';
}

function nf_cpt_xlsx_inline_guess_file_label($value, $filePath) {
    if (is_array($value)) {
        foreach (['file_name', 'saved_name', 'filename', 'name'] as $key) {
            if (!empty($value[$key])) {
                return (string) $value[$key];
            }
        }

        if (!empty($value['url'])) {
            return basename(parse_url($value['url'], PHP_URL_PATH));
        }

        if (!empty($value['value']) && is_string($value['value'])) {
            return basename($value['value']);
        }

        if (isset($value[0]) && is_array($value[0])) {
            return nf_cpt_xlsx_inline_guess_file_label($value[0], $filePath);
        }
    }

    if ($filePath) {
        return basename($filePath);
    }

    return '';
}

function nf_cpt_xlsx_inline_extract_value($field, $normalized) {
    $payload = ['text' => '', 'images' => []];

    $candidates = [];
    $candidates[] = (string) $field['id'];

    if (!empty($field['key'])) {
        $candidates[] = (string) $field['key'];
    }

    foreach ($candidates as $candidate) {
        if (isset($normalized[$candidate])) {
            $payload = $normalized[$candidate];
            break;
        }
    }

    return $payload;
}

function nf_cpt_xlsx_inline_embed_images($sheet, $coordinate, array $images, $rowIndex) {
    if (empty($images)) {
        return;
    }

    $maxImages = 5; // Safety guard.
    $offset = 0;
    $imageCount = min(count($images), $maxImages);
    $rowHeight = max(80, $imageCount * 80);
    $sheet->getRowDimension($rowIndex)->setRowHeight($rowHeight);

    foreach ($images as $index => $path) {
        if ($index >= $maxImages) {
            break;
        }

        if (!file_exists($path) || !nf_cpt_xlsx_inline_is_image_path($path)) {
            continue;
        }

        $drawing = new Drawing();
        $drawing->setName(basename($path));
        $drawing->setDescription(basename($path));
        $drawing->setPath($path);
        $drawing->setCoordinates($coordinate);
        $drawing->setOffsetY($offset);
        $drawing->setWorksheet($sheet);
        $offset += 80;
    }
}

function nf_cpt_xlsx_inline_format_date($date) {
    if (!$date) {
        return '';
    }

    $timestamp = strtotime($date);
    if (!$timestamp) {
        return $date;
    }

    return wp_date(get_option('date_format') . ' ' . get_option('time_format'), $timestamp);
}

function nf_cpt_xlsx_inline_is_image_path($path) {
    if (!is_string($path) || $path === '') {
        return false;
    }

    $filetype = wp_check_filetype($path);
    if (!empty($filetype['ext'])) {
        $allowed = ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'webp'];
        if (in_array(strtolower($filetype['ext']), $allowed, true)) {
            return true;
        }
    }

    $imageInfo = @getimagesize($path);
    return $imageInfo !== false;
}
