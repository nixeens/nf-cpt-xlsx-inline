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

// -----------------------------------------------------------------------------
// Autoload bundled libraries (PhpSpreadsheet + lightweight dependencies).
// -----------------------------------------------------------------------------
function nf_xlsx_register_local_autoloaders() {
    static $registered = false;

    if ($registered) {
        return;
    }

    $base = plugin_dir_path(__FILE__) . 'lib/';

    spl_autoload_register(static function ($class) use ($base) {
        $prefix = 'PhpOffice\\PhpSpreadsheet\\';
        $length = strlen($prefix);

        if (strncmp($class, $prefix, $length) !== 0) {
            return;
        }

        $relative = substr($class, $length);
        $file     = $base . 'PhpOffice/PhpSpreadsheet/' . str_replace('\\', '/', $relative) . '.php';

        if (file_exists($file)) {
            require_once $file;
        }
    });

    spl_autoload_register(static function ($class) use ($base) {
        if ($class === 'Psr\\SimpleCache\\CacheInterface') {
            $file = $base . 'Psr/SimpleCache/CacheInterface.php';
            if (file_exists($file)) {
                require_once $file;
            }
        }
    });

    spl_autoload_register(static function ($class) use ($base) {
        if ($class === 'Composer\\Pcre\\Preg') {
            $file = $base . 'Composer/Pcre/Preg.php';
            if (file_exists($file)) {
                require_once $file;
            }
        }
    });

    $registered = true;
}
add_action('plugins_loaded', 'nf_xlsx_register_local_autoloaders', 1);

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

// -----------------------------------------------------------------------------
// Admin UI registration.
// -----------------------------------------------------------------------------
add_action('admin_menu', static function () {
    add_submenu_page(
        'ninja-forms',
        __('Export to XLSX', 'nf-cpt-xlsx-inline'),
        __('Export to XLSX', 'nf-cpt-xlsx-inline'),
        'manage_options',
        'nf-cpt-xlsx-inline',
        'nf_xlsx_render_admin_page'
    );
});

add_action('admin_post_nf_xlsx_export', 'nf_xlsx_handle_export');

add_action('admin_notices', static function () {
    if (!isset($_GET['page']) || $_GET['page'] !== 'nf-cpt-xlsx-inline') {
        return;
    }

    if (!empty($_GET['nf_xlsx_notice']) && $_GET['nf_xlsx_notice'] === 'success') {
        $uploads = wp_upload_dir();
        $file    = isset($_GET['nf_xlsx_file']) ? sanitize_file_name(wp_unslash(rawurldecode($_GET['nf_xlsx_file']))) : '';

        if ($file) {
            $url = trailingslashit($uploads['url']) . $file;
            printf(
                '<div class="notice notice-success"><p>%s <a href="%s">%s</a></p></div>',
                esc_html__('Excel export completed successfully.', 'nf-cpt-xlsx-inline'),
                esc_url($url),
                esc_html__('Download file', 'nf-cpt-xlsx-inline')
            );
        } else {
            printf(
                '<div class="notice notice-success"><p>%s</p></div>',
                esc_html__('Excel export completed successfully.', 'nf-cpt-xlsx-inline')
            );
        }
    }

    if (!empty($_GET['nf_xlsx_notice']) && $_GET['nf_xlsx_notice'] === 'error') {
        $message = isset($_GET['nf_xlsx_message']) ? wp_strip_all_tags(wp_unslash(rawurldecode($_GET['nf_xlsx_message']))) : '';
        $message = $message ?: __('Unknown error.', 'nf-cpt-xlsx-inline');

        printf(
            '<div class="notice notice-error"><p>%s</p></div>',
            esc_html($message)
        );
    }
});

// -----------------------------------------------------------------------------
// Admin page markup.
// -----------------------------------------------------------------------------
function nf_xlsx_render_admin_page() {
    if (!current_user_can('manage_options')) {
        wp_die(__('You are not allowed to access this page.', 'nf-cpt-xlsx-inline'));
    }

    $forms       = nf_xlsx_get_forms();
    $selected_id = isset($_GET['form_id']) ? absint($_GET['form_id']) : 0;
    ?>
    <div class="wrap">
        <h1><?php esc_html_e('Ninja Forms → Export to XLSX', 'nf-cpt-xlsx-inline'); ?></h1>
        <p><?php esc_html_e('Create a UTF-8 Excel workbook of Ninja Forms submissions. Uploaded images are embedded alongside their field values.', 'nf-cpt-xlsx-inline'); ?></p>

        <?php if (empty($forms)) : ?>
            <div class="notice notice-warning"><p><?php esc_html_e('No Ninja Forms available. Create a form to enable exports.', 'nf-cpt-xlsx-inline'); ?></p></div>
        <?php else : ?>
            <form method="post" action="<?php echo esc_url(admin_url('admin-post.php')); ?>">
                <?php wp_nonce_field('nf_xlsx_export', '_nf_xlsx_nonce'); ?>
                <input type="hidden" name="action" value="nf_xlsx_export">
                <table class="form-table" role="presentation">
                    <tbody>
                        <tr>
                            <th scope="row"><label for="nf-xlsx-form-id"><?php esc_html_e('Select form', 'nf-cpt-xlsx-inline'); ?></label></th>
                            <td>
                                <select name="form_id" id="nf-xlsx-form-id">
                                    <?php foreach ($forms as $id => $title) : ?>
                                        <option value="<?php echo esc_attr($id); ?>" <?php selected($selected_id, $id); ?>><?php echo esc_html($title); ?></option>
                                    <?php endforeach; ?>
                                </select>
                                <p class="description"><?php esc_html_e('Exports include all field labels, values, and uploaded images.', 'nf-cpt-xlsx-inline'); ?></p>
                            </td>
                        </tr>
                    </tbody>
                </table>
                <?php submit_button(__('Export submissions', 'nf-cpt-xlsx-inline')); ?>
            </form>
        <?php endif; ?>
    </div>
    <?php
}

// -----------------------------------------------------------------------------
// Export handler.
// -----------------------------------------------------------------------------
function nf_xlsx_handle_export() {
    if (!current_user_can('manage_options')) {
        wp_die(__('You are not allowed to export data.', 'nf-cpt-xlsx-inline'));
    }

    check_admin_referer('nf_xlsx_export', '_nf_xlsx_nonce');

    $form_id = isset($_POST['form_id']) ? absint($_POST['form_id']) : 0;
    if (!$form_id) {
        nf_xlsx_redirect_error(__('Invalid form selection.', 'nf-cpt-xlsx-inline'));
    }

    try {
        $form = nf_xlsx_get_form($form_id);
        if (!$form) {
            nf_xlsx_redirect_error(__('Selected form does not exist.', 'nf-cpt-xlsx-inline'));
        }

        $fields      = nf_xlsx_get_form_fields($form_id);
        $submissions = nf_xlsx_get_submissions($form_id);

        $spreadsheet = nf_xlsx_build_workbook($form, $fields, $submissions);

        $filename = sprintf('nf-export-%d-%s.xlsx', (int) $form_id, gmdate('Y-m-d-H-i'));
        $uploads  = wp_upload_dir();

        if (!empty($uploads['error'])) {
            throw new RuntimeException($uploads['error']);
        }

        if (!wp_mkdir_p($uploads['path'])) {
            throw new RuntimeException(__('Unable to create upload directory.', 'nf-cpt-xlsx-inline'));
        }

        $filepath = trailingslashit($uploads['path']) . $filename;

        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->setPreCalculateFormulas(false);
        $writer->save($filepath);
        $spreadsheet->disconnectWorksheets();
        unset($spreadsheet);

        $redirect = add_query_arg(
            [
                'page'           => 'nf-cpt-xlsx-inline',
                'form_id'        => $form_id,
                'nf_xlsx_notice' => 'success',
                'nf_xlsx_file'   => rawurlencode($filename),
            ],
            admin_url('admin.php')
        );

        wp_safe_redirect($redirect);
        exit;
    } catch (Throwable $exception) {
        error_log('NF XLSX Export Error: ' . $exception->getMessage());
        nf_xlsx_redirect_error($exception->getMessage());
    }
}

function nf_xlsx_redirect_error($message) {
    $redirect = add_query_arg(
        [
            'page'             => 'nf-cpt-xlsx-inline',
            'nf_xlsx_notice'   => 'error',
            'nf_xlsx_message'  => rawurlencode($message),
        ],
        admin_url('admin.php')
    );

    wp_safe_redirect($redirect);
    exit;
}

// -----------------------------------------------------------------------------
// Data access helpers.
// -----------------------------------------------------------------------------
function nf_xlsx_get_forms() {
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

function nf_xlsx_get_form($form_id) {
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

function nf_xlsx_get_form_fields($form_id) {
    global $wpdb;

    $table = $wpdb->prefix . 'nf3_fields';
    $rows  = $wpdb->get_results(
        $wpdb->prepare("SELECT id, `key`, label, type FROM {$table} WHERE parent_id = %d ORDER BY id ASC", $form_id),
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

function nf_xlsx_get_submissions($form_id) {
    global $wpdb;

    $table = $wpdb->prefix . 'nf3_subs';
    $rows  = $wpdb->get_results(
        $wpdb->prepare("SELECT id, form_id, sub_date, fields FROM {$table} WHERE form_id = %d ORDER BY sub_date ASC", $form_id),
        ARRAY_A
    );

    return $rows ?: [];
}

// -----------------------------------------------------------------------------
// Spreadsheet builder.
// -----------------------------------------------------------------------------
function nf_xlsx_build_workbook(array $form, array $fields, array $submissions) {
    $spreadsheet = new Spreadsheet();
    $spreadsheet->getProperties()
        ->setCreator(get_bloginfo('name'))
        ->setTitle(sprintf(__('Ninja Forms Export – %s', 'nf-cpt-xlsx-inline'), $form['title']))
        ->setDescription(__('Generated with NF CPT → XLSX Inline Export.', 'nf-cpt-xlsx-inline'));

    $sheet = $spreadsheet->getActiveSheet();
    $sheet->setTitle(__('Submissions', 'nf-cpt-xlsx-inline'));

    $headers = [
        'submission_id' => __('Submission ID', 'nf-cpt-xlsx-inline'),
        'submitted_at'  => __('Submitted At', 'nf-cpt-xlsx-inline'),
    ];

    foreach ($fields as $field) {
        $column_key = 'field_' . $field['id'];
        $label      = $field['label'] !== '' ? $field['label'] : $field['key'];

        if (isset($headers[$column_key])) {
            $label .= ' (' . $field['id'] . ')';
        }

        $headers[$column_key] = $label;
    }

    $column = 1;
    foreach ($headers as $header) {
        $coordinate = Coordinate::stringFromColumnIndex($column) . '1';
        $sheet->setCellValue($coordinate, $header);
        $sheet->getStyle($coordinate)->getFont()->setBold(true);
        $sheet->getStyle($coordinate)->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
        $sheet->getStyle($coordinate)->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
        ++$column;
    }

    if (empty($submissions)) {
        $sheet->setCellValue('A2', __('No submissions available for this form.', 'nf-cpt-xlsx-inline'));
        $lastColumn = Coordinate::stringFromColumnIndex(count($headers));
        $sheet->mergeCells('A2:' . $lastColumn . '2');
        $sheet->getStyle('A2')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
        $sheet->getRowDimension(2)->setRowHeight(24);
    } else {
        $rowIndex = 2;

        foreach ($submissions as $submission) {
            $normalized = nf_xlsx_normalize_submission_fields($submission['fields']);

            $sheet->setCellValue('A' . $rowIndex, $submission['id']);
            $sheet->setCellValue('B' . $rowIndex, nf_xlsx_format_date($submission['sub_date']));

            $columnIndex = 3;
            foreach ($fields as $field) {
                $coordinate   = Coordinate::stringFromColumnIndex($columnIndex) . $rowIndex;
                $valuePayload = nf_xlsx_resolve_field_value($field, $normalized);

                $sheet->setCellValue($coordinate, $valuePayload['text']);
                $sheet->getStyle($coordinate)->getAlignment()->setWrapText(true);
                $sheet->getStyle($coordinate)->getAlignment()->setVertical(Alignment::VERTICAL_TOP);

                if (!empty($valuePayload['images'])) {
                    nf_xlsx_embed_images($sheet, $coordinate, $valuePayload['images'], $rowIndex);
                }

                ++$columnIndex;
            }

            ++$rowIndex;
        }
    }

    $sheet->freezePane('A2');
    $sheet->getColumnDimension('A')->setWidth(14);
    $sheet->getColumnDimension('B')->setWidth(22);

    $totalColumns = count($headers);
    for ($index = 3; $index <= $totalColumns; $index++) {
        $sheet->getColumnDimension(Coordinate::stringFromColumnIndex($index))->setWidth(30);
    }

    $lastColumn = Coordinate::stringFromColumnIndex($totalColumns);
    $lastRow    = max(2, count($submissions) + 1);
    $sheet->getStyle('A1:' . $lastColumn . $lastRow)->getAlignment()->setWrapText(true);

    return $spreadsheet;
}

// -----------------------------------------------------------------------------
// Submission parsing helpers.
// -----------------------------------------------------------------------------
function nf_xlsx_normalize_submission_fields($raw) {
    $fields = maybe_unserialize($raw);

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

        $valuePayload = nf_xlsx_prepare_value_payload($field['value'] ?? '');
        $keys         = [];

        if (isset($field['id'])) {
            $keys[] = (string) $field['id'];
        }
        if (!empty($field['key'])) {
            $keys[] = (string) $field['key'];
        }

        if (!$keys) {
            continue;
        }

        foreach ($keys as $key) {
            if (!isset($normalized[$key])) {
                $normalized[$key] = $valuePayload;
            }
        }
    }

    return $normalized;
}

function nf_xlsx_resolve_field_value(array $field, array $normalized) {
    $candidates = [];
    $candidates[] = (string) $field['id'];

    if (!empty($field['key'])) {
        $candidates[] = (string) $field['key'];
    }

    foreach ($candidates as $candidate) {
        if (isset($normalized[$candidate])) {
            return $normalized[$candidate];
        }
    }

    return ['text' => '', 'images' => []];
}

function nf_xlsx_prepare_value_payload($value) {
    $payload = [
        'text'   => '',
        'images' => [],
    ];

    if (is_array($value)) {
        if (nf_xlsx_is_upload_payload($value)) {
            $filePath         = nf_xlsx_locate_file_path($value);
            $payload['text']  = nf_xlsx_guess_file_label($value, $filePath);
            if ($filePath) {
                $payload['images'][] = $filePath;
            }
        } else {
            $texts  = [];
            $images = [];
            foreach ($value as $item) {
                $itemPayload = nf_xlsx_prepare_value_payload($item);
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

function nf_xlsx_is_upload_payload(array $value) {
    $uploadKeys = ['tmp_name', 'file_path', 'file_name', 'url', 'path', 'saved_name'];

    foreach ($uploadKeys as $key) {
        if (array_key_exists($key, $value)) {
            return true;
        }
    }

    if (isset($value[0]) && is_array($value[0])) {
        foreach ($value[0] as $key => $unused) {
            if (in_array($key, $uploadKeys, true)) {
                return true;
            }
        }
    }

    return false;
}

function nf_xlsx_locate_file_path($value) {
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
            $candidates[] = nf_xlsx_locate_file_path($value[0]);
        }
    }

    $uploads = wp_upload_dir();
    $baseDir = trailingslashit($uploads['basedir']);
    $baseUrl = trailingslashit($uploads['baseurl']);

    foreach ($candidates as $candidate) {
        if (!$candidate) {
            continue;
        }

        $candidate = trim((string) $candidate);
        if ($candidate === '') {
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

        $maybe = $baseDir . ltrim($candidate, '/');
        if (file_exists($maybe)) {
            return realpath($maybe);
        }
    }

    return '';
}

function nf_xlsx_guess_file_label($value, $filePath) {
    if (is_array($value)) {
        foreach (['file_name', 'saved_name', 'filename', 'name'] as $key) {
            if (!empty($value[$key])) {
                return (string) $value[$key];
            }
        }

        if (!empty($value['url'])) {
            $path = parse_url($value['url'], PHP_URL_PATH);
            if ($path) {
                return basename($path);
            }
        }

        if (!empty($value['value']) && is_string($value['value'])) {
            return basename($value['value']);
        }

        if (isset($value[0]) && is_array($value[0])) {
            return nf_xlsx_guess_file_label($value[0], $filePath);
        }
    }

    if ($filePath) {
        return basename($filePath);
    }

    return '';
}

function nf_xlsx_embed_images($sheet, $coordinate, array $images, $rowIndex) {
    if (empty($images)) {
        return;
    }

    $maxImages  = 5;
    $imageCount = min(count($images), $maxImages);
    $rowHeight  = max(80, $imageCount * 80);
    $sheet->getRowDimension($rowIndex)->setRowHeight($rowHeight);

    $offsetY = 0;
    foreach ($images as $index => $path) {
        if ($index >= $maxImages) {
            break;
        }

        if (!file_exists($path)) {
            continue;
        }

        $drawing = new Drawing();
        $drawing->setName(basename($path));
        $drawing->setDescription(basename($path));
        $drawing->setPath($path);
        $drawing->setCoordinates($coordinate);
        $drawing->setOffsetY($offsetY);
        $drawing->setWorksheet($sheet);

        $offsetY += 80;
    }
}

function nf_xlsx_format_date($date) {
    if (!$date) {
        return '';
    }

    $timestamp = strtotime($date);
    if (!$timestamp) {
        return $date;
    }

    return wp_date(get_option('date_format') . ' ' . get_option('time_format'), $timestamp);
}
