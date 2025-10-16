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

    spl_autoload_register(static function ($class) use ($base) {
        $prefix = 'ZipStream\\';
        $length = strlen($prefix);

        if (strncmp($class, $prefix, $length) !== 0) {
            return;
        }

        $relative = substr($class, $length);
        $file     = $base . 'ZipStream/' . str_replace('\\', '/', $relative) . '.php';

        if (file_exists($file)) {
            require_once $file;
        }
    });

    $registered = true;
}
add_action('plugins_loaded', 'nf_xlsx_register_local_autoloaders', 1);

require_once __DIR__ . '/class-nf-xlsx-stream-exporter.php';

register_activation_hook(__FILE__, 'nf_xlsx_generate_activation_sample');

// -----------------------------------------------------------------------------
// Admin UI registration.
// -----------------------------------------------------------------------------
add_action('admin_menu', static function () {
    add_menu_page(
        __('NF Submissions Export', 'nf-cpt-xlsx-inline'),
        __('NF Submissions Export', 'nf-cpt-xlsx-inline'),
        'manage_options',
        'nf-cpt-xlsx-inline',
        'nf_xlsx_render_admin_page',
        'dashicons-media-spreadsheet',
        58
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

    if (!$selected_id && !empty($forms)) {
        $firstKey    = array_key_first($forms);
        $selected_id = $firstKey ? (int) $firstKey : 0;
    }

    $fields = [];
    if ($selected_id) {
        $fields = nf_xlsx_get_form_fields($selected_id);
    }

    $availableColumns = $selected_id ? nf_xlsx_prepare_columns($fields) : [];
    $availableIds     = array_map(static function ($column) {
        return isset($column['id']) ? (string) $column['id'] : '';
    }, $availableColumns);
    $availableIds     = array_values(array_filter($availableIds, static function ($value) {
        return $value !== '';
    }));

    $rawSelectedColumns = isset($_GET['columns']) ? (array) wp_unslash($_GET['columns']) : [];
    $rawSelectedColumns = array_map('sanitize_text_field', $rawSelectedColumns);
    $rawSelectedColumns = array_values(array_unique(array_filter($rawSelectedColumns, static function ($value) {
        return $value !== '';
    })));

    $selectedColumns = $rawSelectedColumns;
    if ($availableColumns) {
        $selectedColumns = nf_xlsx_normalize_column_selection($selectedColumns, $availableColumns);

        if (!$selectedColumns) {
            $selectedColumns = $availableIds;
        }
    } else {
        $selectedColumns = [];
    }

    $previewColumns     = $availableColumns ? nf_xlsx_prepare_columns($fields, $selectedColumns) : [];
    $previewSubmissions = $selected_id ? nf_xlsx_get_submissions($selected_id, 5) : [];
    ?>
    <div class="wrap">
        <h1><?php esc_html_e('NF Submissions Export', 'nf-cpt-xlsx-inline'); ?></h1>
        <p><?php esc_html_e('Download an .xlsx file containing all submissions for the selected Ninja Form.', 'nf-cpt-xlsx-inline'); ?></p>

        <?php if (empty($forms)) : ?>
            <div class="notice notice-warning"><p><?php esc_html_e('No Ninja Forms available. Create a form to enable exports.', 'nf-cpt-xlsx-inline'); ?></p></div>
        <?php else : ?>
            <form method="get" action="<?php echo esc_url(admin_url('admin.php')); ?>" class="nf-xlsx-options-form">
                <input type="hidden" name="page" value="nf-cpt-xlsx-inline">
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
                                <p class="description"><?php esc_html_e('Choose a form and adjust which columns appear in the export. Update the preview after making changes.', 'nf-cpt-xlsx-inline'); ?></p>
                            </td>
                        </tr>
                        <?php if ($availableColumns) : ?>
                            <tr>
                                <th scope="row"><?php esc_html_e('Columns', 'nf-cpt-xlsx-inline'); ?></th>
                                <td>
                                    <fieldset>
                                        <legend class="screen-reader-text"><?php esc_html_e('Select columns to export', 'nf-cpt-xlsx-inline'); ?></legend>
                                        <?php foreach ($availableColumns as $column) :
                                            $columnId    = isset($column['id']) ? (string) $column['id'] : '';
                                            $isChecked   = in_array($columnId, $selectedColumns, true);
                                            $inputId     = 'nf-xlsx-column-' . sanitize_html_class($columnId);
                                            ?>
                                            <label for="<?php echo esc_attr($inputId); ?>" style="display: inline-block; margin-right: 12px;">
                                                <input type="checkbox" id="<?php echo esc_attr($inputId); ?>" name="columns[]" value="<?php echo esc_attr($columnId); ?>" <?php checked($isChecked); ?>>
                                                <?php echo esc_html($column['header']); ?>
                                            </label>
                                        <?php endforeach; ?>
                                    </fieldset>
                                </td>
                            </tr>
                        <?php endif; ?>
                    </tbody>
                </table>
                <?php submit_button(__('Update Preview', 'nf-cpt-xlsx-inline'), 'secondary', 'submit', false); ?>
            </form>

            <?php if ($selected_id) : ?>
                <h2><?php esc_html_e('Preview (first 5 submissions)', 'nf-cpt-xlsx-inline'); ?></h2>
                <?php if (!$availableColumns) : ?>
                    <p><?php esc_html_e('This form does not have any fields available for export.', 'nf-cpt-xlsx-inline'); ?></p>
                <?php else : ?>
                    <table class="widefat striped">
                        <thead>
                            <tr>
                                <?php foreach ($previewColumns as $column) : ?>
                                    <th scope="col"><?php echo esc_html($column['header']); ?></th>
                                <?php endforeach; ?>
                            </tr>
                        </thead>
                        <tbody>
                            <?php if ($previewSubmissions) : ?>
                                <?php foreach ($previewSubmissions as $submission) : ?>
                                    <tr>
                                        <?php foreach ($previewColumns as $column) : ?>
                                            <td>
                                                <?php
                                                if ($column['field'] === null) {
                                                    echo wp_kses_post(nl2br(esc_html(nf_xlsx_format_date($submission['sub_date']))));
                                                } else {
                                                    $payload = nf_xlsx_extract_submission_field_payload($submission, $column['field']);
                                                    echo wp_kses_post(nl2br(esc_html($payload['text'])));
                                                }
                                                ?>
                                            </td>
                                        <?php endforeach; ?>
                                    </tr>
                                <?php endforeach; ?>
                            <?php else : ?>
                                <tr>
                                    <td colspan="<?php echo esc_attr(max(1, count($previewColumns))); ?>"><?php esc_html_e('No submissions found for this form.', 'nf-cpt-xlsx-inline'); ?></td>
                                </tr>
                            <?php endif; ?>
                        </tbody>
                    </table>
                <?php endif; ?>

                <form method="post" action="<?php echo esc_url(admin_url('admin-post.php')); ?>" style="margin-top: 20px;">
                    <?php wp_nonce_field('nf_xlsx_export', '_nf_xlsx_nonce'); ?>
                    <input type="hidden" name="action" value="nf_xlsx_export">
                    <input type="hidden" name="form_id" value="<?php echo esc_attr($selected_id); ?>">
                    <?php foreach ($selectedColumns as $columnId) : ?>
                        <input type="hidden" name="columns[]" value="<?php echo esc_attr($columnId); ?>">
                    <?php endforeach; ?>
                    <?php
                    if (!$selectedColumns) {
                        echo '<p class="description">' . esc_html__('Select at least one column before exporting.', 'nf-cpt-xlsx-inline') . '</p>';
                    }
                    submit_button(__('Export to XLSX', 'nf-cpt-xlsx-inline'));
                    ?>
                </form>
            <?php endif; ?>
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

        $fields = nf_xlsx_get_form_fields($form_id);

        $availableColumns = nf_xlsx_prepare_columns($fields);
        if (!$availableColumns) {
            nf_xlsx_redirect_error(__('No columns available for export.', 'nf-cpt-xlsx-inline'));
        }

        $rawSelectedColumns = isset($_POST['columns']) ? (array) wp_unslash($_POST['columns']) : [];
        $rawSelectedColumns = array_map('sanitize_text_field', $rawSelectedColumns);
        $rawSelectedColumns = array_values(array_unique(array_filter($rawSelectedColumns, static function ($value) {
            return $value !== '';
        })));

        $selectedColumns = nf_xlsx_normalize_column_selection($rawSelectedColumns, $availableColumns);
        if (!$selectedColumns) {
            nf_xlsx_redirect_error(__('Please select at least one column to export.', 'nf-cpt-xlsx-inline'));
        }

        $submissions = nf_xlsx_get_submissions($form_id);

        $exporter = nf_xlsx_build_workbook($form, $fields, $submissions, $selectedColumns);

        $filename = sprintf('nf-export-%d-%s.xlsx', (int) $form_id, gmdate('Y-m-d-H-i'));
        $uploads  = wp_upload_dir();

        if (!empty($uploads['error'])) {
            throw new RuntimeException($uploads['error']);
        }

        if (!wp_mkdir_p($uploads['path'])) {
            throw new RuntimeException(__('Unable to create upload directory.', 'nf-cpt-xlsx-inline'));
        }

        $filepath = trailingslashit($uploads['path']) . $filename;

        $exporter->save($filepath);

        $redirectArgs = [
            'page'           => 'nf-cpt-xlsx-inline',
            'form_id'        => $form_id,
            'nf_xlsx_notice' => 'success',
            'nf_xlsx_file'   => rawurlencode($filename),
            'columns'        => $selectedColumns,
        ];

        $redirect = add_query_arg($redirectArgs, admin_url('admin.php'));

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
// Activation sample workbook (simple smoke test).
// -----------------------------------------------------------------------------
function nf_xlsx_seed_sample_asset(string $directory, string $urlBase, string $filename, string $base64): ?array {
    $binary = base64_decode($base64, true);
    if ($binary === false) {
        return null;
    }

    if (!wp_mkdir_p($directory)) {
        return null;
    }

    $unique = function_exists('wp_unique_filename')
        ? wp_unique_filename($directory, $filename)
        : (uniqid('nf-sample-', true) . '-' . $filename);

    $targetDirectory = trailingslashit($directory);
    $path            = $targetDirectory . $unique;

    if (file_put_contents($path, $binary) === false) {
        return null;
    }

    return [
        'name' => $unique,
        'path' => $path,
        'url'  => trailingslashit($urlBase) . $unique,
    ];
}

function nf_xlsx_sample_image_base64(): string {
    return 'iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAIAAAAlC+aJAAAATklEQVR42u3PQQkAAAgEsIt2/UtpBN/CYAWWaV+LgICAgICAgICAgICAgICAgICAgICAgICAgICAg'
        . 'ICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAg'
        . 'ICAgICAgICAgICAgICAgICAgICAgMBlAflF8VpbbnTDAAAAAElFTkSuQmCC';
}

function nf_xlsx_sample_pdf_base64(): string {
    return 'JVBERi0xLjQKMSAwIG9iago8PCAvVHlwZSAvQ2F0YWxvZyAvUGFnZXMgMiAwIFIgPj4KZW5kb2JqCjIgMCBvYmoKPDwgL1R5cGUgL1BhZ2VzIC9LaWRzIFszIDAgUl0g'
        . 'L0NvdW50IDEgPj4KZW5kb2JqCjMgMCBvYmoKPDwgL1R5cGUgL1BhZ2UgL1BhcmVudCAyIDAgUiAvTWVkaWFCb3ggWzAgMCA2MTIgNzkyXSAvQ29udGVudHMgNCAwIFIg'
        . 'L1Jlc291cmNlcyA8PCAvRm9udCA8PCAvRjEgNSAwIFIgPj4gPj4gPj4KZW5kb2JqCjQgMCBvYmoKPDwgL0xlbmd0aCA2MSA+PgpzdHJlYW0KQlQgL0YxIDI0IFRmIDcy'
        . 'IDcyMCBUZCAoTkYgWExTWCBJbmxpbmUgU2FtcGxlIFBERikgVGogRVQKZW5kc3RyZWFtCmVuZG9iago1IDAgb2JqCjw8IC9UeXBlIC9Gb250IC9TdWJ0eXBlIC9UeXBl'
        . 'MSAvQmFzZUZvbnQgL0hlbHZldGljYSA+PgplbmRvYmoKeHJlZgowIDYKMDAwMDAwMDAwMCA2NTUzNSBmIAowMDAwMDAwMDA5IDAwMDAwIG4gCjAwMDAwMDAwNTggMDAw'
        . 'MDAgbiAKMDAwMDAwMDExNSAwMDAwMCBuIAowMDAwMDAwMjQxIDAwMDAwIG4gCjAwMDAwMDAzNDcgMDAwMDAgbiAKdHJhaWxlcgo8PCAvU2l6ZSA2IC9Sb290IDEgMCBS'
        . 'ID4+CnN0YXJ0eHJlZgo0MTcKJSVFT0Y=';
}

function nf_xlsx_generate_activation_sample() {
    if (!function_exists('wp_upload_dir')) {
        return;
    }

    try {
        nf_xlsx_register_local_autoloaders();

        $form = [
            'id'    => 0,
            'title' => __('Activation Sample', 'nf-cpt-xlsx-inline'),
        ];

        $fields = [
            [
                'id'         => 9001,
                'key'        => 'sample_text',
                'label'      => __('Sample Text', 'nf-cpt-xlsx-inline'),
                'type'       => 'textbox',
                'identifier' => nf_xlsx_field_identifier(9001),
            ],
            [
                'id'         => 9002,
                'key'        => 'sample_image',
                'label'      => __('Sample Image Upload', 'nf-cpt-xlsx-inline'),
                'type'       => 'file',
                'identifier' => nf_xlsx_field_identifier(9002),
            ],
            [
                'id'         => 9003,
                'key'        => 'sample_pdf',
                'label'      => __('Sample PDF Upload', 'nf-cpt-xlsx-inline'),
                'type'       => 'file',
                'identifier' => nf_xlsx_field_identifier(9003),
            ],
            [
                'id'         => 9004,
                'key'        => 'rich_text',
                'label'      => __('Rich Text Content', 'nf-cpt-xlsx-inline'),
                'type'       => 'textarea',
                'identifier' => nf_xlsx_field_identifier(9004),
            ],
        ];

        $columns = nf_xlsx_prepare_columns($fields);
        if (!$columns) {
            return;
        }

        $uploads = wp_upload_dir();
        if (!empty($uploads['error'])) {
            return;
        }

        $sampleDirectory = trailingslashit($uploads['path']) . 'nf-xlsx-inline-samples';
        $sampleUrlBase   = trailingslashit($uploads['url']) . 'nf-xlsx-inline-samples';

        if (!wp_mkdir_p($sampleDirectory)) {
            return;
        }

        $imageSeed = nf_xlsx_seed_sample_asset($sampleDirectory, $sampleUrlBase, 'sample-image.png', nf_xlsx_sample_image_base64());
        $pdfSeed   = nf_xlsx_seed_sample_asset($sampleDirectory, $sampleUrlBase, 'sample.pdf', nf_xlsx_sample_pdf_base64());

        if (!$imageSeed || !$pdfSeed) {
            return;
        }

        $imageMeta = wp_json_encode([
            'file_name' => $imageSeed['name'],
            'file_path' => $imageSeed['path'],
            'url'       => $imageSeed['url'],
        ]);

        $pdfMeta = wp_json_encode([
            'file_name' => $pdfSeed['name'],
            'file_path' => $pdfSeed['path'],
            'url'       => $pdfSeed['url'],
        ]);

        $submissions = [
            [
                'id'       => 1,
                'sub_date' => current_time('mysql', true),
                'status'   => 'publish',
                'meta'     => [
                    'sample_text'  => [__('This workbook was generated on activation.', 'nf-cpt-xlsx-inline')],
                    'sample_image' => [$imageMeta],
                    'sample_pdf'   => [$pdfMeta],
                    'rich_text'    => [
                        sprintf(
                            __('Inline assets for review: %1$s and %2$s', 'nf-cpt-xlsx-inline'),
                            $imageSeed['url'],
                            $pdfSeed['url']
                        ),
                    ],
                ],
            ],
        ];

        if (!wp_mkdir_p($uploads['path'])) {
            return;
        }

        $filename = 'nf-xlsx-inline-sample-' . gmdate('Y-m-d-H-i-s') . '.xlsx';
        $filepath = trailingslashit($uploads['path']) . $filename;

        $exporter = new NF_XLSX_Stream_Exporter($form, $columns, $submissions);
        $exporter->save($filepath);
    } catch (Throwable $exception) {
        error_log('NF XLSX sample generation failed: ' . $exception->getMessage());
    }
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
            $key   = isset($row['key']) ? trim((string) $row['key']) : '';
            $label = isset($row['label']) ? trim((string) $row['label']) : '';

            if ($label === '' && $key !== '') {
                $label = $key;
            }

            $fields[] = [
                'id'         => (int) $row['id'],
                'key'        => $key,
                'label'      => $label,
                'type'       => isset($row['type']) ? $row['type'] : '',
                'identifier' => nf_xlsx_field_identifier((int) $row['id']),
            ];
        }
    }

    return array_values($fields);
}

function nf_xlsx_get_submissions($form_id, $limit = 0) {
    global $wpdb;

    $form_id     = (int) $form_id;
    $limit       = (int) $limit;
    $posts_table = $wpdb->posts;
    $meta_table  = $wpdb->postmeta;

    $order = $limit > 0 ? 'DESC' : 'ASC';

    $sql = $wpdb->prepare(
        "SELECT p.ID AS sub_id, p.post_date AS sub_date, p.post_status
         FROM {$posts_table} p
         INNER JOIN {$meta_table} fm
             ON fm.post_id = p.ID
             AND fm.meta_key = '_form_id'
             AND fm.meta_value = %d
         WHERE p.post_type = 'nf_sub'
         ORDER BY p.ID {$order}",
        $form_id
    );

    if ($limit > 0) {
        $sql .= $wpdb->prepare(' LIMIT %d', $limit);
    }

    $rows = $wpdb->get_results($sql, ARRAY_A);

    if (!$rows) {
        return [];
    }

    $submissions = [];
    $submissionIds = [];

    foreach ($rows as $row) {
        $id = (int) $row['sub_id'];

        if (!isset($submissions[$id])) {
            $submissions[$id] = [
                'id'       => $id,
                'sub_date' => $row['sub_date'],
                'status'   => isset($row['post_status']) ? (string) $row['post_status'] : '',
                'meta'     => [],
            ];
            $submissionIds[] = $id;
        }
    }

    if ($submissionIds) {
        $placeholders = implode(',', array_fill(0, count($submissionIds), '%d'));
        $metaSql      = $wpdb->prepare(
            "SELECT post_id, meta_key, meta_value
             FROM {$meta_table}
             WHERE post_id IN ({$placeholders})",
            $submissionIds
        );

        $metaRows = $wpdb->get_results($metaSql, ARRAY_A);

        if ($metaRows) {
            foreach ($metaRows as $metaRow) {
                $postId = (int) $metaRow['post_id'];

                if (!isset($submissions[$postId])) {
                    continue;
                }

                $metaKey = isset($metaRow['meta_key']) ? (string) $metaRow['meta_key'] : '';

                if (!isset($submissions[$postId]['meta'][$metaKey])) {
                    $submissions[$postId]['meta'][$metaKey] = [];
                }

                $submissions[$postId]['meta'][$metaKey][] = $metaRow['meta_value'];
            }
        }
    }

    return array_values($submissions);
}

// -----------------------------------------------------------------------------
// Spreadsheet builder.
// -----------------------------------------------------------------------------
function nf_xlsx_build_workbook(array $form, array $fields, array $submissions, array $selectedColumnIds) {
    $columns = nf_xlsx_prepare_columns($fields, $selectedColumnIds);

    if (!$columns) {
        throw new RuntimeException(__('No columns selected for export.', 'nf-cpt-xlsx-inline'));
    }

    return new NF_XLSX_Stream_Exporter($form, $columns, $submissions);
}

function nf_xlsx_prepare_columns(array $fields, array $selectedColumnIds = []) {
    $columns         = [];
    $usedHeaders     = [];
    $fallbackCounter = 1;

    $firstHeader = nf_xlsx_register_unique_header(__('Submission Date', 'nf-cpt-xlsx-inline'), $usedHeaders);
    $columns[]   = [
        'id'     => 'submission_date',
        'field'  => null,
        'header' => $firstHeader,
        'index'  => 1,
        'letter' => nf_xlsx_column_letter_from_position(1),
    ];

    $position = 2;

    foreach ($fields as $field) {
        $header = nf_xlsx_resolve_field_header($field, $usedHeaders, $fallbackCounter);

        $columns[] = [
            'id'     => isset($field['identifier']) ? (string) $field['identifier'] : nf_xlsx_field_identifier($field['id']),
            'field'  => $field,
            'header' => $header,
            'index'  => $position,
            'letter' => nf_xlsx_column_letter_from_position($position),
        ];

        ++$position;
    }

    if ($selectedColumnIds) {
        $selectedLookup = array_fill_keys(array_map('strval', $selectedColumnIds), true);
        $columns        = array_values(array_filter($columns, static function ($column) use ($selectedLookup) {
            $id = isset($column['id']) ? (string) $column['id'] : '';
            return isset($selectedLookup[$id]);
        }));
    }

    foreach ($columns as $index => &$column) {
        $column['index']  = $index + 1;
        $column['letter'] = nf_xlsx_column_letter_from_position($column['index']);
    }
    unset($column);

    return $columns;
}

function nf_xlsx_resolve_field_header(array $field, array &$usedHeaders, int &$fallbackCounter) {
    $candidates = [];

    if (isset($field['label'])) {
        $candidates[] = $field['label'];
    }

    if (isset($field['key'])) {
        $candidates[] = $field['key'];
    }

    $label = '';

    foreach ($candidates as $candidate) {
        $candidate = trim((string) $candidate);

        if ($candidate !== '') {
            $label = $candidate;
            break;
        }
    }

    if ($label === '') {
        $label = 'Column_' . $fallbackCounter;
        ++$fallbackCounter;
    }

    return nf_xlsx_register_unique_header($label, $usedHeaders);
}

function nf_xlsx_register_unique_header($label, array &$usedHeaders) {
    $label     = trim((string) $label);
    $baseLabel = $label !== '' ? $label : 'Column_' . (count($usedHeaders) + 1);
    $label     = $baseLabel;
    $suffix    = 2;
    $key       = nf_xlsx_normalize_header_key($label);

    while (isset($usedHeaders[$key])) {
        $label = $baseLabel . ' (' . $suffix . ')';
        $key   = nf_xlsx_normalize_header_key($label);
        ++$suffix;
    }

    $usedHeaders[$key] = true;

    return $label;
}

function nf_xlsx_normalize_header_key($label) {
    if (function_exists('mb_strtolower')) {
        return mb_strtolower($label, 'UTF-8');
    }

    return strtolower($label);
}

function nf_xlsx_column_letter_from_position($position) {
    return nf_col_from_index($position);
}

// -----------------------------------------------------------------------------
// Submission parsing helpers.
// -----------------------------------------------------------------------------
function nf_xlsx_extract_submission_field_payload(array $submission, array $field) {
    $meta       = isset($submission['meta']) && is_array($submission['meta']) ? $submission['meta'] : [];
    $candidates = nf_xlsx_candidate_meta_keys($field);

    foreach ($candidates as $candidate) {
        if (!isset($meta[$candidate])) {
            continue;
        }

        $values  = (array) $meta[$candidate];
        $payload = nf_xlsx_normalize_meta_values($values);

        if ($payload['text'] !== '' || !empty($payload['links'])) {
            return $payload;
        }
    }

    return ['text' => '', 'links' => [], 'images' => [], 'pdfs' => []];
}

function nf_xlsx_candidate_meta_keys(array $field) {
    $candidates = [];

    if (!empty($field['key'])) {
        $candidates[] = (string) $field['key'];
    }

    $candidates[] = (string) $field['id'];
    $candidates[] = 'field_' . $field['id'];
    $candidates[] = '_field_' . $field['id'];

    return array_values(array_unique(array_filter($candidates, static function ($value) {
        return $value !== '';
    })));
}

function nf_xlsx_normalize_meta_values(array $values) {
    $texts  = [];
    $links  = [];
    $images = [];
    $pdfs   = [];

    foreach ($values as $value) {
        $decoded = nf_xlsx_decode_meta_value($value);
        $payload = nf_xlsx_prepare_value_payload($decoded);

        if ($payload['text'] !== '') {
            $texts[] = $payload['text'];
        }

        if (!empty($payload['links'])) {
            $links = array_merge($links, $payload['links']);
        }

        if (!empty($payload['images'])) {
            $images = array_merge($images, $payload['images']);
        }

        if (!empty($payload['pdfs'])) {
            $pdfs = array_merge($pdfs, $payload['pdfs']);
        }
    }

    $texts = array_values(array_filter($texts, static function ($text) {
        return $text !== '';
    }));
    $texts = array_values(array_unique($texts));

    $links = array_values(array_unique($links));
    $images = array_values(array_unique($images));
    $pdfs   = array_values(array_unique($pdfs));

    if (!$texts && $links) {
        $texts = $links;
    }

    return [
        'text'   => $texts ? implode("\n", $texts) : '',
        'links'  => $links,
        'images' => $images,
        'pdfs'   => $pdfs,
    ];
}

function nf_xlsx_decode_meta_value($value) {
    if (is_string($value)) {
        $maybe = maybe_unserialize($value);
    } else {
        $maybe = $value;
    }

    if (is_string($maybe)) {
        $decoded = json_decode($maybe, true);
        if (json_last_error() === JSON_ERROR_NONE) {
            return $decoded;
        }
    }

    return $maybe;
}

function nf_xlsx_prepare_value_payload($value) {
    $payload = [
        'text'   => '',
        'links'  => [],
        'images' => [],
        'pdfs'   => [],
    ];

    if (is_array($value)) {
        if (nf_xlsx_is_upload_payload($value)) {
            $fileUrl         = nf_xlsx_resolve_file_url($value);
            $payload['text'] = nf_xlsx_guess_file_label($value, $fileUrl);

            if ($payload['text'] === '' && $fileUrl) {
                $payload['text'] = $fileUrl;
            }

            if ($fileUrl && nf_xlsx_is_linkable_url($fileUrl)) {
                $payload['links'][] = $fileUrl;

                if (nf_xlsx_is_image_url($fileUrl)) {
                    $payload['images'][] = $fileUrl;
                }

                if (nf_xlsx_is_pdf_url($fileUrl)) {
                    $payload['pdfs'][] = $fileUrl;
                }
            }
        } else {
            if (array_key_exists('value', $value)) {
                $valuePayload = nf_xlsx_prepare_value_payload($value['value']);

                if ($valuePayload['text'] === '' && !empty($value['label']) && is_scalar($value['label'])) {
                    $valuePayload['text'] = (string) $value['label'];
                }

                if ($valuePayload['text'] !== '' || !empty($valuePayload['links'])) {
                    if (!empty($valuePayload['links'])) {
                        $valuePayload['links'] = array_values(array_unique($valuePayload['links']));
                    }

                    if (!empty($valuePayload['images'])) {
                        $valuePayload['images'] = array_values(array_unique($valuePayload['images']));
                    }

                    if (!empty($valuePayload['pdfs'])) {
                        $valuePayload['pdfs'] = array_values(array_unique($valuePayload['pdfs']));
                    }

                    return $valuePayload;
                }
            }

            $texts = [];
            $links = [];
            $images = [];
            $pdfs   = [];

            foreach ($value as $key => $item) {
                if ($key === 'value') {
                    continue;
                }

                $itemPayload = nf_xlsx_prepare_value_payload($item);

                if ($itemPayload['text'] !== '') {
                    $texts[] = $itemPayload['text'];
                }

                if (!empty($itemPayload['links'])) {
                    $links = array_merge($links, $itemPayload['links']);
                }

                if (!empty($itemPayload['images'])) {
                    $images = array_merge($images, $itemPayload['images']);
                }

                if (!empty($itemPayload['pdfs'])) {
                    $pdfs = array_merge($pdfs, $itemPayload['pdfs']);
                }
            }

            if ($texts) {
                $payload['text'] = implode(', ', $texts);
            }

            if ($links) {
                $payload['links'] = array_values(array_unique($links));
            }

            if ($images) {
                $payload['images'] = array_values(array_unique($images));
            }

            if ($pdfs) {
                $payload['pdfs'] = array_values(array_unique($pdfs));
            }
        }
    } elseif (is_scalar($value)) {
        $string = trim((string) $value);
        if ($string !== '') {
            $payload['text'] = $string;

            if (nf_xlsx_is_linkable_url($string)) {
                $payload['links'][] = $string;

                if (nf_xlsx_is_image_url($string)) {
                    $payload['images'][] = $string;
                }

                if (nf_xlsx_is_pdf_url($string)) {
                    $payload['pdfs'][] = $string;
                }
            }
        }
    }

    if (!empty($payload['links'])) {
        $payload['links'] = array_values(array_unique(array_map('trim', $payload['links'])));
    }

    if (!empty($payload['images'])) {
        $payload['images'] = array_values(array_unique(array_map('trim', $payload['images'])));
    }

    if (!empty($payload['pdfs'])) {
        $payload['pdfs'] = array_values(array_unique(array_map('trim', $payload['pdfs'])));
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

    if (isset($value['value'])) {
        $maybe = $value['value'];

        if (is_string($maybe) && (nf_xlsx_is_linkable_url($maybe) || nf_xlsx_convert_path_to_url($maybe))) {
            return true;
        }

        if (is_array($maybe) && nf_xlsx_is_upload_payload($maybe)) {
            return true;
        }

        if (is_array($maybe)) {
            foreach ($maybe as $subMaybe) {
                if (is_array($subMaybe) && nf_xlsx_is_upload_payload($subMaybe)) {
                    return true;
                }

                if (is_string($subMaybe) && nf_xlsx_is_linkable_url($subMaybe)) {
                    return true;
                }
            }
        }
    }

    if (isset($value[0]) && is_array($value[0])) {
        foreach ($value[0] as $key => $unused) {
            if (in_array($key, $uploadKeys, true)) {
                return true;
            }
        }

        if (nf_xlsx_is_upload_payload($value[0])) {
            return true;
        }
    }

    return false;
}

function nf_xlsx_resolve_file_url(array $value) {
    $candidates = [];

    foreach (['url', 'value', 'file_url'] as $key) {
        if (!empty($value[$key]) && is_string($value[$key])) {
            $candidates[] = $value[$key];
        }
    }

    foreach (['file_path', 'path', 'tmp_name', 'saved_name'] as $key) {
        if (!empty($value[$key]) && is_string($value[$key])) {
            $candidates[] = $value[$key];
        }
    }

    if (isset($value[0]) && is_array($value[0])) {
        $nested = nf_xlsx_resolve_file_url($value[0]);
        if ($nested) {
            $candidates[] = $nested;
        }
    }

    foreach ($candidates as $candidate) {
        $candidate = trim((string) $candidate);

        if ($candidate === '') {
            continue;
        }

        if (nf_xlsx_is_linkable_url($candidate)) {
            return $candidate;
        }

        $maybe = nf_xlsx_convert_path_to_url($candidate);
        if ($maybe && nf_xlsx_is_linkable_url($maybe)) {
            return $maybe;
        }
    }

    $path = nf_xlsx_locate_file_path($value);
    if ($path) {
        $maybe = nf_xlsx_convert_path_to_url($path);
        if ($maybe && nf_xlsx_is_linkable_url($maybe)) {
            return $maybe;
        }
    }

    return '';
}

function nf_xlsx_convert_path_to_url($path) {
    if (!is_string($path) || $path === '') {
        return '';
    }

    $uploads = wp_upload_dir();
    if (!empty($uploads['error'])) {
        return '';
    }

    $baseDir = wp_normalize_path(trailingslashit($uploads['basedir']));
    $baseUrl = trailingslashit($uploads['baseurl']);
    $path    = wp_normalize_path($path);

    if (strpos($path, $baseDir) === 0) {
        $relative = ltrim(substr($path, strlen($baseDir)), '/');
        return $baseUrl . str_replace('\\', '/', $relative);
    }

    if (strpos($path, '/') === 0) {
        $maybe = $baseDir . ltrim($path, '/');
        if (file_exists($maybe)) {
            $normalized = wp_normalize_path($maybe);
            $relative   = ltrim(substr($normalized, strlen($baseDir)), '/');
            return $baseUrl . str_replace('\\', '/', $relative);
        }
    }

    return '';
}

function nf_xlsx_guess_file_label($value, $fileReference) {
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
            return nf_xlsx_guess_file_label($value[0], $fileReference);
        }
    }

    if ($fileReference) {
        $path = parse_url($fileReference, PHP_URL_PATH);
        if ($path) {
            return basename($path);
        }
    }

    return '';
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

function nf_xlsx_is_linkable_url($url) {
    $extension = nf_xlsx_url_extension($url);

    if ($extension === '') {
        return false;
    }

    return nf_xlsx_is_image_extension($extension) || nf_xlsx_is_pdf_extension($extension);
}

function nf_xlsx_is_image_url($url) {
    $extension = nf_xlsx_url_extension($url);

    return nf_xlsx_is_image_extension($extension);
}

function nf_xlsx_is_pdf_url($url) {
    return nf_xlsx_url_extension($url) === 'pdf';
}

function nf_xlsx_is_image_extension($extension) {
    if (!is_string($extension) || $extension === '') {
        return false;
    }

    $extension = strtolower($extension);

    return in_array($extension, ['jpg', 'jpeg', 'png', 'gif', 'webp'], true);
}

function nf_xlsx_is_pdf_extension($extension) {
    return is_string($extension) && strtolower($extension) === 'pdf';
}

function nf_xlsx_url_extension($url) {
    if (!is_string($url) || $url === '') {
        return '';
    }

    $url = trim($url);
    if ($url === '') {
        return '';
    }

    $path = parse_url($url, PHP_URL_PATH);
    if (!$path) {
        return '';
    }

    return strtolower(pathinfo($path, PATHINFO_EXTENSION));
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

function nf_col_from_index($index) {
    $index = (int) $index;

    if ($index < 1) {
        $index = 1;
    }

    $letters = '';

    while ($index > 0) {
        $index--;
        $letters = chr($index % 26 + 65) . $letters;
        $index   = (int) floor($index / 26);
    }

    return $letters !== '' ? $letters : 'A';
}

function nf_addr($column, $row) {
    if (is_numeric($column)) {
        $column = nf_col_from_index((int) $column);
    } else {
        $column = strtoupper(trim((string) $column));

        if ($column === '') {
            $column = 'A';
        }
    }

    $row = (int) $row;
    if ($row < 1) {
        $row = 1;
    }

    return $column . $row;
}

function nf_xlsx_field_identifier($fieldId) {
    return 'field-' . (int) $fieldId;
}

function nf_xlsx_normalize_column_selection(array $selectedColumns, array $availableColumns) {
    $selectedColumns = array_values(array_unique(array_filter(array_map('strval', $selectedColumns), static function ($value) {
        return $value !== '';
    })));

    if (!$selectedColumns) {
        return [];
    }

    $selectedLookup = array_fill_keys($selectedColumns, true);
    $normalized     = [];

    foreach ($availableColumns as $column) {
        $id = isset($column['id']) ? (string) $column['id'] : '';
        if ($id === '') {
            continue;
        }

        if (isset($selectedLookup[$id])) {
            $normalized[] = $id;
        }
    }

    return $normalized;
}
