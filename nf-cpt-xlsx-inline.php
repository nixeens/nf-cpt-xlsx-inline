<?php
/**
 * Plugin Name: NF Export - Codex V2
 * Description: Export custom post type data to XLSX with bundled attachments.
 * Version: 2.0.0
 * Author: Your Name
 * License: GPL2
 */

namespace CodexV2;

use RuntimeException;
use WP_Post;
use WP_Query;

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
\add_action('plugins_loaded', __NAMESPACE__ . '\\nf_xlsx_register_local_autoloaders', 1);

require_once __DIR__ . '/class-nf-xlsx-stream-exporter.php';

// -----------------------------------------------------------------------------
// Admin UI registration.
// -----------------------------------------------------------------------------
\add_action('admin_menu', static function () {
    \add_menu_page(
        __('NF → XLSX Export', 'nf-cpt-xlsx-inline'),
        __('NF → XLSX Export', 'nf-cpt-xlsx-inline'),
        'manage_options',
        'nf-cpt-xlsx-inline',
        __NAMESPACE__ . '\\nf_xlsx_render_admin_page',
        'dashicons-media-spreadsheet',
        58
    );
});

\add_action('admin_post_nf_xlsx_export', __NAMESPACE__ . '\\nf_xlsx_handle_export');

\add_action('admin_notices', static function () {
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
                esc_html__('Export generated successfully.', 'nf-cpt-xlsx-inline'),
                esc_url($url),
                esc_html__('Download ZIP', 'nf-cpt-xlsx-inline')
            );
        } else {
            printf(
                '<div class="notice notice-success"><p>%s</p></div>',
                esc_html__('Export generated successfully.', 'nf-cpt-xlsx-inline')
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

    $postTypes = nf_cpt_xlsx_inline_get_post_types();
    $selected  = isset($_GET['post_type']) ? sanitize_key(wp_unslash($_GET['post_type'])) : '';

    if (!$selected && !empty($postTypes)) {
        $firstKey = array_key_first($postTypes);
        if ($firstKey !== null) {
            $selected = (string) $firstKey;
        }
    }

    if ($selected && !isset($postTypes[$selected])) {
        $selected = '';
    }

    $zipPreview = nf_cpt_xlsx_inline_next_export_name();
    $includeAttachments = true;
    $treatSignatures    = true;
    ?>
    <div class="wrap">
        <h1><?php esc_html_e('NF → XLSX Export', 'nf-cpt-xlsx-inline'); ?></h1>
        <p><?php esc_html_e('Generate an export.zip containing export.xlsx, all referenced attachments, and a manifest.', 'nf-cpt-xlsx-inline'); ?></p>

        <?php if (empty($postTypes)) : ?>
            <div class="notice notice-warning"><p><?php esc_html_e('No public post types available for export.', 'nf-cpt-xlsx-inline'); ?></p></div>
        <?php else : ?>
            <form method="post" action="<?php echo esc_url(admin_url('admin-post.php')); ?>" class="nf-xlsx-options-form">
                <?php wp_nonce_field('nf_cpt_xlsx_inline_export', '_nf_cpt_xlsx_inline_nonce'); ?>
                <input type="hidden" name="action" value="nf_xlsx_export">
                <table class="form-table" role="presentation">
                    <tbody>
                        <tr>
                            <th scope="row"><label for="nf-xlsx-post-type"><?php esc_html_e('Post type', 'nf-cpt-xlsx-inline'); ?></label></th>
                            <td>
                                <select name="post_type" id="nf-xlsx-post-type">
                                    <?php foreach ($postTypes as $slug => $label) : ?>
                                        <option value="<?php echo esc_attr($slug); ?>" <?php selected($selected, $slug); ?>><?php echo esc_html($label); ?></option>
                                    <?php endforeach; ?>
                                </select>
                            </td>
                        </tr>
                        <tr>
                            <th scope="row"><?php esc_html_e('Options', 'nf-cpt-xlsx-inline'); ?></th>
                            <td>
                                <fieldset>
                                    <legend class="screen-reader-text"><?php esc_html_e('Export options', 'nf-cpt-xlsx-inline'); ?></legend>
                                    <label for="nf-xlsx-include-attachments">
                                        <input type="checkbox" id="nf-xlsx-include-attachments" name="include_attachments" value="1" <?php checked($includeAttachments); ?>>
                                        <?php esc_html_e('Include attachments in ZIP package', 'nf-cpt-xlsx-inline'); ?>
                                    </label>
                                    <br>
                                    <label for="nf-xlsx-treat-signatures">
                                        <input type="checkbox" id="nf-xlsx-treat-signatures" name="treat_signatures" value="1" <?php checked($treatSignatures); ?>>
                                        <?php esc_html_e('Treat signatures separately from other images', 'nf-cpt-xlsx-inline'); ?>
                                    </label>
                                </fieldset>
                                <p class="description"><?php esc_html_e('Attachments will be renamed sequentially (attachment1.ext, attachment2.ext, …).', 'nf-cpt-xlsx-inline'); ?></p>
                            </td>
                        </tr>
                        <tr>
                            <th scope="row"><?php esc_html_e('Next ZIP name', 'nf-cpt-xlsx-inline'); ?></th>
                            <td>
                                <input type="text" readonly class="regular-text" value="<?php echo esc_attr($zipPreview); ?>">
                                <p class="description"><?php esc_html_e('The final filename may include a numeric suffix if needed to avoid collisions.', 'nf-cpt-xlsx-inline'); ?></p>
                            </td>
                        </tr>
                    </tbody>
                </table>
                <?php submit_button(__('Export to XLSX (ZIP)', 'nf-cpt-xlsx-inline')); ?>
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

    check_admin_referer('nf_cpt_xlsx_inline_export', '_nf_cpt_xlsx_inline_nonce');

    $postType = isset($_POST['post_type']) ? sanitize_key(wp_unslash($_POST['post_type'])) : '';

    $postTypes = nf_cpt_xlsx_inline_get_post_types();
    if (!$postType || !isset($postTypes[$postType])) {
        nf_xlsx_redirect_error(__('Invalid post type selection.', 'nf-cpt-xlsx-inline'));
    }

    $includeAttachments = !empty($_POST['include_attachments']);
    $treatSignatures    = !empty($_POST['treat_signatures']);

    $options = [
        'include_attachments'        => (bool) $includeAttachments,
        'treat_signatures_separately' => (bool) $treatSignatures,
    ];

    try {
        $headers = nf_cpt_xlsx_inline_prepare_headers($postType);
        if (empty($headers)) {
            throw new RuntimeException(__('No columns available for export.', 'nf-cpt-xlsx-inline'));
        }

        $postIds = nf_cpt_xlsx_inline_query_post_ids($postType);
        $rows    = nf_cpt_xlsx_inline_build_rows($postIds, $headers, $options);

        list($rowsWithAttachments) = nf_cpt_xlsx_inline_assign_zip_names($rows);

        $exporter = new NF_XLSX_Stream_Exporter($headers, $rowsWithAttachments, $options);

        $uploads = wp_upload_dir();
        if (!empty($uploads['error'])) {
            throw new RuntimeException($uploads['error']);
        }

        if (!wp_mkdir_p($uploads['path'])) {
            throw new RuntimeException(__('Unable to create upload directory.', 'nf-cpt-xlsx-inline'));
        }

        $baseFilename = nf_cpt_xlsx_inline_next_export_name();
        $filename     = function_exists('wp_unique_filename')
            ? wp_unique_filename($uploads['path'], $baseFilename)
            : $baseFilename;

        $zipPath = trailingslashit($uploads['path']) . $filename;

        $manifest = $exporter->exportZip($zipPath);

        do_action('nf_cpt_xlsx_inline_after_export', $zipPath, $manifest, $options);

        $redirectArgs = [
            'page'           => 'nf-cpt-xlsx-inline',
            'post_type'      => $postType,
            'nf_xlsx_notice' => 'success',
            'nf_xlsx_file'   => rawurlencode($filename),
        ];

        $redirect = add_query_arg($redirectArgs, admin_url('admin.php'));
        wp_safe_redirect($redirect);
        exit;
    } catch (RuntimeException $exception) {
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
function nf_cpt_xlsx_inline_get_post_types(): array {
    $objects = get_post_types([
        'public' => true,
    ], 'objects');

    $postTypes = [];
    foreach ($objects as $slug => $object) {
        $label = $object->labels->singular_name ?? $object->label ?? $slug;
        $postTypes[$slug] = $label;
    }

    $postTypes = apply_filters('nf_cpt_xlsx_inline_post_types', $postTypes);

    if (is_array($postTypes)) {
        asort($postTypes, SORT_FLAG_CASE | SORT_STRING);
    } else {
        $postTypes = [];
    }

    return $postTypes;
}

function nf_cpt_xlsx_inline_prepare_headers(string $postType): array {
    $defaultHeaders = [
        ['key' => 'id', 'label' => __('ID', 'nf-cpt-xlsx-inline')],
        ['key' => 'title', 'label' => __('Title', 'nf-cpt-xlsx-inline')],
        ['key' => 'status', 'label' => __('Status', 'nf-cpt-xlsx-inline')],
        ['key' => 'author', 'label' => __('Author', 'nf-cpt-xlsx-inline')],
        ['key' => 'date', 'label' => __('Date', 'nf-cpt-xlsx-inline')],
        ['key' => 'permalink', 'label' => __('Permalink', 'nf-cpt-xlsx-inline')],
    ];

    $headers = apply_filters('nf_cpt_xlsx_inline_headers', $defaultHeaders, $postType);
    if (!is_array($headers)) {
        return $defaultHeaders;
    }

    $normalized = [];
    foreach ($headers as $header) {
        if (is_string($header)) {
            $key   = sanitize_key($header);
            $label = $header;
        } elseif (is_array($header)) {
            $key   = isset($header['key']) ? sanitize_key((string) $header['key']) : '';
            $label = isset($header['label']) ? (string) $header['label'] : '';
        } else {
            continue;
        }

        if ($key === '' || $label === '') {
            continue;
        }

        $normalized[] = [
            'key'   => $key,
            'label' => $label,
        ];
    }

    return $normalized;
}

function nf_cpt_xlsx_inline_query_post_ids(string $postType): array {
    $queryArgs = [
        'post_type'              => $postType,
        'posts_per_page'         => -1,
        'post_status'            => 'any',
        'orderby'                => 'ID',
        'order'                  => 'ASC',
        'fields'                 => 'ids',
        'no_found_rows'          => true,
        'update_post_meta_cache' => false,
        'update_post_term_cache' => false,
    ];

    $queryArgs = apply_filters('nf_cpt_xlsx_inline_query_args', $queryArgs, $postType);

    $query = new WP_Query($queryArgs);
    $ids   = $query->posts;
    wp_reset_postdata();

    if (!is_array($ids)) {
        return [];
    }

    return array_map('intval', $ids);
}

function nf_cpt_xlsx_inline_build_rows(array $postIds, array $headers, array $options): array {
    $rows = [];

    foreach ($postIds as $postId) {
        $post = get_post($postId);
        if (!$post instanceof WP_Post) {
            continue;
        }

        $defaultValues = nf_cpt_xlsx_inline_default_row_values($post);
        $rowValues     = [];

        foreach ($headers as $header) {
            $key          = $header['key'];
            $rowValues[$key] = isset($defaultValues[$key]) ? (string) $defaultValues[$key] : '';
        }

        $attachments = nf_cpt_xlsx_inline_collect_post_attachments($post->ID, $options);

        $row = [
            'post_id'     => (int) $post->ID,
            'values'      => $rowValues,
            'attachments' => $attachments,
        ];

        $row = apply_filters('nf_cpt_xlsx_inline_row', $row, $post, $headers, $options);

        if (!is_array($row) || empty($row['values'])) {
            continue;
        }

        $row['values'] = nf_cpt_xlsx_inline_merge_row_values($rowValues, $row['values']);
        $row['attachments'] = nf_cpt_xlsx_inline_normalize_attachments($row['attachments'] ?? [], $row['post_id']);

        $rows[] = $row;
    }

    return $rows;
}

function nf_cpt_xlsx_inline_default_row_values(WP_Post $post): array {
    $authorName = '';
    if ($post->post_author) {
        $authorName = get_the_author_meta('display_name', $post->post_author);
    }

    $dateGmt = $post->post_date_gmt;
    if (!$dateGmt && $post->post_date) {
        $dateGmt = get_gmt_from_date($post->post_date);
    }

    $dateDisplay = '';
    if ($dateGmt) {
        $dateDisplay = get_date_from_gmt($dateGmt, 'Y-m-d H:i:s');
    }

    return [
        'id'        => (string) $post->ID,
        'title'     => get_the_title($post),
        'status'    => (string) $post->post_status,
        'author'    => $authorName,
        'date'      => $dateDisplay,
        'permalink' => get_permalink($post),
    ];
}

function nf_cpt_xlsx_inline_collect_post_attachments(int $postId, array $options): array {
    $media = get_attached_media('', $postId);

    if (!$media) {
        return [];
    }

    $items = array_values($media);
    usort($items, static function (WP_Post $a, WP_Post $b): int {
        if ($a->menu_order === $b->menu_order) {
            return $a->ID <=> $b->ID;
        }

        return $a->menu_order <=> $b->menu_order;
    });

    $attachments = [];
    foreach ($items as $attachment) {
        $path = get_attached_file($attachment->ID);
        if (!$path || !file_exists($path)) {
            continue;
        }

        $mime = get_post_mime_type($attachment->ID);
        $type = nf_cpt_xlsx_inline_classify_mime($mime);
        $url  = wp_get_attachment_url($attachment->ID);

        $isSignature = apply_filters('nf_cpt_xlsx_is_signature', false, $attachment->ID);

        $attachments[] = [
            'attachment_id'    => (int) $attachment->ID,
            'path'             => $path,
            'source_url'       => $url ?: '',
            'mime'             => $mime ?: '',
            'type'             => $type,
            'is_signature'     => (bool) $isSignature,
            'original_filename'=> wp_basename($path),
            'post_id'          => $postId,
        ];
    }

    return $attachments;
}

function nf_cpt_xlsx_inline_classify_mime(?string $mime): string {
    $mime = is_string($mime) ? strtolower(trim($mime)) : '';

    if ($mime !== '' && strpos($mime, 'image/') === 0) {
        return 'image';
    }

    if ($mime === 'application/pdf') {
        return 'pdf';
    }

    return 'other';
}

function nf_cpt_xlsx_inline_merge_row_values(array $defaults, array $overrides): array {
    $merged = $defaults;
    foreach ($overrides as $key => $value) {
        $key = sanitize_key((string) $key);
        if ($key === '') {
            continue;
        }
        $merged[$key] = is_scalar($value) ? (string) $value : wp_json_encode($value);
    }

    return $merged;
}

function nf_cpt_xlsx_inline_normalize_attachments($attachments, int $postId): array {
    if (!is_array($attachments)) {
        return [];
    }

    $normalized = [];

    foreach ($attachments as $attachment) {
        if (!is_array($attachment)) {
            continue;
        }

        $path = isset($attachment['path']) ? (string) $attachment['path'] : '';
        if ($path === '' || !file_exists($path)) {
            continue;
        }

        $mime = isset($attachment['mime']) ? (string) $attachment['mime'] : '';
        $type = isset($attachment['type']) ? (string) $attachment['type'] : nf_cpt_xlsx_inline_classify_mime($mime);
        if (!in_array($type, ['image', 'pdf', 'other'], true)) {
            $type = 'other';
        }

        $normalized[] = [
            'attachment_id'     => isset($attachment['attachment_id']) ? (int) $attachment['attachment_id'] : null,
            'path'              => $path,
            'source_url'        => isset($attachment['source_url']) ? (string) $attachment['source_url'] : '',
            'mime'              => $mime,
            'type'              => $type,
            'is_signature'      => !empty($attachment['is_signature']),
            'original_filename' => isset($attachment['original_filename']) && $attachment['original_filename'] !== ''
                ? (string) $attachment['original_filename']
                : wp_basename($path),
            'post_id'           => isset($attachment['post_id']) ? (int) $attachment['post_id'] : $postId,
        ];
    }

    return $normalized;
}

function nf_cpt_xlsx_inline_assign_zip_names(array $rows): array {
    $counter = 0;

    foreach ($rows as &$row) {
        if (empty($row['attachments']) || !is_array($row['attachments'])) {
            continue;
        }

        foreach ($row['attachments'] as &$attachment) {
            $extension = nf_cpt_xlsx_inline_attachment_extension($attachment);
            ++$counter;
            $attachment['zip_name'] = sprintf('attachment%d.%s', $counter, $extension);
        }
        unset($attachment);
    }
    unset($row);

    return [$rows];
}

function nf_cpt_xlsx_inline_attachment_extension(array $attachment): string {
    $filename = isset($attachment['original_filename']) ? (string) $attachment['original_filename'] : '';
    $extension = strtolower(pathinfo($filename, PATHINFO_EXTENSION));

    if ($extension === '' && !empty($attachment['path'])) {
        $extension = strtolower(pathinfo((string) $attachment['path'], PATHINFO_EXTENSION));
    }

    if ($extension === '' && !empty($attachment['mime'])) {
        $extension = nf_cpt_xlsx_inline_extension_from_mime((string) $attachment['mime']);
    }

    if ($extension === '') {
        $extension = 'bin';
    }

    return $extension;
}

function nf_cpt_xlsx_inline_extension_from_mime(string $mime): string {
    $map = [
        'image/jpeg'      => 'jpg',
        'image/jpg'       => 'jpg',
        'image/png'       => 'png',
        'image/gif'       => 'gif',
        'image/webp'      => 'webp',
        'application/pdf' => 'pdf',
    ];

    $mime = strtolower(trim($mime));

    return $map[$mime] ?? '';
}

function nf_cpt_xlsx_inline_next_export_name(): string {
    return sprintf('export-%s.zip', gmdate('Ymd-Hi'));
}
