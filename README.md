# NF CPT → XLSX Inline Export

WordPress plugin that exports post (or custom post type) data to a downloadable XLSX file using the bundled PhpSpreadsheet library—no Composer step required.

## Features

- Bundled PhpSpreadsheet 1.29 classes for offline-friendly deployments.
- Secure admin-post export endpoint protected by capability and nonce checks.
- Auto-generated Excel headers/rows with filter hook (`nf_cpt_xlsx_inline_export_data`) for custom datasets.
- Optional image embedding support via the filtered payload.

## Usage

1. Upload the plugin folder to `wp-content/plugins/`.
2. Activate **NF CPT → XLSX Inline Export** from the WordPress dashboard.
3. Visit **Tools → NF → XLSX** (or the top-level menu depending on your admin menu setup).
4. Click **Export Test XLSX** to download the spreadsheet. Use the filter hook in theme/plugin code to provide custom data.

```
add_filter('nf_cpt_xlsx_inline_export_data', function ($payload) {
    $payload['headers'] = ['ID', 'Title', 'Custom'];
    $payload['rows'] = [
        ['1', 'Example', 'Value'],
    ];

    return $payload;
});
```

## Version

- **1.0.6** — Stable release with nonce-protected export endpoint and default CPT dataset loader.
