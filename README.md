# NF CPT → XLSX Inline Export

This WordPress plugin bundles PhpSpreadsheet so editors can export any public post type directly from the admin without installing Composer on the server.

## Features

- WordPress-style autoloader that maps bundled `PhpOffice`, `Psr\SimpleCache`, and `Composer\Pcre` classes.
- Admin page under **NF → XLSX** that lets administrators pick the target post type.
- Exports ID, Title, Status, Author, Date, and Permalink columns out of the box.
- Filters (`nf_cpt_xlsx_inline_post_types`, `nf_cpt_xlsx_inline_headers`, `nf_cpt_xlsx_inline_row`, `nf_cpt_xlsx_inline_query_args`) for deep customization.
- Inline XLSX streaming with proper cache headers and worksheet styling.

## Usage

1. Upload/activate the plugin.
2. Navigate to **Tools → NF → XLSX** in wp-admin.
3. Pick the post type you want to export and click **Export to XLSX**.
4. Extend with hooks if you need to add/remove columns or tweak the query.

> **Note:** The bundled libraries live in `/lib` so the plugin is self-contained on hosts without Composer.
