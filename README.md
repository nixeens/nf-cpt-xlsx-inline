# NF Export - codex

This WordPress plugin bundles PhpSpreadsheet so editors can export any public post type directly from the admin without installing Composer on the server.

## Features

- WordPress-style autoloader that maps bundled `PhpOffice`, `Psr\SimpleCache`, and `Composer\Pcre` classes.
- Admin page under **NF → XLSX** that lets administrators pick the target post type and tweak export options.
- Exports ID, Title, Status, Author, Date, and Permalink columns out of the box.
- Packages `export.xlsx`, the referenced attachments, and `manifest.json` into a single ZIP archive with sequential attachment names.
- Filters (`nf_cpt_xlsx_inline_post_types`, `nf_cpt_xlsx_inline_headers`, `nf_cpt_xlsx_inline_row`, `nf_cpt_xlsx_inline_query_args`) for deep customization.
- Inline XLSX streaming with repaired one-cell anchors and worksheet styling that avoids Excel “repair” prompts.

## Usage

1. Upload/activate the plugin.
2. Navigate to **NF → XLSX Export** in wp-admin.
3. Pick the post type you want to export and click **Export to XLSX**.
4. Extend with hooks if you need to add/remove columns or tweak the query.

> **Note:** The bundled libraries live in `/lib` so the plugin is self-contained on hosts without Composer.
