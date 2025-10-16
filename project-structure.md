# NF CPT → XLSX Inline Export - Code Logic Analysis

## Overview

This WordPress plugin exports Ninja Forms submissions to Excel (.xlsx) files with support for inline images and PDF attachments. It bundles PhpSpreadsheet to work on servers without Composer.

## Architecture Components

### 1. Plugin Bootstrap (`nf-cpt-xlsx-inline.php`)

#### Autoloading System (Lines 17-78)

- **Purpose**: Load bundled libraries without Composer
- **Libraries**: PhpSpreadsheet, Psr\SimpleCache, Composer\Pcre, ZipStream
- **Implementation**: 4 separate `spl_autoload_register` handlers
- **Pattern**: PSR-4 autoloading from `/lib` directory
- **Execution**: Runs on `plugins_loaded` hook with priority 1

#### Admin Interface (Lines 87-290)

**Menu Registration** (Lines 87-97)

- Creates admin page: "NF Submissions Export"
- Menu icon: `dashicons-media-spreadsheet`
- Position: 58
- Capability: `manage_options`

**Admin Page Rendering** (`nf_xlsx_render_admin_page`, Lines 140-291)

- **Form Selection**: Dropdown of all Ninja Forms (from `nf3_forms` table)
- **Column Selection**: Checkboxes for each field + submission date
- **Preview**: First 5 submissions displayed in table
- **State Management**: Uses GET parameters (`form_id`, `columns[]`) to maintain selections

**Column Selection Logic** (Lines 158-181)

- Fetches available columns from form fields
- Normalizes user selection against available columns
- Defaults to all columns if none selected
- Maintains selection through preview updates

#### Export Handler (Lines 296-367)

**Security** (Lines 297-301)

- Capability check: `manage_options`
- Nonce verification: `nf_xlsx_export`

**Export Flow** (`nf_xlsx_handle_export`):

1. Validate form selection
2. Fetch form metadata and fields
3. Prepare columns based on user selection
4. Fetch all submissions for the form
5. Build workbook via `NF_XLSX_Stream_Exporter`
6. Save to uploads directory with timestamped filename
7. Redirect with success/error notice

**Error Handling** (Lines 363-366)

- Try-catch wrapper
- Logs to error_log
- Redirects with error message

#### Activation Hook (Lines 430-542)

**Sample Workbook Generation** (`nf_xlsx_generate_activation_sample`):

- Creates mock form with 4 fields (text, image, PDF, textarea)
- Generates base64-encoded sample image (64x64 PNG)
- Generates base64-encoded sample PDF
- Creates one sample submission
- Saves example XLSX to uploads directory
- Purpose: Smoke test for PhpSpreadsheet functionality

### 2. Data Access Layer

#### Ninja Forms Integration

**Forms Table** (`nf_xlsx_get_forms`, Lines 547-561)

- Table: `{prefix}nf3_forms`
- Columns: `id`, `title`
- Returns: Array of form_id => title

**Fields Table** (`nf_xlsx_get_form_fields`, Lines 579-609)

- Table: `{prefix}nf3_fields`
- Query: WHERE `parent_id` = form_id
- Columns: `id`, `key`, `label`, `type`
- Fallback: Uses `key` if `label` is empty
- Returns: Array of field metadata with identifiers

**Submissions** (`nf_xlsx_get_submissions`, Lines 611-691)

- Table: `{prefix}posts` (post_type = 'nf_sub')
- Join: `{prefix}postmeta` (meta_key = '_form_id')
- Returns: Submissions with all metadata
- Performance: Two-query approach (submissions first, then metadata batch fetch)
- Order: DESC for preview (first 5), ASC for export (all)

#### Column Preparation (Lines 706-751)

**`nf_xlsx_prepare_columns`**:

1. Adds "Submission Date" as first column
2. Iterates fields and assigns unique headers
3. Filters by `$selectedColumnIds` if provided
4. Re-indexes columns and assigns Excel letters (A, B, C...)
5. Returns: Array with `id`, `field`, `header`, `index`, `letter`

**Header Uniqueness** (`nf_xlsx_register_unique_header`, Lines 783-799)

- Prevents duplicate column headers
- Appends (2), (3)... to duplicates
- Case-insensitive comparison

### 3. Submission Parsing Engine

#### Field Value Extraction (Lines 816-834)

**`nf_xlsx_extract_submission_field_payload`**:

- Searches submission meta for field value using candidate keys
- Candidate order: `field.key`, `field.id`, `field_{id}`, `_field_{id}`
- Returns normalized payload with `text`, `links`, `images`, `pdfs`

#### Meta Value Normalization (Lines 852-898)

**`nf_xlsx_normalize_meta_values`**:

- Handles arrays of values (multi-select, file uploads)
- Decodes JSON and serialized data
- Aggregates text, links, images, PDFs
- Deduplicates and filters empty values
- Fallback: Uses first link as text if no text present

#### Value Payload Preparation (Lines 917-1075)

**`nf_xlsx_prepare_value_payload`** - Recursive parser:

**Array Handling**:

- Detects file upload payloads (keys: `tmp_name`, `file_path`, `url`, `files`)
- Processes nested `files` array
- Handles `value`/`label` structure
- Recursively processes all array elements

**Upload Detection** (`nf_xlsx_is_upload_payload`, Lines 1077-1127):

- Checks for upload-specific keys
- Searches nested structures
- Validates URLs and paths

**URL Resolution** (`nf_xlsx_resolve_file_url`, Lines 1129-1190):

- Priority: `url` > `file_path` > nested structures
- Converts local paths to URLs using `wp_upload_dir`
- Special handling for `ninja-forms/` subdirectory

**Path to URL Conversion** (`nf_xlsx_convert_path_to_url`, Lines 1192-1231):

- Maps filesystem paths to public URLs
- Handles uploads directory structure
- Validates file existence

**File Label Extraction** (`nf_xlsx_guess_file_label`, Lines 1233-1265):

- Searches: `file_name`, `saved_name`, `filename`, `name`
- Fallback: Basename from URL
- Used for display in Excel cells

### 4. Excel Export Engine (`class-nf-xlsx-stream-exporter.php`)

#### Architecture

**State Management** (Lines 17-38):

- `$form`, `$columns`, `$submissions`: Input data
- `$spreadsheet`, `$submissionsSheet`, `$attachmentsSheet`: PhpSpreadsheet objects
- `$imageCache`, `$pdfCache`: Prevents duplicate downloads
- `$tempFiles`: Tracks files for cleanup
- `$rowHeights`, `$cellOffsets`: Layout tracking for inline images

#### Initialization (Lines 66-83)

**`initialiseSheets`**:

1. Creates Spreadsheet instance
2. Configures Office theme
3. Sets Calibri 11pt as default font
4. Creates "Submissions" sheet
5. Adds headers (bold, wrap text)
6. Populates data or "No submissions" message
7. Freezes header row (pane A2)

#### Header Row Creation (Lines 84-100)

**`addSubmissionHeaders`**:

- Writes column headers from `$columns` array
- Applies bold font + center alignment
- Enables text wrapping
- Sets AutoSize for each column

#### Submissions Population (Lines 109-154)

**`populateSubmissions`** - Main loop:

- Iterates submissions (row 2+)
- For each column:
    - **Date column** (`field === null`): Formats submission date
    - **Field columns**: Extracts payload via helper functions
    - **Text**: Writes to cell with wrap text
    - **Links**: Adds hyperlink to first URL
    - **Images**: Calls `addImage()` for inline embedding
    - **PDFs**: Calls `addPdf()` for icon + link

#### Image Handling (Lines 165-211)

**`addImage`**:

1. **Cache check**: MD5 hash of URL
2. **Download**: HTTP request or local file read
3. **Validation**: `getimagesizefromstring()` check
4. **Scaling**: Max 240×220px, proportional
5. **Temp file**: Write image data to temp directory
6. **Drawing object**: PhpSpreadsheet Drawing with coordinates
7. **Offset management**: `reserveCellOffset()` for vertical stacking
8. **Row height**: Auto-adjusts for image height

**Image Fetching** (`fetch_image_bin`, Lines 556-600):

- HTTP request via `wp_remote_get` (or `file_get_contents` fallback)
- Extracts MIME type and dimensions
- Returns: `data`, `mime`, `extension`, `width`, `height`
- Fallback: Tries local file path conversion

#### PDF Handling (Lines 212-261)

**`addPdf`**:

1. **Download**: Fetches PDF binary (cached)
2. **Temp storage**: Writes PDF to temp file
3. **Icon embedding**: Uses base64-encoded PDF icon PNG
4. **Drawing**: 22px height icon with hyperlink
5. **Attachments sheet**: Logs PDF metadata in secondary sheet
6. **Offset tracking**: Stacks icons vertically in cell

**Attachments Sheet** (`ensureAttachmentsSheet`, Lines 283-310):

- Created on-demand when first PDF is added
- Columns: Row, Column, Original URL, Status
- Status: "Icon linked" or "Download failed"
- All URLs are hyperlinked
- Purpose: Reference table for external files

#### Layout Management

**Row Height** (`ensureRowHeight`, Lines 339-348):

- Tracks maximum height needed per row
- Converts pixels to points (72 DPI vs 96 DPI)
- Only increases height, never decreases

**Cell Offset** (`reserveCellOffset`, Lines 350-362):

- Tracks vertical position within cell for stacking
- Maintains `next` offset and `total` height per cell
- Returns offset Y coordinate for Drawing object
- Auto-adjusts row height based on content

#### Resource Management

**Temp Files** (`writeTempFile`, Lines 363-398):

- Uses `wp_tempnam` or `tempnam()`
- Adds extension to filename for MIME detection
- Tracks all paths in `$tempFiles` array

**Cleanup** (`cleanupTempFiles`, Lines 413-420):

- Called in `__destruct` and after save
- Deletes all temporary image/PDF files
- Prevents disk space leaks

**Disconnection** (Line 63):

- `$spreadsheet->disconnectWorksheets()` after save
- Frees memory for large exports

#### HTTP & File Operations

**HTTP Requests** (`perform_http_request`, Lines 622-676):

- Primary: `wp_remote_get` with 15s timeout
- Fallback: `file_get_contents` with stream context
- Accepts: `image/*,application/pdf`
- Returns body + content-type

**Local File Reading** (`read_local_file`, Lines 678-705):

- Converts URL to filesystem path
- Uses `wp_check_filetype` for MIME detection
- Fallback: `mime_content_type()`

**URL to Path Mapping** (`url_to_local_path`, Lines 707-748):

- Checks uploads directory (`wp_upload_dir`)
- Checks site root (`ABSPATH`)
- Checks content directory (`WP_CONTENT_DIR`)
- Validates file existence

#### Theme Configuration (Lines 464-491)

**`configureTheme`**:

- Sets Office theme colors
- Configures fonts: Cambria (major), Calibri (minor)
- Defines 12 theme colors (lt1, dk1, accents, hyperlinks)
- Resets theme fonts for consistency

### 5. Utility Functions

#### URL Detection (Lines 1353-1403)

- `nf_xlsx_is_linkable_url`: Checks for image/PDF extensions
- `nf_xlsx_is_image_url`: Validates image extensions
- `nf_xlsx_is_pdf_url`: Checks .pdf extension
- `nf_xlsx_url_extension`: Extracts extension from URL path

#### Excel Helpers (Lines 1418-1457)

- `nf_col_from_index`: Converts 1→A, 27→AA (base-26 conversion)
- `nf_addr`: Creates cell address like "A1", "B2"
- `nf_xlsx_field_identifier`: Creates "field-{id}" identifier

#### Date Formatting (Lines 1405-1416)

- Uses WordPress date/time settings
- `wp_date()` with timezone support
- Fallback: Returns raw string if parsing fails

## Data Flow

### Export Request Flow

```
User clicks "Export to XLSX"
  ↓
nf_xlsx_handle_export (nonce + capability check)
  ↓
Fetch form, fields, submissions from DB
  ↓
nf_xlsx_prepare_columns (normalize selection)
  ↓
new NF_XLSX_Stream_Exporter($form, $columns, $submissions)
  ↓
  initialiseSheets()
    ↓
    addSubmissionHeaders() - Bold headers, freeze pane
    ↓
    populateSubmissions()
      ↓
      For each submission:
        For each column:
          - Extract field payload (text, links, images, PDFs)
          - Write text to cell
          - Add hyperlinks
          - Download & embed images (with caching)
          - Download PDFs, add icon + attachments sheet entry
          - Track offsets for vertical stacking
          - Adjust row heights
  ↓
save($filepath) - Write to uploads directory
  ↓
Disconnect worksheets + cleanup temp files
  ↓
Redirect to admin page with success notice + download link
```

### Submission Field Value Resolution

```
Submission meta array
  ↓
nf_xlsx_candidate_meta_keys() - Try: key, id, field_{id}, _field_{id}
  ↓
nf_xlsx_normalize_meta_values() - Decode JSON/serialized, aggregate
  ↓
nf_xlsx_prepare_value_payload() - Recursive parsing
  ↓
  Is array?
    ↓
    Is upload? (has file_path, url, files keys)
      → Extract URL, label, create payload
    ↓
    Has 'value' key?
      → Recurse on value
    ↓
    Otherwise:
      → Recurse on all elements, combine results
  ↓
  Is scalar?
    → Use as text, detect if URL
  ↓
Return { text, links, images[], pdfs[] }
```

## Key Design Patterns

1. **Caching Strategy**: MD5-keyed cache for images/PDFs prevents duplicate downloads
2. **Recursive Parsing**: Handles deeply nested Ninja Forms field structures
3. **Fallback Chain**: Multiple methods for URL resolution and file access
4. **Resource Cleanup**: Destructor pattern + explicit cleanup prevents temp file leaks
5. **Lazy Initialization**: Attachments sheet only created when needed
6. **Offset Stacking**: Allows multiple images/PDFs vertically in one cell
7. **Two-query Pattern**: Batch-loads submission metadata for performance

## Security Considerations

- **Capability Checks**: All admin actions require `manage_options`
- **Nonce Verification**: Export action uses WordPress nonce
- **Input Sanitization**: `sanitize_text_field`, `absint` on all inputs
- **File Validation**: MIME type and extension checks for uploads
- **SQL Preparation**: All queries use `$wpdb->prepare()`
- **Output Escaping**: `esc_html`, `esc_url`, `esc_attr` in templates

## Performance Optimizations

- **Batch Meta Fetching**: Single query for all submission metadata
- **Image/PDF Caching**: Downloads once per unique URL
- **Local Path Detection**: Avoids HTTP requests for local files
- **AutoSize Columns**: PhpSpreadsheet calculates optimal widths
- **Disconnect Worksheets**: Frees memory after save
- **Limit Preview**: Only shows 5 submissions in admin

## WordPress Integration Points

- **Hooks**: `plugins_loaded`, `admin_menu`, `admin_post_{action}`, `admin_notices`, `register_activation_hook`
- **Functions**: `wp_remote_get`, `wp_upload_dir`, `wp_tempnam`, `wp_date`, `wp_nonce_field`, `add_menu_page`
- **Database**: Direct `$wpdb` access to Ninja Forms custom tables
- **File System**: Uses WordPress upload directory structure
- **Localization**: All strings wrapped in `__()` with `nf-cpt-xlsx-inline` text domain

## File Structure

```
nf-cpt-xlsx-inline/
├── nf-cpt-xlsx-inline.php         Main plugin (1,484 lines)
├── class-nf-xlsx-stream-exporter.php  Excel builder (805 lines)
├── lib/                           Bundled dependencies
│   ├── PhpOffice/PhpSpreadsheet/  Excel library
│   ├── Psr/SimpleCache/           PSR-16 interface
│   ├── Composer/Pcre/             Regex wrapper
│   └── ZipStream/                 ZIP streaming for XLSX
└── README.md                      Documentation
```

## Extension Points

The plugin provides filters for customization:

- `nf_cpt_xlsx_inline_post_types` - Modify available post types
- `nf_cpt_xlsx_inline_headers` - Customize column headers
- `nf_cpt_xlsx_inline_row` - Modify row data before export
- `nf_cpt_xlsx_inline_query_args` - Adjust submission query

(Note: Filter names referenced in README but not implemented in current code)