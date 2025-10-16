<?php

class NF_XLSX_Stream_Exporter {
    private array $form;
    private array $columns;
    private array $submissions;
    private array $sharedStrings = [];
    private array $sharedLookup = [];
    private array $sheetRows = [
        'submissions' => [],
        'attachments' => [],
    ];
    private array $rowMeta = [
        'submissions' => [],
        'attachments' => [],
    ];
    private array $images = [];
    private array $imageCache = [];
    private array $pdfs = [];
    private array $pdfCache = [];
    private int $imageCounter = 0;
    private int $pdfCounter = 0;
    private int $attachmentsRowIndex = 1;
    private bool $attachmentsSheetInitialized = false;
    private string $submissionsSheetName;
    private string $attachmentsSheetName;

    public function __construct(array $form, array $columns, array $submissions) {
        $this->form        = $form;
        $this->columns     = array_values($columns);
        $this->submissions = $submissions;

        $this->submissionsSheetName = self::sanitize_sheet_name(__('Submissions', 'nf-cpt-xlsx-inline'));
        $this->attachmentsSheetName = self::sanitize_sheet_name(__('Attachments', 'nf-cpt-xlsx-inline'));

        $this->initialise_sheets();
    }

    private function initialise_sheets(): void {
        $this->add_submission_headers();
        if (empty($this->submissions)) {
            $this->add_no_submissions_row();
        } else {
            $this->populate_submissions();
        }
    }

    private function add_submission_headers(): void {
        foreach ($this->columns as $column) {
            $this->add_cell('submissions', 1, (int) $column['index'], (string) $column['header']);
        }
    }

    private function add_attachments_header(): void {
        $headers = [
            __('Row', 'nf-cpt-xlsx-inline'),
            __('Column', 'nf-cpt-xlsx-inline'),
            __('Original URL', 'nf-cpt-xlsx-inline'),
            __('Embedded Part', 'nf-cpt-xlsx-inline'),
        ];

        $columnIndex = 1;
        foreach ($headers as $header) {
            $this->add_cell('attachments', 1, $columnIndex, (string) $header);
            ++$columnIndex;
        }
    }

    private function add_no_submissions_row(): void {
        $message = __('No submissions available for this form.', 'nf-cpt-xlsx-inline');
        $this->add_cell('submissions', 2, 1, $message);
        $this->rowMeta['submissions'][2]['max'] = max(1, $this->rowMeta['submissions'][2]['max'] ?? 1);
    }

    private function populate_submissions(): void {
        $rowIndex = 2;

        foreach ($this->submissions as $submission) {
            foreach ($this->columns as $column) {
                $columnIndex = (int) $column['index'];

                if ($column['field'] === null) {
                    $dateText = nf_xlsx_format_date($submission['sub_date']);
                    if ($dateText !== '') {
                        $this->add_cell('submissions', $rowIndex, $columnIndex, $dateText);
                    }
                    continue;
                }

                $payload = nf_xlsx_extract_submission_field_payload($submission, $column['field']);

                if ($payload['text'] !== '') {
                    $this->add_cell('submissions', $rowIndex, $columnIndex, $payload['text']);
                }

                $imageUrls = self::image_entries_from_value($payload);
                foreach ($imageUrls as $imageUrl) {
                    $this->add_image($imageUrl, $rowIndex, $columnIndex);
                }

                $pdfUrls = self::pdf_entries_from_value($payload);
                foreach ($pdfUrls as $pdfUrl) {
                    $this->add_pdf($pdfUrl, $rowIndex, $column);
                }
            }

            ++$rowIndex;
        }
    }

    private function add_image(string $url, int $rowIndex, int $columnIndex): void {
        $cacheKey = md5($url);

        if (!array_key_exists($cacheKey, $this->imageCache)) {
            $this->imageCache[$cacheKey] = self::fetch_image_bin($url);
        }

        $imageData = $this->imageCache[$cacheKey];
        if (!$imageData) {
            return;
        }

        ++$this->imageCounter;

        $fileName = 'image' . $this->imageCounter . '.' . $imageData['extension'];
        $widthPx  = (float) ($imageData['width'] ?? 120);
        $heightPx = (float) ($imageData['height'] ?? 120);
        $heightPt = max(60.0, self::pixels_to_points($heightPx));

        $this->rowMeta['submissions'][$rowIndex]['height'] = max(
            $heightPt,
            $this->rowMeta['submissions'][$rowIndex]['height'] ?? 0.0
        );

        $this->images[] = [
            'path'       => 'xl/media/' . $fileName,
            'data'       => $imageData['data'],
            'mime'       => $imageData['mime'],
            'extension'  => $imageData['extension'],
            'row'        => $rowIndex,
            'column'     => $columnIndex,
            'width_emu'  => self::pixels_to_emu($widthPx),
            'height_emu' => self::pixels_to_emu($heightPx),
        ];
    }

    private function add_pdf(string $url, int $rowIndex, array $column): void {
        $cacheKey = md5($url);

        if (!array_key_exists($cacheKey, $this->pdfCache)) {
            $this->pdfCache[$cacheKey] = self::fetch_pdf_bin($url);
        }

        $pdfBinary = $this->pdfCache[$cacheKey];
        if ($pdfBinary === null) {
            return;
        }

        $this->ensure_attachments_sheet();

        ++$this->pdfCounter;

        $fileName = 'file' . $this->pdfCounter . '.pdf';
        $partName = 'xl/embeddings/' . $fileName;

        $this->pdfs[] = [
            'path' => $partName,
            'data' => $pdfBinary,
        ];

        $columnLabel = isset($column['header']) ? (string) $column['header'] : '';
        $this->add_attachment_row($rowIndex, $columnLabel, $url, $partName);
    }

    private function add_attachment_row(int $sourceRow, string $columnLabel, string $url, string $partName): void {
        ++$this->attachmentsRowIndex;
        $rowIndex = $this->attachmentsRowIndex;

        $this->add_cell('attachments', $rowIndex, 1, (string) $sourceRow);
        $this->add_cell('attachments', $rowIndex, 2, $columnLabel);
        $this->add_cell('attachments', $rowIndex, 3, $url);
        $this->add_cell('attachments', $rowIndex, 4, $partName);
    }

    private function ensure_attachments_sheet(): void {
        if ($this->attachmentsSheetInitialized) {
            return;
        }

        $this->attachmentsSheetInitialized = true;
        $this->attachmentsRowIndex         = 1;
        $this->sheetRows['attachments']    = [];
        $this->rowMeta['attachments']      = [];

        $this->add_attachments_header();
    }

    private function add_cell(string $sheet, int $rowIndex, int $columnIndex, string $value): void {
        $value = trim((string) $value);

        $this->ensure_row($sheet, $rowIndex);

        if ($value === '') {
            $this->rowMeta[$sheet][$rowIndex]['max'] = max(
                $columnIndex,
                $this->rowMeta[$sheet][$rowIndex]['max'] ?? 0
            );
            return;
        }

        $sharedIndex = $this->register_shared_string($value);

        $this->sheetRows[$sheet][$rowIndex][] = [
            'column' => $columnIndex,
            'shared' => $sharedIndex,
        ];

        $this->rowMeta[$sheet][$rowIndex]['max'] = max(
            $columnIndex,
            $this->rowMeta[$sheet][$rowIndex]['max'] ?? 0
        );
    }

    private function ensure_row(string $sheet, int $rowIndex): void {
        if (!isset($this->sheetRows[$sheet][$rowIndex])) {
            $this->sheetRows[$sheet][$rowIndex] = [];
        }

        if (!isset($this->rowMeta[$sheet][$rowIndex])) {
            $this->rowMeta[$sheet][$rowIndex] = [
                'max'    => 0,
                'height' => 0,
            ];
        }
    }

    private function register_shared_string(string $value): int {
        $value = self::normalise_string($value);

        if (isset($this->sharedLookup[$value])) {
            return $this->sharedLookup[$value];
        }

        $index = count($this->sharedStrings);
        $this->sharedStrings[$index] = $value;
        $this->sharedLookup[$value]  = $index;

        return $index;
    }

    public function save(string $filepath): void {
        if ($this->images) {
            $relationIndex = 1;

            foreach ($this->images as $index => $image) {
                $this->images[$index]['rid']      = 'rId' . $relationIndex;
                $this->images[$index]['shape_id'] = $relationIndex;
                ++$relationIndex;
            }
        }

        $sharedStringsXml = $this->build_shared_strings_xml();
        $sheet1Xml        = $this->build_submissions_sheet_xml();
        $sheet2Xml        = $this->attachmentsSheetInitialized ? $this->build_attachments_sheet_xml() : '';
        $workbookXml      = $this->build_workbook_xml();
        $workbookRelsXml  = $this->build_workbook_rels_xml();
        $contentTypesXml  = $this->build_content_types_xml();
        $packageRelsXml   = $this->build_package_rels_xml();
        $stylesXml        = $this->build_styles_xml();
        $themeXml         = $this->build_theme_xml();
        $coreXml          = $this->build_docprops_core_xml();
        $appXml           = $this->build_docprops_app_xml();

        $zip = new ZipArchive();
        $openResult = $zip->open($filepath, ZipArchive::OVERWRITE | ZipArchive::CREATE);

        if ($openResult !== true) {
            throw new RuntimeException(__('Unable to create workbook archive.', 'nf-cpt-xlsx-inline'));
        }

        $zip->addFromString('[Content_Types].xml', $contentTypesXml);
        $zip->addFromString('_rels/.rels', $packageRelsXml);
        $zip->addFromString('docProps/core.xml', $coreXml);
        $zip->addFromString('docProps/app.xml', $appXml);
        $zip->addFromString('xl/workbook.xml', $workbookXml);
        $zip->addFromString('xl/_rels/workbook.xml.rels', $workbookRelsXml);
        $zip->addFromString('xl/styles.xml', $stylesXml);
        $zip->addFromString('xl/theme/theme1.xml', $themeXml);
        $zip->addFromString('xl/sharedStrings.xml', $sharedStringsXml);
        $zip->addFromString('xl/worksheets/sheet1.xml', $sheet1Xml);

        if ($this->attachmentsSheetInitialized) {
            $zip->addFromString('xl/worksheets/sheet2.xml', $sheet2Xml);
        }

        if ($this->images) {
            $zip->addFromString('xl/worksheets/_rels/sheet1.xml.rels', $this->build_sheet1_rels_xml());
            $zip->addFromString('xl/drawings/drawing1.xml', $this->build_drawing_xml());
            $zip->addFromString('xl/drawings/_rels/drawing1.xml.rels', $this->build_drawing_rels_xml());

            foreach ($this->images as $image) {
                $zip->addFromString($image['path'], $image['data']);
            }
        }

        foreach ($this->pdfs as $pdf) {
            $zip->addFromString($pdf['path'], $pdf['data']);
        }

        $zip->close();
    }

    private function build_submissions_sheet_xml(): string {
        $dimension  = $this->build_dimension_for_sheet('submissions', count($this->columns));
        $sheetViews = '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>';
        $sheetData  = $this->build_sheet_data_xml('submissions');
        $cols       = $this->build_submission_cols_xml();
        $drawingTag = $this->images ? '<drawing r:id="rId1"/>' : '';

        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            . '<dimension ref="' . $dimension . '"/>'
            . $sheetViews
            . '<sheetFormatPr defaultRowHeight="15"/>'
            . $cols
            . $sheetData
            . '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
            . $drawingTag
            . '</worksheet>';
    }

    private function build_attachments_sheet_xml(): string {
        $dimension  = $this->build_dimension_for_sheet('attachments', 4);
        $sheetViews = '<sheetViews><sheetView workbookViewId="0"/></sheetViews>';
        $sheetData  = $this->build_sheet_data_xml('attachments');
        $cols       = $this->build_attachments_cols_xml();

        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            . '<dimension ref="' . $dimension . '"/>'
            . $sheetViews
            . '<sheetFormatPr defaultRowHeight="15"/>'
            . $cols
            . $sheetData
            . '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
            . '</worksheet>';
    }

    private function build_sheet_data_xml(string $sheet): string {
        $rows = $this->sheetRows[$sheet];

        if (!$rows) {
            return '<sheetData/>';
        }

        ksort($rows);

        $xml = '<sheetData>';

        foreach ($rows as $rowIndex => $cells) {
            usort($cells, static function ($a, $b) {
                return $a['column'] <=> $b['column'];
            });

            $fallbackMax = $sheet === 'submissions' ? max(1, count($this->columns)) : 4;
            $metaMax     = $this->rowMeta[$sheet][$rowIndex]['max'] ?? 0;
            $maxCol      = max(1, $metaMax ? $metaMax : $fallbackMax);
            $height   = (float) ($this->rowMeta[$sheet][$rowIndex]['height'] ?? 0);
            $rowAttrs = sprintf('r="%d" spans="1:%d"', $rowIndex, $maxCol);

            if ($height > 0) {
                $rowAttrs .= sprintf(' ht="%.2F" customHeight="1"', $height);
            }

            $xml .= '<row ' . $rowAttrs . '>';

            foreach ($cells as $cell) {
                $coordinate = nf_addr($cell['column'], $rowIndex);
                $xml       .= '<c r="' . $coordinate . '" t="s"><v>' . $cell['shared'] . '</v></c>';
            }

            $xml .= '</row>';
        }

        $xml .= '</sheetData>';

        return $xml;
    }

    private function build_submission_cols_xml(): string {
        if (!$this->columns) {
            return '';
        }

        $cols = '<cols>';

        foreach ($this->columns as $column) {
            $width = $column['field'] === null ? 22.0 : 30.0;
            $index = (int) $column['index'];

            $cols .= sprintf('<col min="%d" max="%d" width="%.2F" customWidth="1"/>', $index, $index, $width);
        }

        $cols .= '</cols>';

        return $cols;
    }

    private function build_attachments_cols_xml(): string {
        $widths = [12.0, 28.0, 60.0, 32.0];
        $cols   = '<cols>';

        foreach ($widths as $index => $width) {
            $colIndex = $index + 1;
            $cols    .= sprintf('<col min="%d" max="%d" width="%.2F" customWidth="1"/>', $colIndex, $colIndex, $width);
        }

        $cols .= '</cols>';

        return $cols;
    }

    private function build_dimension_for_sheet(string $sheet, int $fallbackColumns): string {
        $rows = $this->sheetRows[$sheet];

        if (!$rows) {
            return 'A1:' . nf_addr(max(1, $fallbackColumns), 1);
        }

        $maxRow = max(array_keys($rows));
        $maxCol = $fallbackColumns;

        foreach ($this->rowMeta[$sheet] as $meta) {
            if (!empty($meta['max']) && $meta['max'] > $maxCol) {
                $maxCol = (int) $meta['max'];
            }
        }

        return 'A1:' . nf_addr(max(1, $maxCol), max(1, $maxRow));
    }

    private function build_shared_strings_xml(): string {
        $uniqueCount = count($this->sharedStrings);
        $count       = $uniqueCount;
        $strings     = '';

        foreach ($this->sharedStrings as $string) {
            $escaped = self::xml_escape($string);

            if ($escaped === '') {
                $strings .= '<si><t/></si>';
                continue;
            }

            if (preg_match('/^\s|\s$/', $string)) {
                $strings .= '<si><t xml:space="preserve">' . $escaped . '</t></si>';
            } else {
                $strings .= '<si><t>' . $escaped . '</t></si>';
            }
        }

        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' . $count . '" uniqueCount="' . $uniqueCount . '">' . $strings . '</sst>';
    }

    private function build_workbook_xml(): string {
        $sheetIndex = 1;
        $sheetEntries = [];

        $sheetEntries[] = '<sheet name="' . self::xml_escape($this->submissionsSheetName) . '" sheetId="' . $sheetIndex . '" r:id="rId' . $sheetIndex . '"/>';

        if ($this->attachmentsSheetInitialized) {
            $sheetIndex++;
            $sheetEntries[] = '<sheet name="' . self::xml_escape($this->attachmentsSheetName) . '" sheetId="' . $sheetIndex . '" r:id="rId' . $sheetIndex . '"/>';
        }

        $sheets = '<sheets>' . implode('', $sheetEntries) . '</sheets>';

        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            . '<workbookPr date1904="false"/>'
            . '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="14400"/></bookViews>'
            . $sheets
            . '</workbook>';
    }

    private function build_workbook_rels_xml(): string {
        $relations = [
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        ];

        if ($this->attachmentsSheetInitialized) {
            $relations[] = '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>';
        }

        $relations[] = '<Relationship Id="rIdSharedStrings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>';
        $relations[] = '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';
        $relations[] = '<Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>';

        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            . implode('', $relations)
            . '</Relationships>';
    }

    private function build_package_rels_xml(): string {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            . '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            . '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
            . '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
            . '</Relationships>';
    }

    private function build_content_types_xml(): string {
        $overrides = [
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
            '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>',
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
            '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>',
            '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
            '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
        ];

        if ($this->attachmentsSheetInitialized) {
            $overrides[] = '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
        }

        if ($this->images) {
            $overrides[] = '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>';
        }

        foreach ($this->pdfs as $index => $pdf) {
            $overrides[] = '<Override PartName="/' . self::xml_escape($pdf['path']) . '" ContentType="application/pdf"/>';
        }

        $defaults = [
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
            '<Default Extension="xml" ContentType="application/xml"/>',
            '<Default Extension="png" ContentType="image/png"/>',
            '<Default Extension="jpeg" ContentType="image/jpeg"/>',
            '<Default Extension="jpg" ContentType="image/jpeg"/>',
            '<Default Extension="gif" ContentType="image/gif"/>',
            '<Default Extension="webp" ContentType="image/webp"/>',
            '<Default Extension="pdf" ContentType="application/pdf"/>',
        ];

        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            . implode('', $defaults)
            . implode('', $overrides)
            . '</Types>';
    }

    private function build_styles_xml(): string {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            . '<fonts count="1"><font><name val="Calibri"/><sz val="11"/></font></fonts>'
            . '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
            . '<borders count="1"><border/></borders>'
            . '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            . '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
            . '</styleSheet>';
    }

    private function build_theme_xml(): string {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">'
            . '<a:themeElements>'
            . '<a:clrScheme name="Office">'
            . '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
            . '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
            . '<a:dk2><a:srgbClr val="1F497D"/></a:dk2>'
            . '<a:lt2><a:srgbClr val="EEECE1"/></a:lt2>'
            . '</a:clrScheme>'
            . '</a:themeElements>'
            . '</a:theme>';
    }

    private function build_sheet1_rels_xml(): string {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            . '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>'
            . '</Relationships>';
    }

    private function build_drawing_xml(): string {
        $anchors = '';

        foreach ($this->images as $image) {
            $col = max(0, (int) $image['column'] - 1);
            $row = max(0, (int) $image['row'] - 1);

            $anchors .= '<xdr:oneCellAnchor>'
                . '<xdr:from>'
                . '<xdr:col>' . $col . '</xdr:col>'
                . '<xdr:colOff>0</xdr:colOff>'
                . '<xdr:row>' . $row . '</xdr:row>'
                . '<xdr:rowOff>0</xdr:rowOff>'
                . '</xdr:from>'
                . '<xdr:ext cx="' . $image['width_emu'] . '" cy="' . $image['height_emu'] . '"/>'
                . '<xdr:pic>'
                . '<xdr:nvPicPr>'
                . '<xdr:cNvPr id="' . $image['shape_id'] . '" name="Picture ' . $image['shape_id'] . '"/>'
                . '<xdr:cNvPicPr/>'
                . '</xdr:nvPicPr>'
                . '<xdr:blipFill>'
                . '<a:blip r:embed="' . $image['rid'] . '"/>'
                . '<a:stretch><a:fillRect/></a:stretch>'
                . '</xdr:blipFill>'
                . '<xdr:spPr>'
                . '<a:xfrm><a:off x="0" y="0"/><a:ext cx="' . $image['width_emu'] . '" cy="' . $image['height_emu'] . '"/></a:xfrm>'
                . '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
                . '</xdr:spPr>'
                . '</xdr:pic>'
                . '<xdr:clientData/>'
                . '</xdr:oneCellAnchor>';
        }

        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            . $anchors
            . '</xdr:wsDr>';
    }

    private function build_drawing_rels_xml(): string {
        $relationships = '';

        foreach ($this->images as $image) {
            $relationships .= '<Relationship Id="' . $image['rid'] . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/' . basename($image['path']) . '"/>';
        }

        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            . $relationships
            . '</Relationships>';
    }

    private function build_docprops_core_xml(): string {
        $title    = isset($this->form['title']) ? (string) $this->form['title'] : '';
        $blogName = get_bloginfo('name');
        $creator  = $blogName ? (string) $blogName : 'NF CPT → XLSX Inline Export';
        $now      = self::w3c_date();

        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            . '<dc:title>' . self::xml_escape(sprintf(__('NF Submissions Export – %s', 'nf-cpt-xlsx-inline'), $title)) . '</dc:title>'
            . '<dc:creator>' . self::xml_escape($creator) . '</dc:creator>'
            . '<cp:lastModifiedBy>NF CPT → XLSX Inline Export</cp:lastModifiedBy>'
            . '<dcterms:created xsi:type="dcterms:W3CDTF">' . $now . '</dcterms:created>'
            . '<dcterms:modified xsi:type="dcterms:W3CDTF">' . $now . '</dcterms:modified>'
            . '</cp:coreProperties>';
    }

    private function build_docprops_app_xml(): string {
        $sheetNames = [$this->submissionsSheetName];

        if ($this->attachmentsSheetInitialized) {
            $sheetNames[] = $this->attachmentsSheetName;
        }

        $sheetCount = count($sheetNames);
        $titlesVector = '<TitlesOfParts><vt:vector size="' . $sheetCount . '" baseType="lpstr">';

        foreach ($sheetNames as $name) {
            $titlesVector .= '<vt:lpstr>' . self::xml_escape($name) . '</vt:lpstr>';
        }

        $titlesVector .= '</vt:vector></TitlesOfParts>';

        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            . '<Application>NF CPT → XLSX Inline Export</Application>'
            . '<DocSecurity>0</DocSecurity>'
            . '<ScaleCrop>false</ScaleCrop>'
            . '<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>' . $sheetCount . '</vt:i4></vt:variant></vt:vector></HeadingPairs>'
            . $titlesVector
            . '<Company>' . self::xml_escape(get_bloginfo('name')) . '</Company>'
            . '<LinksUpToDate>false</LinksUpToDate>'
            . '<SharedDoc>false</SharedDoc>'
            . '<HyperlinksChanged>false</HyperlinksChanged>'
            . '<AppVersion>16.0300</AppVersion>'
            . '</Properties>';
    }

    private static function pixels_to_points(float $pixels): float {
        return round($pixels * 72 / 96, 2);
    }

    private static function pixels_to_emu(float $pixels): int {
        return (int) round($pixels * 9525);
    }

    private static function normalise_string(string $value): string {
        $value = str_replace(["\r\n", "\r"], "\n", $value);

        return $value;
    }

    private static function xml_escape(string $value): string {
        return htmlspecialchars($value, ENT_XML1 | ENT_COMPAT, 'UTF-8');
    }

    private static function sanitize_sheet_name(string $name): string {
        $name = preg_replace('/[\\\\\/*\[\]\?:]/', ' ', $name);
        $name = trim((string) $name);

        if ($name === '') {
            $name = 'Sheet';
        }

        if (function_exists('mb_substr')) {
            $name = mb_substr($name, 0, 31, 'UTF-8');
        } else {
            $name = substr($name, 0, 31);
        }

        return $name;
    }

    private static function w3c_date(): string {
        return gmdate('Y-m-d\TH:i:s\Z');
    }

    private static function image_entries_from_value(array $payload): array {
        $urls = [];

        if (!empty($payload['images'])) {
            $urls = array_merge($urls, (array) $payload['images']);
        }

        if (!empty($payload['links'])) {
            foreach ((array) $payload['links'] as $link) {
                if (self::is_image_extension(self::extension_from_url($link))) {
                    $urls[] = $link;
                }
            }
        }

        if (!empty($payload['text'])) {
            foreach (self::extract_urls_from_text($payload['text']) as $link) {
                if (self::is_image_extension(self::extension_from_url($link))) {
                    $urls[] = $link;
                }
            }
        }

        $urls = array_values(array_filter(array_unique($urls)));

        return $urls;
    }

    private static function pdf_entries_from_value(array $payload): array {
        $urls = [];

        if (!empty($payload['pdfs'])) {
            $urls = array_merge($urls, (array) $payload['pdfs']);
        }

        if (!empty($payload['links'])) {
            foreach ((array) $payload['links'] as $link) {
                if (self::is_pdf_extension(self::extension_from_url($link))) {
                    $urls[] = $link;
                }
            }
        }

        if (!empty($payload['text'])) {
            foreach (self::extract_urls_from_text($payload['text']) as $link) {
                if (self::is_pdf_extension(self::extension_from_url($link))) {
                    $urls[] = $link;
                }
            }
        }

        return array_values(array_filter(array_unique($urls)));
    }

    private static function extract_urls_from_text(string $text): array {
        $pattern = '/https?:\/\/[^\s]+/i';
        preg_match_all($pattern, $text, $matches);

        if (empty($matches[0])) {
            return [];
        }

        return array_values(array_unique($matches[0]));
    }

    private static function fetch_image_bin(string $url): ?array {
        if ($url === '') {
            return null;
        }

        $response = self::perform_http_request($url);

        if (!$response['body']) {
            return null;
        }

        $body = $response['body'];
        $info = @getimagesizefromstring($body);

        if ($info === false) {
            return null;
        }

        $mime      = isset($info['mime']) ? $info['mime'] : ($response['content_type'] ?? 'image/png');
        $extension = self::extension_from_mime($mime);

        if ($extension === '') {
            $extension = self::extension_from_url($url);
        }

        if ($extension === '') {
            $extension = 'png';
        }

        return [
            'data'      => $body,
            'mime'      => $mime,
            'extension' => $extension,
            'width'     => isset($info[0]) ? (int) $info[0] : 120,
            'height'    => isset($info[1]) ? (int) $info[1] : 120,
        ];
    }

    private static function fetch_pdf_bin(string $url): ?string {
        if ($url === '') {
            return null;
        }

        $response = self::perform_http_request($url);

        if (!$response['body']) {
            return null;
        }

        return $response['body'];
    }

    private static function perform_http_request(string $url): array {
        $body        = '';
        $contentType = null;

        if (function_exists('wp_remote_get')) {
            $response = wp_remote_get($url, [
                'timeout' => 15,
                'headers' => [
                    'Accept' => 'image/*,application/pdf;q=0.9,*/*;q=0.1',
                ],
            ]);

            if (is_wp_error($response)) {
                return ['body' => '', 'content_type' => null];
            }

            $code = wp_remote_retrieve_response_code($response);
            if ($code >= 400) {
                return ['body' => '', 'content_type' => null];
            }

            $body        = (string) wp_remote_retrieve_body($response);
            $contentType = wp_remote_retrieve_header($response, 'content-type');
        } else {
            $context = stream_context_create([
                'http' => [
                    'timeout' => 15,
                    'header'  => "Accept: image/*,application/pdf;q=0.9,*/*;q=0.1\r\n",
                ],
            ]);

            $body = @file_get_contents($url, false, $context);

            if ($body === false) {
                $body = '';
            }

            if (isset($http_response_header)) {
                foreach ($http_response_header as $headerLine) {
                    if (stripos($headerLine, 'content-type:') === 0) {
                        $contentType = trim(substr($headerLine, strlen('content-type:')));
                        break;
                    }
                }
            }
        }

        return [
            'body'        => $body,
            'content_type' => $contentType,
        ];
    }

    private static function extension_from_url(string $url): string {
        $path = parse_url($url, PHP_URL_PATH);

        if (!$path) {
            return '';
        }

        $extension = strtolower(pathinfo($path, PATHINFO_EXTENSION));

        return $extension;
    }

    private static function extension_from_mime(?string $mime): string {
        $mime = is_string($mime) ? strtolower(trim($mime)) : '';

        $map = [
            'image/jpeg' => 'jpg',
            'image/jpg'  => 'jpg',
            'image/png'  => 'png',
            'image/gif'  => 'gif',
            'image/webp' => 'webp',
            'application/pdf' => 'pdf',
        ];

        return $map[$mime] ?? '';
    }

    private static function mime_from_extension(string $extension): string {
        $map = [
            'jpg'  => 'image/jpeg',
            'jpeg' => 'image/jpeg',
            'png'  => 'image/png',
            'gif'  => 'image/gif',
            'pdf'  => 'application/pdf',
        ];

        $extension = strtolower($extension);

        return $map[$extension] ?? 'application/octet-stream';
    }

    private static function is_image_extension(string $extension): bool {
        $extension = strtolower($extension);

        return in_array($extension, ['jpg', 'jpeg', 'png', 'gif', 'webp'], true);
    }

    private static function is_pdf_extension(string $extension): bool {
        return strtolower($extension) === 'pdf';
    }
}
