<?php

use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Style;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class NF_XLSX_Stream_Exporter {
    private const PDF_ICON_BASE64 = 'iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAIAAAAlC+aJAAAAVklEQVR42u3PQQ0AMAzEsOOPrCDGpeOwSe3HUQg4'
        . 'lfzc2wUAAAAAAAAAAAAAAAAAAOANcGp3AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgMkv31CxpiuECMgAAAAASUVORK5CYII=';
    private array $form;
    private array $columns;
    private array $submissions;

    private Spreadsheet $spreadsheet;
    private Worksheet $submissionsSheet;
    private ?Worksheet $attachmentsSheet = null;
    private int $attachmentsRow = 1;

    private array $imageCache = [];
    private array $pdfCache = [];
    private array $tempFiles = [];
    private array $rowHeights = [];
    private array $cellOffsets = [];

    private int $imageCounter = 0;
    private int $pdfCounter = 0;

    private string $submissionsSheetName;
    private string $attachmentsSheetName;

    private ?string $pdfIconPath = null;
    public function __construct(array $form, array $columns, array $submissions) {
        $this->form        = $form;
        $this->columns     = array_values($columns);
        $this->submissions = $submissions;

        $this->submissionsSheetName = self::sanitize_sheet_name(__('Submissions', 'nf-cpt-xlsx-inline'));
        $this->attachmentsSheetName = self::sanitize_sheet_name(__('Attachments', 'nf-cpt-xlsx-inline'));

        $this->initialiseSheets();
    }

    public function __destruct() {
        $this->cleanupTempFiles();
    }

    public function save(string $filepath): void {
        $this->spreadsheet->setActiveSheetIndex(0);
        $writer = new Xlsx($this->spreadsheet);
        $writer->setPreCalculateFormulas(false);
        $writer->save($filepath);
        $this->spreadsheet->disconnectWorksheets();
        $this->cleanupTempFiles();
    }
    private function initialiseSheets(): void {
        $this->spreadsheet = new Spreadsheet();
        $this->spreadsheet->getDefaultStyle()->getFont()->setName('Calibri')->setSize(11);

        $this->submissionsSheet = $this->spreadsheet->getActiveSheet();
        $this->submissionsSheet->setTitle($this->submissionsSheetName);

        $this->addSubmissionHeaders();

        if (empty($this->submissions)) {
            $this->addNoSubmissionsRow();
        } else {
            $this->populateSubmissions();
        }

        $this->submissionsSheet->freezePane('A2');
    }
    private function addSubmissionHeaders(): void {
        foreach ($this->columns as $column) {
            $columnIndex = (int) $column['index'];
            $headerText  = (string) $column['header'];

            $cell = $this->cell($this->submissionsSheet, $columnIndex, 1);
            $cell->setValueExplicit($headerText, DataType::TYPE_STRING);

            $style = $this->style($this->submissionsSheet, $columnIndex, 1);
            $style->getFont()->setBold(true);
            $style->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
            $style->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
            $style->getAlignment()->setWrapText(true);

            $this->submissionsSheet->getColumnDimensionByColumn($columnIndex)->setAutoSize(true);
        }
    }

    private function addNoSubmissionsRow(): void {
        $message = __('No submissions available for this form.', 'nf-cpt-xlsx-inline');
        $cell    = $this->cell($this->submissionsSheet, 1, 2);
        $cell->setValueExplicit($message, DataType::TYPE_STRING);
        $this->submissionsSheet->getStyle('A2')->getAlignment()->setWrapText(true);
        $this->ensureRowHeight(2, 22.0);
    }
    private function populateSubmissions(): void {
        $rowIndex = 2;

        foreach ($this->submissions as $submission) {
            foreach ($this->columns as $column) {
                $columnIndex = (int) $column['index'];

                if ($column['field'] === null) {
                    $dateText = nf_xlsx_format_date($submission['sub_date'] ?? '');
                    if ($dateText !== '') {
                        $this->writeCellText($columnIndex, $rowIndex, $dateText);
                    }
                    continue;
                }

                $payload = nf_xlsx_extract_submission_field_payload($submission, $column['field']);

                if ($payload['text'] !== '') {
                    $this->writeCellText($columnIndex, $rowIndex, (string) $payload['text']);
                }

                if (!empty($payload['links'])) {
                    $this->cell($this->submissionsSheet, $columnIndex, $rowIndex)
                        ->getHyperlink()
                        ->setUrl($payload['links'][0])
                        ->setTooltip(__('Open link', 'nf-cpt-xlsx-inline'));
                }

                $imageUrls = self::image_entries_from_value($payload);
                if ($imageUrls) {
                    foreach ($imageUrls as $imageUrl) {
                        $this->addImage($imageUrl, $rowIndex, $columnIndex);
                    }
                }

                $pdfUrls = self::pdf_entries_from_value($payload);
                if ($pdfUrls) {
                    foreach ($pdfUrls as $pdfUrl) {
                        $this->addPdf($pdfUrl, $rowIndex, $column);
                    }
                }
            }

            ++$rowIndex;
        }
    }
    private function writeCellText(int $columnIndex, int $rowIndex, string $value): void {
        $cell = $this->cell($this->submissionsSheet, $columnIndex, $rowIndex);
        $cell->setValueExplicit($value, DataType::TYPE_STRING);

        $style = $this->style($this->submissionsSheet, $columnIndex, $rowIndex);
        $style->getAlignment()->setWrapText(true);
        $style->getAlignment()->setVertical(Alignment::VERTICAL_TOP);

        $this->ensureRowHeight($rowIndex, 22.0);
    }
    private function addImage(string $url, int $rowIndex, int $columnIndex): void {
        $cacheKey = md5($url);

        if (!array_key_exists($cacheKey, $this->imageCache)) {
            $this->imageCache[$cacheKey] = self::fetch_image_bin($url);
        }

        $imageData = $this->imageCache[$cacheKey];
        if (!$imageData) {
            $this->fallbackLink($url, $columnIndex, $rowIndex);
            return;
        }

        ++$this->imageCounter;

        $maxWidth  = 240.0;
        $maxHeight = 220.0;
        $widthPx   = (float) ($imageData['width'] ?? 120.0);
        $heightPx  = (float) ($imageData['height'] ?? 120.0);

        $scale = min(1.0, $maxWidth / max(1.0, $widthPx), $maxHeight / max(1.0, $heightPx));
        $targetWidth  = max(20.0, $widthPx * $scale);
        $targetHeight = max(20.0, $heightPx * $scale);

        $tempPath = $this->writeTempFile(
            $imageData['data'],
            'nf-image-' . $this->imageCounter . '.' . $imageData['extension']
        );

        if (!$tempPath) {
            $this->fallbackLink($url, $columnIndex, $rowIndex);
            return;
        }

        $drawing = new Drawing();
        $drawing->setName(sprintf(__('Image %d', 'nf-cpt-xlsx-inline'), $this->imageCounter));
        $drawing->setDescription($url);
        $drawing->setPath($tempPath);
        $drawing->setCoordinates($this->coordinate($columnIndex, $rowIndex));
        $drawing->setResizeProportional(true);
        $drawing->setWidth((int) round($targetWidth));
        $drawing->setWorksheet($this->submissionsSheet);

        $offsetY = $this->reserveCellOffset($rowIndex, $columnIndex, $targetHeight);
        $drawing->setOffsetX(4);
        $drawing->setOffsetY($offsetY);
    }
    private function addPdf(string $url, int $rowIndex, array $column): void {
        $columnIndex = (int) $column['index'];
        $cacheKey    = md5($url);

        if (!array_key_exists($cacheKey, $this->pdfCache)) {
            $this->pdfCache[$cacheKey] = self::fetch_pdf_bin($url);
        }

        $pdfBinary = $this->pdfCache[$cacheKey];
        if ($pdfBinary === null) {
            $this->fallbackLink($url, $columnIndex, $rowIndex);
            $this->addAttachmentRow($rowIndex, $column['header'] ?? '', $url, false);
            return;
        }

        ++$this->pdfCounter;

        $pdfTemp = $this->writeTempFile($pdfBinary, 'nf-pdf-' . $this->pdfCounter . '.pdf');
        if (!$pdfTemp) {
            $this->fallbackLink($url, $columnIndex, $rowIndex);
            $this->addAttachmentRow($rowIndex, $column['header'] ?? '', $url, false);
            return;
        }

        $iconPath = $this->getPdfIconPath();
        if (!$iconPath) {
            $this->fallbackLink($url, $columnIndex, $rowIndex);
            $this->addAttachmentRow($rowIndex, $column['header'] ?? '', $url, true);
            return;
        }

        $drawing = new Drawing();
        $drawing->setName(sprintf(__('PDF %d', 'nf-cpt-xlsx-inline'), $this->pdfCounter));
        $drawing->setDescription($url);
        $drawing->setPath($iconPath);
        $drawing->setCoordinates($this->coordinate($columnIndex, $rowIndex));
        $drawing->setResizeProportional(true);
        $drawing->setHeight(22);
        $drawing->setWorksheet($this->submissionsSheet);

        $offsetY = $this->reserveCellOffset($rowIndex, $columnIndex, 22.0);
        $drawing->setOffsetX(2);
        $drawing->setOffsetY($offsetY);

        $cell = $this->cell($this->submissionsSheet, $columnIndex, $rowIndex);
        $cell->getHyperlink()->setUrl($url);
        $cell->getHyperlink()->setTooltip(__('Open PDF', 'nf-cpt-xlsx-inline'));

        $this->addAttachmentRow($rowIndex, $column['header'] ?? '', $url, true);
    }
    private function fallbackLink(string $url, int $columnIndex, int $rowIndex): void {
        if ($url === '') {
            return;
        }

        $cell = $this->cell($this->submissionsSheet, $columnIndex, $rowIndex);
        $existing = (string) $cell->getValue();

        if ($existing !== '') {
            if (strpos($existing, $url) === false) {
                $cell->setValueExplicit($existing . "\n" . $url, DataType::TYPE_STRING);
            }
        } else {
            $cell->setValueExplicit($url, DataType::TYPE_STRING);
        }

        $style = $this->style($this->submissionsSheet, $columnIndex, $rowIndex);
        $style->getAlignment()->setWrapText(true);

        $this->ensureRowHeight($rowIndex, 24.0);
    }
    private function ensureAttachmentsSheet(): Worksheet {
        if ($this->attachmentsSheet instanceof Worksheet) {
            return $this->attachmentsSheet;
        }

        $this->attachmentsSheet = new Worksheet($this->spreadsheet, $this->attachmentsSheetName);
        $this->spreadsheet->addSheet($this->attachmentsSheet);
        $this->attachmentsSheet->setCellValueExplicitByColumnAndRow(1, 1, __('Row', 'nf-cpt-xlsx-inline'), DataType::TYPE_STRING);
        $this->attachmentsSheet->setCellValueExplicitByColumnAndRow(2, 1, __('Column', 'nf-cpt-xlsx-inline'), DataType::TYPE_STRING);
        $this->attachmentsSheet->setCellValueExplicitByColumnAndRow(3, 1, __('Original URL', 'nf-cpt-xlsx-inline'), DataType::TYPE_STRING);
        $this->attachmentsSheet->setCellValueExplicitByColumnAndRow(4, 1, __('Status', 'nf-cpt-xlsx-inline'), DataType::TYPE_STRING);

        for ($col = 1; $col <= 4; $col++) {
            $style = $this->style($this->attachmentsSheet, $col, 1);
            $style->getFont()->setBold(true);
            $style->getAlignment()->setWrapText(true);
            $style->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
        }

        $this->attachmentsSheet->getColumnDimensionByColumn(1)->setWidth(12.0);
        $this->attachmentsSheet->getColumnDimensionByColumn(2)->setWidth(28.0);
        $this->attachmentsSheet->getColumnDimensionByColumn(3)->setWidth(60.0);
        $this->attachmentsSheet->getColumnDimensionByColumn(4)->setWidth(26.0);

        $this->attachmentsRow = 1;

        return $this->attachmentsSheet;
    }

    private function addAttachmentRow(int $sourceRow, string $columnLabel, string $url, bool $downloaded): void {
        $sheet = $this->ensureAttachmentsSheet();
        ++$this->attachmentsRow;

        $sheet->setCellValueExplicitByColumnAndRow(1, $this->attachmentsRow, (string) $sourceRow, DataType::TYPE_STRING);
        $sheet->setCellValueExplicitByColumnAndRow(2, $this->attachmentsRow, (string) $columnLabel, DataType::TYPE_STRING);
        $sheet->setCellValueExplicitByColumnAndRow(3, $this->attachmentsRow, (string) $url, DataType::TYPE_STRING);

        if ($url !== '') {
            $this->cell($sheet, 3, $this->attachmentsRow)
                ->getHyperlink()
                ->setUrl($url)
                ->setTooltip(__('Open original file', 'nf-cpt-xlsx-inline'));
        }

        $status = $downloaded
            ? __('Icon linked', 'nf-cpt-xlsx-inline')
            : __('Download failed', 'nf-cpt-xlsx-inline');

        $sheet->setCellValueExplicitByColumnAndRow(4, $this->attachmentsRow, $status, DataType::TYPE_STRING);

        for ($col = 1; $col <= 4; $col++) {
            $style = $this->style($sheet, $col, $this->attachmentsRow);
            $style->getAlignment()->setWrapText(true);
            $style->getAlignment()->setVertical(Alignment::VERTICAL_TOP);
        }
    }
    private function ensureRowHeight(int $rowIndex, float $heightPx): void {
        $heightPx = max($heightPx, 20.0);
        $current  = $this->rowHeights[$rowIndex] ?? 0.0;

        if ($heightPx > $current) {
            $this->rowHeights[$rowIndex] = $heightPx;
            $points = self::pixels_to_points($heightPx + 4.0);
            $this->submissionsSheet->getRowDimension($rowIndex)->setRowHeight($points);
        }
    }

    private function reserveCellOffset(int $rowIndex, int $columnIndex, float $heightPx): int {
        $key   = $rowIndex . ':' . $columnIndex;
        $state = $this->cellOffsets[$key] ?? ['next' => 2, 'total' => 0];

        $offset = (int) $state['next'];
        $state['next']  = $offset + (int) ceil($heightPx) + 6;
        $state['total'] = max($state['total'], $offset + (int) ceil($heightPx));
        $this->cellOffsets[$key] = $state;

        $this->ensureRowHeight($rowIndex, (float) $state['total'] + 8.0);

        return $offset;
    }
    private function writeTempFile(string $binary, string $filename): ?string {
        if ($binary === '') {
            return null;
        }

        $tempPath = '';

        if (function_exists('wp_tempnam')) {
            $tempPath = wp_tempnam($filename);
        }

        if (!$tempPath) {
            $tempPath = tempnam($this->tempDirectory(), 'nfx');
            if ($tempPath && $filename) {
                $extension = pathinfo($filename, PATHINFO_EXTENSION);
                if ($extension) {
                    $newPath = $tempPath . '.' . $extension;
                    if (@rename($tempPath, $newPath)) {
                        $tempPath = $newPath;
                    }
                }
            }
        }

        if (!$tempPath) {
            return null;
        }

        if (file_put_contents($tempPath, $binary) === false) {
            return null;
        }

        $this->tempFiles[] = $tempPath;

        return $tempPath;
    }

    private function tempDirectory(): string {
        if (function_exists('get_temp_dir')) {
            return get_temp_dir();
        }

        $uploadDir = function_exists('wp_upload_dir') ? wp_upload_dir() : null;
        if (is_array($uploadDir) && empty($uploadDir['error']) && !empty($uploadDir['path'])) {
            return trailingslashit($uploadDir['path']);
        }

        return sys_get_temp_dir();
    }

    private function cleanupTempFiles(): void {
        foreach ($this->tempFiles as $path) {
            if ($path && file_exists($path)) {
                @unlink($path);
            }
        }
        $this->tempFiles = [];
    }

    private function getPdfIconPath(): ?string {
        if ($this->pdfIconPath !== null) {
            return $this->pdfIconPath === '' ? null : $this->pdfIconPath;
        }

        $binary = base64_decode(self::PDF_ICON_BASE64, true);
        if ($binary === false) {
            $this->pdfIconPath = '';
            return null;
        }

        $path = $this->writeTempFile($binary, 'nf-pdf-icon.png');
        if (!$path) {
            $this->pdfIconPath = '';
            return null;
        }

        $this->pdfIconPath = $path;

        return $this->pdfIconPath;
    }

    private function coordinate(int $columnIndex, int $rowIndex): string {
        return Coordinate::stringFromColumnIndex($columnIndex) . (string) $rowIndex;
    }

    private function cell(Worksheet $sheet, int $columnIndex, int $rowIndex): Cell {
        return $sheet->getCell($this->coordinate($columnIndex, $rowIndex));
    }

    private function style(Worksheet $sheet, int $columnIndex, int $rowIndex): Style {
        return $sheet->getStyle($this->coordinate($columnIndex, $rowIndex));
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
            $localPath = self::url_to_local_path($url);
            if ($localPath) {
                $info = @getimagesize($localPath);
                if ($info === false) {
                    return null;
                }
                $body = (string) file_get_contents($localPath);
            } else {
                return null;
            }
        }

        $mime      = isset($info['mime']) ? (string) $info['mime'] : ($response['content_type'] ?? 'image/png');
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
        if ($response['body']) {
            return $response['body'];
        }

        $localPath = self::url_to_local_path($url);
        if ($localPath && file_exists($localPath)) {
            $contents = @file_get_contents($localPath);
            if ($contents !== false) {
                return $contents;
            }
        }

        return null;
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

            if (!is_wp_error($response)) {
                $code = wp_remote_retrieve_response_code($response);
                if ($code < 400) {
                    $body        = (string) wp_remote_retrieve_body($response);
                    $contentType = wp_remote_retrieve_header($response, 'content-type');
                }
            }
        } else {
            $context = stream_context_create([
                'http' => [
                    'timeout' => 15,
                    'header'  => "Accept: image/*,application/pdf;q=0.9,*/*;q=0.1\r\n",
                ],
            ]);

            $fetched = @file_get_contents($url, false, $context);
            if ($fetched !== false) {
                $body = $fetched;
            }

            if (isset($http_response_header) && is_array($http_response_header)) {
                foreach ($http_response_header as $headerLine) {
                    if (stripos($headerLine, 'content-type:') === 0) {
                        $contentType = trim(substr($headerLine, strlen('content-type:')));
                        break;
                    }
                }
            }
        }

        if ($body === '') {
            $local = self::read_local_file($url);
            if ($local) {
                $body        = $local['body'];
                $contentType = $local['content_type'];
            }
        }

        return [
            'body'         => $body,
            'content_type' => $contentType,
        ];
    }

    private static function read_local_file(string $url): ?array {
        $path = self::url_to_local_path($url);
        if (!$path || !file_exists($path)) {
            return null;
        }

        $body = @file_get_contents($path);
        if ($body === false) {
            return null;
        }

        $type = null;
        if (function_exists('wp_check_filetype')) {
            $typeInfo = wp_check_filetype($path);
            if (!empty($typeInfo['type'])) {
                $type = $typeInfo['type'];
            }
        }

        if (!$type && function_exists('mime_content_type')) {
            $type = @mime_content_type($path) ?: null;
        }

        return [
            'body'         => $body,
            'content_type' => $type,
        ];
    }

    private static function url_to_local_path(string $url): string {
        if (!is_string($url) || $url === '') {
            return '';
        }

        $url = trim($url);
        if ($url === '') {
            return '';
        }

        $uploads = function_exists('wp_upload_dir') ? wp_upload_dir() : null;
        if (is_array($uploads) && empty($uploads['error']) && !empty($uploads['baseurl']) && !empty($uploads['basedir'])) {
            $baseUrl = trailingslashit($uploads['baseurl']);
            if (stripos($url, $baseUrl) === 0) {
                $relative = ltrim(substr($url, strlen($baseUrl)), '/');
                $path     = trailingslashit($uploads['basedir']) . str_replace(['\\', '//'], '/', $relative);
                if (file_exists($path)) {
                    return $path;
                }
            }
        }

        $siteUrl = function_exists('site_url') ? trailingslashit(site_url()) : '';
        if ($siteUrl && stripos($url, $siteUrl) === 0) {
            $relative = ltrim(substr($url, strlen($siteUrl)), '/');
            $path     = trailingslashit(ABSPATH) . str_replace(['\\', '//'], '/', $relative);
            if (file_exists($path)) {
                return $path;
            }
        }

        $contentUrl = function_exists('content_url') ? trailingslashit(content_url()) : '';
        if ($contentUrl && stripos($url, $contentUrl) === 0) {
            $relative = ltrim(substr($url, strlen($contentUrl)), '/');
            $path     = trailingslashit(WP_CONTENT_DIR) . str_replace(['\\', '//'], '/', $relative);
            if (file_exists($path)) {
                return $path;
            }
        }

        return '';
    }
    private static function extension_from_url(string $url): string {
        $path = parse_url($url, PHP_URL_PATH);

        if (!$path) {
            return '';
        }

        return strtolower(pathinfo($path, PATHINFO_EXTENSION));
    }

    private static function extension_from_mime(?string $mime): string {
        $mime = is_string($mime) ? strtolower(trim($mime)) : '';

        $map = [
            'image/jpeg'      => 'jpg',
            'image/jpg'       => 'jpg',
            'image/png'       => 'png',
            'image/gif'       => 'gif',
            'image/webp'      => 'webp',
            'application/pdf' => 'pdf',
        ];

        return $map[$mime] ?? '';
    }

    private static function is_image_extension(string $extension): bool {
        $extension = strtolower($extension);

        return in_array($extension, ['jpg', 'jpeg', 'png', 'gif', 'webp'], true);
    }

    private static function is_pdf_extension(string $extension): bool {
        return strtolower($extension) === 'pdf';
    }

    private static function pixels_to_points(float $pixels): float {
        return round($pixels * 72 / 96, 2);
    }

    private static function sanitize_sheet_name(string $name): string {
        $name = preg_replace('/[\\\\\\/*\[\]\?:]/', ' ', $name);
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
}
