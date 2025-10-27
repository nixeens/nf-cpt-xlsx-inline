<?php

namespace CodexV2;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use RuntimeException;
use ZipArchive;

class NF_XLSX_Stream_Exporter
{
    private const PDF_ICON_BASE64 = 'iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAIAAAAlC+aJAAAAVklEQVR42u3PQQ0AMAzEsOOPrCDGpeOwSe3HUQg4'
        . 'lfzc2wUAAAAAAAAAAAAAAAAAAOANcGp3AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgMkv31CxpiuECMgAAAAASUVORK5CYII=';

    private array $headers;
    private array $rows;
    private array $options;

    private Spreadsheet $spreadsheet;
    private Worksheet $dataSheet;
    private ?Worksheet $attachmentsSheet = null;

    private int $attachmentsSheetRow = 1;
    private array $rowOffsets = [];
    private array $rowHeights = [];
    private array $attachments = [];
    private array $tempFiles = [];

    private int $imageIndex = 0;
    private int $pdfIndex = 0;

    private ?string $pdfIconPath = null;

    private int $imageColumnIndex;
    private int $signatureColumnIndex;
    private int $pdfColumnIndex;

    private const IMAGE_MAX_WIDTH = 170.0;
    private const IMAGE_MAX_HEIGHT = 220.0;

    public function __construct(array $headers, array $rows, array $options = [])
    {
        $this->headers = $headers;
        $this->rows = $rows;
        $this->options = array_merge(
            [
                'include_attachments' => true,
                'treat_signatures_separately' => true,
            ],
            $options
        );

        $this->buildWorkbook();
    }

    public function __destruct()
    {
        $this->cleanupTempFiles();
    }

    public function exportZip(string $zipPath): array
    {
        $xlsxTemp = $this->createTempFile('export.xlsx');
        $this->saveSpreadsheet($xlsxTemp);

        $zip = new ZipArchive();
        if ($zip->open($zipPath, ZipArchive::CREATE | ZipArchive::OVERWRITE) !== true) {
            throw new RuntimeException('Unable to create ZIP archive at ' . $zipPath);
        }

        if (!$zip->addFile($xlsxTemp, 'export.xlsx')) {
            throw new RuntimeException('Unable to add workbook to archive.');
        }

        if (!empty($this->options['include_attachments'])) {
            foreach ($this->attachments as $attachment) {
                $path = $attachment['path'] ?? '';
                if ($path && file_exists($path)) {
                    $zip->addFile($path, $attachment['zip_name']);
                }
            }
        }

        $manifest = $this->createManifest();
        $json = function_exists('wp_json_encode')
            ? wp_json_encode($manifest, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES)
            : json_encode($manifest, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES);

        $zip->addFromString('manifest.json', (string) $json);
        $zip->close();

        return $manifest;
    }

    private function buildWorkbook(): void
    {
        $this->spreadsheet = new Spreadsheet();
        $this->configureTheme();
        $this->spreadsheet->getDefaultStyle()->getFont()->setName('Calibri')->setSize(11);

        $this->dataSheet = $this->spreadsheet->getActiveSheet();
        $this->dataSheet->setTitle('Data');

        $dataColumns = count($this->headers);
        $this->imageColumnIndex = $dataColumns + 1;
        $this->signatureColumnIndex = $dataColumns + 2;
        $this->pdfColumnIndex = $dataColumns + 3;

        $this->addHeaders();
        $this->addRows();

        $this->dataSheet->freezePane('A2');
    }

    private function addHeaders(): void
    {
        $columnIndex = 1;
        foreach ($this->headers as $header) {
            $label = isset($header['label']) ? (string) $header['label'] : '';
            $this->setCellValue($this->dataSheet, $columnIndex, 1, $label);
            $style = $this->dataSheet->getStyle(Coordinate::stringFromColumnIndex($columnIndex) . '1');
            $style->getFont()->setBold(true);
            $style->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
            $style->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
            $style->getAlignment()->setWrapText(true);
            $this->dataSheet->getColumnDimensionByColumn($columnIndex)->setAutoSize(true);
            ++$columnIndex;
        }

        $this->setAttachmentHeader($this->imageColumnIndex, __('Images', 'nf-cpt-xlsx-inline'));
        $this->setAttachmentHeader($this->signatureColumnIndex, __('Signatures', 'nf-cpt-xlsx-inline'));
        $this->setAttachmentHeader($this->pdfColumnIndex, __('PDFs', 'nf-cpt-xlsx-inline'));
    }

    private function setAttachmentHeader(int $columnIndex, string $label): void
    {
        $this->setCellValue($this->dataSheet, $columnIndex, 1, $label);
        $style = $this->dataSheet->getStyle(Coordinate::stringFromColumnIndex($columnIndex) . '1');
        $style->getFont()->setBold(true);
        $style->getAlignment()->setWrapText(true);
        $style->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
        $this->dataSheet->getColumnDimensionByColumn($columnIndex)->setWidth(32.0);
    }

    private function addRows(): void
    {
        $rowNumber = 2;
        foreach ($this->rows as $row) {
            $this->addRow($row, $rowNumber);
            ++$rowNumber;
        }
    }

    private function addRow(array $row, int $rowNumber): void
    {
        foreach ($this->headers as $index => $header) {
            $key = $header['key'] ?? '';
            $value = '';
            if ($key !== '' && isset($row['values'][$key])) {
                $value = (string) $row['values'][$key];
            }

            $this->setCellValue($this->dataSheet, $index + 1, $rowNumber, $value);
            $style = $this->dataSheet->getStyle(Coordinate::stringFromColumnIndex($index + 1) . $rowNumber);
            $style->getAlignment()->setWrapText(true);
            $style->getAlignment()->setVertical(Alignment::VERTICAL_TOP);
        }

        $this->ensureRowHeight($rowNumber, 24.0);

        if (!empty($row['attachments']) && is_array($row['attachments'])) {
            foreach ($row['attachments'] as $attachment) {
                $this->attachments[] = $attachment;
                $this->processAttachment($attachment, $rowNumber);
            }
        }
    }

    private function processAttachment(array $attachment, int $rowNumber): void
    {
        $type = $attachment['type'] ?? 'other';

        if ($type === 'image') {
            $isSignature = !empty($attachment['is_signature']) && !empty($this->options['treat_signatures_separately']);
            $columnIndex = $isSignature ? $this->signatureColumnIndex : $this->imageColumnIndex;
            $this->embedImage($attachment, $rowNumber, $columnIndex);
        } elseif ($type === 'pdf') {
            $this->embedPdf($attachment, $rowNumber);
        } else {
            $this->addAttachmentSheetRow($attachment);
        }
    }

    private function embedImage(array $attachment, int $rowNumber, int $columnIndex): void
    {
        $resolved = $this->resolveImagePath($attachment);
        if ($resolved === null) {
            return;
        }

        $path = $resolved['path'];
        $width = $resolved['width'];
        $height = $resolved['height'];

        $scale = min(1.0, self::IMAGE_MAX_WIDTH / max(1.0, $width), self::IMAGE_MAX_HEIGHT / max(1.0, $height));
        $targetWidth = (int) round(max(20.0, $width * $scale));
        $targetHeight = (int) round(max(20.0, $height * $scale));

        $offsetY = $this->reserveOffset($rowNumber, $columnIndex, (float) $targetHeight);

        $drawing = new Drawing();
        $drawing->setName(sprintf(__('Image %d', 'nf-cpt-xlsx-inline'), ++$this->imageIndex));
        $drawing->setDescription($attachment['source_url'] ?? '');
        $drawing->setPath($path);
        $drawing->setCoordinates(Coordinate::stringFromColumnIndex($columnIndex) . $rowNumber);
        $drawing->setResizeProportional(false);
        $drawing->setWidth($targetWidth);
        $drawing->setHeight($targetHeight);
        $drawing->setOffsetX(4);
        $drawing->setOffsetY($offsetY);
        $drawing->setWorksheet($this->dataSheet);

        $this->ensureRowHeight($rowNumber, $offsetY + $targetHeight + 10.0);
    }

    private function embedPdf(array $attachment, int $rowNumber): void
    {
        $iconPath = $this->getPdfIconPath();
        if (!$iconPath) {
            return;
        }

        $columnIndex = $this->pdfColumnIndex;
        $offsetY = $this->reserveOffset($rowNumber, $columnIndex, 22.0);

        $drawing = new Drawing();
        $drawing->setName(sprintf(__('PDF %d', 'nf-cpt-xlsx-inline'), ++$this->pdfIndex));
        $drawing->setDescription($attachment['source_url'] ?? '');
        $drawing->setPath($iconPath);
        $drawing->setCoordinates(Coordinate::stringFromColumnIndex($columnIndex) . $rowNumber);
        $drawing->setHeight(22);
        $drawing->setOffsetX(4);
        $drawing->setOffsetY($offsetY);
        $drawing->setWorksheet($this->dataSheet);

        $coordinate = Coordinate::stringFromColumnIndex($columnIndex) . $rowNumber;
        $cell = $this->dataSheet->getCell($coordinate);
        $label = $attachment['original_filename'] ?? $attachment['zip_name'] ?? __('Document', 'nf-cpt-xlsx-inline');
        $cell->setValueExplicit((string) $label, DataType::TYPE_STRING);
        $cell->getStyle()->getAlignment()->setWrapText(true);
        $cell->getStyle()->getAlignment()->setVertical(Alignment::VERTICAL_TOP);

        if (!empty($attachment['source_url'])) {
            $cell->getHyperlink()->setUrl($attachment['source_url']);
        }

        if (!empty($attachment['zip_name'])) {
            $richText = new RichText();
            $richText->createText(sprintf(__('Linked file: %s', 'nf-cpt-xlsx-inline'), $attachment['zip_name']));
            $cell->getComment()->setText($richText);
            $cell->getComment()->setAuthor('NF Export');
        }

        $this->ensureRowHeight($rowNumber, $offsetY + 32.0);
    }

    private function addAttachmentSheetRow(array $attachment): void
    {
        $sheet = $this->ensureAttachmentsSheet();
        ++$this->attachmentsSheetRow;

        $this->setCellValue($sheet, 1, $this->attachmentsSheetRow, (string) ($attachment['zip_name'] ?? ''));
        $this->setCellValue($sheet, 2, $this->attachmentsSheetRow, (string) ($attachment['original_filename'] ?? ''));
        $this->setCellValue($sheet, 3, $this->attachmentsSheetRow, (string) ($attachment['mime'] ?? ''));
        $this->setCellValue($sheet, 4, $this->attachmentsSheetRow, (string) ($attachment['source_url'] ?? ''));
        $this->setCellValue($sheet, 5, $this->attachmentsSheetRow, (string) ($attachment['post_id'] ?? ''));

        if (!empty($attachment['source_url'])) {
            $coordinate = Coordinate::stringFromColumnIndex(4) . $this->attachmentsSheetRow;
            $sheet->getCell($coordinate)->getHyperlink()->setUrl($attachment['source_url']);
        }

        for ($col = 1; $col <= 5; ++$col) {
            $style = $sheet->getStyle(Coordinate::stringFromColumnIndex($col) . $this->attachmentsSheetRow);
            $style->getAlignment()->setWrapText(true);
            $style->getAlignment()->setVertical(Alignment::VERTICAL_TOP);
        }
    }

    private function ensureAttachmentsSheet(): Worksheet
    {
        if ($this->attachmentsSheet instanceof Worksheet) {
            return $this->attachmentsSheet;
        }

        $this->attachmentsSheet = new Worksheet($this->spreadsheet, 'Attachments');
        $this->spreadsheet->addSheet($this->attachmentsSheet);

        $headers = [
            __('ZIP Name', 'nf-cpt-xlsx-inline'),
            __('Original Filename', 'nf-cpt-xlsx-inline'),
            __('MIME', 'nf-cpt-xlsx-inline'),
            __('Source URL', 'nf-cpt-xlsx-inline'),
            __('Post ID', 'nf-cpt-xlsx-inline'),
        ];

        foreach ($headers as $index => $label) {
            $this->setCellValue($this->attachmentsSheet, $index + 1, 1, $label);
            $style = $this->attachmentsSheet->getStyle(Coordinate::stringFromColumnIndex($index + 1) . '1');
            $style->getFont()->setBold(true);
            $style->getAlignment()->setWrapText(true);
            $style->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
        }

        $this->attachmentsSheet->getColumnDimensionByColumn(1)->setWidth(24.0);
        $this->attachmentsSheet->getColumnDimensionByColumn(2)->setWidth(36.0);
        $this->attachmentsSheet->getColumnDimensionByColumn(3)->setWidth(20.0);
        $this->attachmentsSheet->getColumnDimensionByColumn(4)->setWidth(50.0);
        $this->attachmentsSheet->getColumnDimensionByColumn(5)->setWidth(12.0);

        $this->attachmentsSheetRow = 1;

        return $this->attachmentsSheet;
    }

    private function reserveOffset(int $rowNumber, int $columnIndex, float $height): int
    {
        $key = $rowNumber . ':' . $columnIndex;
        $state = $this->rowOffsets[$key] ?? ['next' => 2];

        $offset = (int) $state['next'];
        $state['next'] = $offset + (int) ceil($height) + 8;
        $this->rowOffsets[$key] = $state;

        return $offset;
    }

    private function ensureRowHeight(int $rowNumber, float $heightPx): void
    {
        $heightPx = max($heightPx, 20.0);
        $current = $this->rowHeights[$rowNumber] ?? 0.0;

        if ($heightPx > $current) {
            $this->rowHeights[$rowNumber] = $heightPx;
            $points = $heightPx * 72 / 96;
            $this->dataSheet->getRowDimension($rowNumber)->setRowHeight($points);
        }
    }

    private function resolveImagePath(array $attachment): ?array
    {
        $path = $attachment['path'] ?? '';
        if (!$path || !file_exists($path)) {
            return null;
        }

        $info = @getimagesize($path);
        if ($info && max($info[0], $info[1]) > 2000) {
            $scaled = $this->createScaledImage($path);
            if ($scaled !== null) {
                $path = $scaled;
                $info = @getimagesize($path);
            }
        }

        if (!$info) {
            return [
                'path'  => $path,
                'width' => self::IMAGE_MAX_WIDTH,
                'height'=> self::IMAGE_MAX_HEIGHT,
            ];
        }

        return [
            'path'  => $path,
            'width' => (float) $info[0],
            'height'=> (float) $info[1],
        ];
    }

    private function createScaledImage(string $path): ?string
    {
        if (!function_exists('wp_get_image_editor')) {
            return null;
        }

        $editor = wp_get_image_editor($path);
        if (is_wp_error($editor)) {
            return null;
        }

        $editor->resize(2000, 2000, false);

        $temp = $this->createTempFile(basename($path));
        $result = $editor->save($temp);
        if (is_wp_error($result)) {
            return null;
        }

        return $result['path'] ?? $temp;
    }

    private function setCellValue(Worksheet $sheet, int $columnIndex, int $rowIndex, string $value): void
    {
        $sheet->setCellValueExplicit(
            Coordinate::stringFromColumnIndex($columnIndex) . $rowIndex,
            $value,
            DataType::TYPE_STRING
        );
    }

    private function createTempFile(string $filename): string
    {
        $extension = pathinfo($filename, PATHINFO_EXTENSION);
        $baseName = $extension ? substr($filename, 0, -(strlen($extension) + 1)) : $filename;
        if ($baseName === '') {
            $baseName = 'nf-export';
        }

        $tempPath = function_exists('wp_tempnam') ? wp_tempnam($filename) : tempnam($this->tempDirectory(), 'nfx');
        if ($tempPath === false || $tempPath === '') {
            throw new RuntimeException('Unable to allocate temporary file.');
        }

        if ($extension !== '') {
            $target = $tempPath . '.' . $extension;
            if (@rename($tempPath, $target)) {
                $tempPath = $target;
            }
        }

        $this->tempFiles[] = $tempPath;

        return $tempPath;
    }

    private function tempDirectory(): string
    {
        if (function_exists('get_temp_dir')) {
            return get_temp_dir();
        }

        $uploadDir = function_exists('wp_upload_dir') ? wp_upload_dir() : null;
        if (is_array($uploadDir) && empty($uploadDir['error']) && !empty($uploadDir['path'])) {
            return trailingslashit($uploadDir['path']);
        }

        return sys_get_temp_dir();
    }

    private function cleanupTempFiles(): void
    {
        foreach ($this->tempFiles as $path) {
            if ($path && file_exists($path)) {
                @unlink($path);
            }
        }

        $this->tempFiles = [];
    }

    private function saveSpreadsheet(string $path): void
    {
        $writer = IOFactory::createWriter($this->spreadsheet, 'Xlsx');
        if ($writer instanceof Xlsx) {
            $writer->setPreCalculateFormulas(false);
        }

        $writer->save($path);
        $this->spreadsheet->disconnectWorksheets();
    }

    private function getPdfIconPath(): ?string
    {
        if ($this->pdfIconPath === '') {
            return null;
        }

        if ($this->pdfIconPath !== null) {
            return $this->pdfIconPath;
        }

        $binary = base64_decode(self::PDF_ICON_BASE64, true);
        if ($binary === false) {
            $this->pdfIconPath = '';
            return null;
        }

        $path = $this->createTempFile('pdf-icon.png');
        if (file_put_contents($path, $binary) === false) {
            $this->pdfIconPath = '';
            return null;
        }

        $this->pdfIconPath = $path;

        return $this->pdfIconPath;
    }

    private function createManifest(): array
    {
        $manifest = ['attachments' => []];

        foreach ($this->attachments as $attachment) {
            $manifest['attachments'][] = [
                'zip_name'          => $attachment['zip_name'] ?? '',
                'original_filename' => $attachment['original_filename'] ?? '',
                'mime'              => $attachment['mime'] ?? '',
                'type'              => $attachment['type'] ?? '',
                'is_signature'      => !empty($attachment['is_signature']),
                'source_url'        => $attachment['source_url'] ?? '',
                'post_id'           => isset($attachment['post_id']) ? (int) $attachment['post_id'] : 0,
            ];
        }

        return $manifest;
    }

    private function configureTheme(): void
    {
        $theme = $this->spreadsheet->getTheme();
        $theme->setThemeColorName('Office');
        $theme->setThemeFontName('Office');
        $theme->setMajorFontValues('Cambria', '', 'Times New Roman', []);
        $theme->setMinorFontValues('Calibri', '', 'Times New Roman', []);

        $defaultColours = [
            'lt1'      => 'FFFFFF',
            'dk1'      => '000000',
            'lt2'      => 'EEECE1',
            'dk2'      => '1F497D',
            'accent1'  => '4F81BD',
            'accent2'  => 'C0504D',
            'accent3'  => '9BBB59',
            'accent4'  => '8064A2',
            'accent5'  => '4BACC6',
            'accent6'  => 'F79646',
            'hlink'    => '0000FF',
            'folHlink' => '800080',
        ];

        foreach ($defaultColours as $name => $value) {
            $theme->setThemeColor($name, $value);
        }

        $this->spreadsheet->resetThemeFonts();
    }
}
