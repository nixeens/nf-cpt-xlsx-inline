<?php

declare(strict_types=1);

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

require __DIR__ . '/../vendor/autoload.php';

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle('Embedded Images');

$images = [
    [
        'path' => '/tmp/image1.png',
        'cell' => 'A1',
        'offsetX' => 5,
        'offsetY' => 5,
        'height' => 80,
    ],
    [
        'path' => '/tmp/image2.png',
        'cell' => 'C5',
        'offsetX' => 10,
        'offsetY' => 12,
        'height' => 120,
    ],
    [
        'path' => '/tmp/pdf1.png',
        'cell' => 'E10',
        'offsetX' => 2,
        'offsetY' => 2,
        'height' => 100,
    ],
];

foreach ($images as $imageSpec) {
    $drawing = new Drawing();
    $drawing->setPath($imageSpec['path']);
    $drawing->setCoordinates($imageSpec['cell']);
    $drawing->setOffsetX($imageSpec['offsetX']);
    $drawing->setOffsetY($imageSpec['offsetY']);
    $drawing->setHeight($imageSpec['height']);
    $drawing->setWorksheet($sheet);
}

$writer = new Xlsx($spreadsheet);
$writer->save('/tmp/export_with_embedded_images.xlsx');

