<?php

declare(strict_types=1);

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

require __DIR__ . '/../vendor/autoload.php';

if (!extension_loaded('gd')) {
    throw new \RuntimeException('The GD extension is required to generate sample images.');
}

function createSolidColorPng(string $path, array $rgb, int $width = 240, int $height = 160): void
{
    $image = imagecreatetruecolor($width, $height);

    if ($image === false) {
        throw new \RuntimeException('Unable to create GD image resource.');
    }

    imagealphablending($image, false);
    imagesavealpha($image, true);

    $background = imagecolorallocatealpha($image, $rgb[0], $rgb[1], $rgb[2], 0);
    imagefilledrectangle($image, 0, 0, $width, $height, $background);

    $borderColor = imagecolorallocatealpha($image, 255, 255, 255, 0);
    imagerectangle($image, 0, 0, $width - 1, $height - 1, $borderColor);

    if (!imagepng($image, $path)) {
        imagedestroy($image);
        throw new \RuntimeException(sprintf('Unable to write PNG image to %s', $path));
    }

    imagedestroy($image);
}

$outputDir = __DIR__ . '/output';
if (!is_dir($outputDir) && !mkdir($outputDir, 0777, true) && !is_dir($outputDir)) {
    throw new \RuntimeException(sprintf('Unable to create output directory at %s', $outputDir));
}

$imageDir = $outputDir . '/generated-images';
if (!is_dir($imageDir) && !mkdir($imageDir, 0777, true) && !is_dir($imageDir)) {
    throw new \RuntimeException(sprintf('Unable to create image directory at %s', $imageDir));
}

$imageDefinitions = [
    [
        'filename' => 'leaf-green.png',
        'rgb' => [39, 174, 96],
        'cell' => 'B2',
        'offsetX' => 12,
        'offsetY' => 6,
        'height' => 120,
    ],
    [
        'filename' => 'sky-blue.png',
        'rgb' => [41, 128, 185],
        'cell' => 'D8',
        'offsetX' => 18,
        'offsetY' => 4,
        'height' => 90,
    ],
];

foreach ($imageDefinitions as &$definition) {
    $definition['path'] = $imageDir . '/' . $definition['filename'];
    createSolidColorPng($definition['path'], $definition['rgb']);
}
unset($definition);

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle('Theme + Images');

$sheet->setCellValue('A1', 'Embedded images anchored to a single cell.');

foreach ($imageDefinitions as $index => $definition) {
    $drawing = new Drawing();
    $drawing->setPath($definition['path']);
    $drawing->setCoordinates($definition['cell']);
    $drawing->setOffsetX($definition['offsetX']);
    $drawing->setOffsetY($definition['offsetY']);
    $drawing->setHeight($definition['height']);
    $drawing->setWorksheet($sheet);

    $sheet->setCellValue($definition['cell'], sprintf('Image %d', $index + 1));
}

$outputPath = $outputDir . '/one-cell-anchor-theme.xlsx';
$writer = new Xlsx($spreadsheet);
$writer->save($outputPath);

echo sprintf("Workbook saved to %s\n", $outputPath);
