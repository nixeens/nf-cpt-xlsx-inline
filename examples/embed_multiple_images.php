<?php

declare(strict_types=1);

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

require __DIR__ . '/../vendor/autoload.php';

/**
 * Ensure a square image exists with a solid fill and optional decorator.
 *
 * @param array{0:int,1:int,2:int} $backgroundRgb
 */
function ensureImage(string $path, string $type, array $backgroundRgb, ?callable $decorator = null): void
{
    if (is_file($path)) {
        return;
    }

    $size = 140;
    $image = imagecreatetruecolor($size, $size);

    if ($type === 'png') {
        imagealphablending($image, false);
        imagesavealpha($image, true);
        $transparent = imagecolorallocatealpha($image, 0, 0, 0, 127);
        imagefilledrectangle($image, 0, 0, $size - 1, $size - 1, $transparent);
        imagealphablending($image, true);
    }

    $background = imagecolorallocate($image, $backgroundRgb[0], $backgroundRgb[1], $backgroundRgb[2]);
    imagefilledrectangle($image, 0, 0, $size - 1, $size - 1, $background);

    if ($decorator !== null) {
        $decorator($image);
    }

    if ($type === 'png') {
        imagepng($image, $path);
    } else {
        imagejpeg($image, $path, 90);
    }

    imagedestroy($image);
}

$assetsDir = __DIR__ . '/assets';
if (!is_dir($assetsDir)) {
    mkdir($assetsDir, 0777, true);
}

ensureImage($assetsDir . '/image1.png', 'png', [52, 152, 219]);
ensureImage($assetsDir . '/image2.jpg', 'jpg', [46, 204, 113]);
ensureImage(
    $assetsDir . '/pdf1.png',
    'png',
    [231, 76, 60],
    static function ($image): void {
        $textColor = imagecolorallocate($image, 255, 255, 255);
        $font = 5;
        $text = 'PDF';
        $textWidth = imagefontwidth($font) * strlen($text);
        $textHeight = imagefontheight($font);
        $x = (int) ((imagesx($image) - $textWidth) / 2);
        $y = (int) ((imagesy($image) - $textHeight) / 2);
        imagestring($image, $font, $x, $y, $text, $textColor);
    }
);

$imagePaths = [
    'image1' => realpath($assetsDir . '/image1.png'),
    'image2' => realpath($assetsDir . '/image2.jpg'),
    'pdf1' => realpath($assetsDir . '/pdf1.png'),
];

foreach ($imagePaths as $label => $absolutePath) {
    if ($absolutePath === false) {
        throw new RuntimeException(sprintf('Unable to resolve absolute path for %s.', $label));
    }
}

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle('Embedded Media');

$drawings = [
    [
        'path' => $imagePaths['image1'],
        'cell' => 'A1',
        'offsetX' => 5,
        'offsetY' => 5,
        'height' => 80,
    ],
    [
        'path' => $imagePaths['image2'],
        'cell' => 'C5',
        'offsetX' => 5,
        'offsetY' => 5,
        'height' => 100,
    ],
    [
        'path' => $imagePaths['pdf1'],
        'cell' => 'E10',
        'offsetX' => 5,
        'offsetY' => 5,
        'height' => 90,
    ],
];

foreach ($drawings as $index => $spec) {
    $drawing = new Drawing();
    $drawing->setPath($spec['path']);
    $drawing->setCoordinates($spec['cell']);
    $drawing->setOffsetX($spec['offsetX']);
    $drawing->setOffsetY($spec['offsetY']);
    $drawing->setHeight($spec['height']);
    $drawing->setWorksheet($sheet);
    $sheet->setCellValue($spec['cell'], sprintf('Image %d', $index + 1));
}

$outputDir = __DIR__ . '/output';
if (!is_dir($outputDir)) {
    mkdir($outputDir, 0777, true);
}

$outputPath = $outputDir . '/embedded-images-one-cell-anchor.xlsx';

$writer = new Xlsx($spreadsheet);
$writer->save($outputPath);

printf("Workbook saved to %s\n", $outputPath);

