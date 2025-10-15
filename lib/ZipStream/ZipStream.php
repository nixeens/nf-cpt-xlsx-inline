<?php

namespace ZipStream;

use ZipStream\Exception\OverflowException;
use ZipStream\Option\Archive;

class ZipStream
{
    /** @var resource */
    private $outputStream;

    private ?\ZipArchive $zip = null;

    private string $tempFile;

    private bool $finished = false;

    private bool $closeOutputStream = false;

    public function __construct(
        ?string $name = null,
        ?Archive $options = null,
        bool $enableZip64 = true,
        $outputStream = null,
        bool $sendHttpHeaders = true,
        bool $defaultEnableZeroHeader = true
    ) {
        if ($options instanceof Archive) {
            $enableZip64 = $options->getEnableZip64();
            $outputStream = $outputStream ?? $options->getOutputStream();
            $sendHttpHeaders = $options->getSendHttpHeaders();
        }

        $this->tempFile = tempnam(sys_get_temp_dir(), 'zipstream_');
        if ($this->tempFile === false) {
            throw new \RuntimeException('Unable to create temporary file for ZipStream.');
        }

        if (!class_exists(\ZipArchive::class)) {
            throw new \RuntimeException('The ZipArchive extension is required to use the bundled ZipStream implementation.');
        }

        $this->zip = new \ZipArchive();
        $openResult = $this->zip->open($this->tempFile, \ZipArchive::OVERWRITE | \ZipArchive::CREATE);
        if ($openResult !== true) {
            throw new \RuntimeException('Unable to initialise ZipArchive backend.');
        }

        if ($outputStream === null) {
            $outputStream = fopen('php://output', 'wb');
            $this->closeOutputStream = true;
        }

        if (!is_resource($outputStream)) {
            throw new \InvalidArgumentException('Output stream must be a valid resource.');
        }

        $this->outputStream = $outputStream;
    }

    public function __destruct()
    {
        $this->cleanup();
    }

    public function addFile(string $name, string $content, array $options = []): void
    {
        if ($this->finished) {
            throw new \LogicException('Cannot add files after finish() has been called.');
        }

        $this->zip->addFromString($name, $content);
    }

    public function finish(): void
    {
        if ($this->finished) {
            return;
        }

        $closeResult = $this->zip->close();
        if ($closeResult !== true) {
            throw new OverflowException('Unable to finalise ZIP archive.');
        }

        $source = fopen($this->tempFile, 'rb');
        if ($source === false) {
            throw new OverflowException('Unable to read temporary ZIP archive.');
        }

        try {
            while (!feof($source)) {
                $chunk = fread($source, 1048576);
                if ($chunk === false) {
                    throw new OverflowException('Unable to read data from temporary ZIP archive.');
                }

                if ($chunk === '') {
                    continue;
                }

                $written = fwrite($this->outputStream, $chunk);
                if ($written === false || $written !== strlen($chunk)) {
                    throw new OverflowException('Unable to write ZIP archive to output stream.');
                }
            }
        } finally {
            fclose($source);
        }

        if ($this->closeOutputStream) {
            fclose($this->outputStream);
        } else {
            fflush($this->outputStream);
        }

        $this->finished = true;
        $this->cleanup();
    }

    private function cleanup(): void
    {
        if ($this->zip instanceof \ZipArchive) {
            // ZipArchive does not provide a method to check if close has been called, but
            // calling close multiple times is harmless for our usage. Suppress any warnings.
            if (!$this->finished) {
                try {
                    $this->zip->close();
                } catch (\Throwable $e) {
                    // Ignore cleanup errors.
                }
            }

            $this->zip = null;
        }

        if ($this->tempFile !== '' && file_exists($this->tempFile)) {
            @unlink($this->tempFile);
        }

        $this->tempFile = '';
    }
}
