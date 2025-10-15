<?php

namespace ZipStream\Option;

class Archive
{
    private bool $enableZip64 = true;

    /** @var resource|null */
    private $outputStream = null;

    private bool $sendHttpHeaders = true;

    public function setEnableZip64(bool $enable): void
    {
        $this->enableZip64 = $enable;
    }

    public function getEnableZip64(): bool
    {
        return $this->enableZip64;
    }

    /**
     * @param resource|null $stream
     */
    public function setOutputStream($stream): void
    {
        $this->outputStream = $stream;
    }

    /**
     * @return resource|null
     */
    public function getOutputStream()
    {
        return $this->outputStream;
    }

    public function setSendHttpHeaders(bool $send): void
    {
        $this->sendHttpHeaders = $send;
    }

    public function getSendHttpHeaders(): bool
    {
        return $this->sendHttpHeaders;
    }
}
