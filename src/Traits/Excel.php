<?php

namespace Dcat\EasyExcel\Traits;

use Box\Spout\Common\Type;
use Box\Spout\Reader\CSV\Reader as CSVReader;
use Box\Spout\Writer\CSV\Writer as CSVWriter;

trait Excel
{
    /**
     * @var string
     */
    protected $type;

    /**
     * @var \Closure
     */
    protected $optionCallback;

    /**
     * @var array|false
     */
    protected $headers = [];

    /**
     * @var \Closure
     */
    protected $headerCallback;

    /**
     * @var array
     */
    protected $csvConfiguration = [
        'delimiter' => ',',
        'enclosure' => '"',
        'encoding'  => 'UTF-8',
        'bom'       => true,
    ];

    /**
     * @param string|null $type
     * @return $this
     */
    public function type(?string $type)
    {
        $this->type = $type;

        return $this;
    }

    /**
     * @param \Closure $callback
     * @return $this
     */
    public function option(\Closure $callback)
    {
        $this->optionCallback = $callback;

        return $this;
    }

    /**
     * @param array $headers
     * @param \Closure|null $callback
     * @return $this
     */
    public function headers(array $headers, \Closure $callback = null)
    {
        $this->headers = $headers;

        $this->headerCallback = $callback;

        return $this;
    }

    /**
     * @return array|false
     */
    public function getHeaders()
    {
        return $this->headers;
    }

    /**
     * @return $this
     */
    public function withoutHeaders()
    {
        $this->headers = false;

        return $this;
    }

    /**
     * @return $this
     */
    public function csv()
    {
        return $this->type(Type::CSV);
    }

    /**
     * @return $this
     */
    public function ods()
    {
        return $this->type(Type::ODS);
    }

    /**
     * @return $this
     */
    public function xlsx()
    {
        return $this->type(Type::XLSX);
    }

    /**
     * @return string|null
     */
    public function getType()
    {
        return $this->type;
    }

    /**
     * @param string $delimiter
     * @param string $enclosure
     * @param string $encoding
     * @param bool $bom
     *
     * @return $this
     */
    public function configureCsv(
        string $delimiter = ',',
        string $enclosure = '"',
        string $encoding = 'UTF-8',
        bool $bom = false
    )
    {
        $this->csvConfiguration = compact('delimiter', 'enclosure', 'encoding', 'bom');

        return $this;
    }

    /**
     * @param \Box\Spout\Reader\ReaderInterface|\Box\Spout\Writer\WriterInterface $readerOrWriter
     */
    protected function configure(&$readerOrWriter)
    {
        if ($readerOrWriter instanceof CSVReader || $readerOrWriter instanceof CSVWriter) {
            $readerOrWriter->setFieldDelimiter($this->csvConfiguration['delimiter']);
            $readerOrWriter->setFieldEnclosure($this->csvConfiguration['enclosure']);

            if ($readerOrWriter instanceof CSVReader) {
                $readerOrWriter->setEncoding($this->csvConfiguration['encoding']);
            }

            if ($readerOrWriter instanceof CSVWriter) {
                $readerOrWriter->setShouldAddBOM($this->csvConfiguration['bom']);
            }
        }

        if ($this->optionCallback) {
            ($this->optionCallback)($readerOrWriter);
        }
    }

    /**
     * @param string $fileName
     * @return string
     */
    protected function prepareFileName(string $fileName)
    {
        if ($this->type && strpos($fileName, '.') === false) {
            return $fileName.'.'.$this->type;
        }

        return $fileName;
    }
}
