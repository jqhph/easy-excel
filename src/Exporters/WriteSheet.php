<?php

namespace Dcat\EasyExcel\Exporters;

use Box\Spout\Common\Entity\Row;
use Box\Spout\Common\Entity\Style\Style;
use Box\Spout\Writer\Common\Creator\WriterEntityFactory;
use Box\Spout\Writer\CSV\Writer as CsvWriter;
use Box\Spout\Writer\WriterInterface;
use Dcat\EasyExcel\Support\SheetCollection;
use Generator;
use Dcat\EasyExcel\Support\Arr;

/**
 * @mixin \Dcat\EasyExcel\Contracts\Exporter
 */
trait WriteSheet
{
    /**
     * @var array
     */
    protected $writedHeadings = [];

    /**
     * @param WriterInterface $writer
     * @return WriterInterface
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    protected function writeSheets(WriterInterface $writer)
    {
        $data    = $this->makeSheetsArray();
        $keys    = array_keys($data);
        $lastKey = end($keys);

        foreach ($data as $index => &$collection) {
            if ($collection instanceof \Generator) {
                $this->writeRowsFromGenerator($writer, $index, $collection);
            } else {
                $collection = $this->convertToArray($collection);

                $this->writeRowsFromArray($writer, $index, $collection);
            }

            if (is_string($index) && method_exists($writer, 'getCurrentSheet')) {
                $writer->getCurrentSheet()->setName($index);
            }

            if ($lastKey !== $index && method_exists($writer, 'addNewSheetAndMakeItCurrent')) {
                $writer->addNewSheetAndMakeItCurrent();
            }
        }

        return $writer;
    }

    /**
     * @param WriterInterface $writer
     * @param int|string $index
     * @param array $rows
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    protected function writeRowsFromArray(WriterInterface $writer, $index, array &$rows)
    {
        // Add heading row.
        if ($this->canWriteHeadings($writer, $index)) {
            $this->writeHeadings($writer, current($rows));
        }

        foreach ($rows as &$row) {
            $this->writeRow($writer, $row, $index);
        }
    }

    /**
     * @param WriterInterface $writer
     * @param $index
     * @param Generator $generator
     */
    protected function writeRowsFromGenerator(WriterInterface $writer, $index, Generator $generator)
    {
        foreach ($generator as $key => $items) {
            $items = $this->convertToArray($items);

            if (! is_array(current($items))) {
                $items = [$items];
            }

            $this->writeRowsFromArray($writer, $index, $items);
        }
    }

    /**
     * @param WriterInterface $writer
     * @param array $item
     * @param string|int $index
     */
    protected function writeRow(WriterInterface $writer, array &$item, $index)
    {
        $item = $this->formatRow($item);

        if ($this->rowCallback) {
            $item = call_user_func($this->rowCallback, $item, $index);
        }

        if ($item && is_array($item)) {
            $item = $this->makeDefaultRow($item);
        }

        if ($item instanceof Row) {
            $writer->addRow($item);
        }
    }

    /**
     * @param WriterInterface $writer
     * @param array $firstRow
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    protected function writeHeadings(WriterInterface $writer, array $firstRow)
    {
        $writer->addRow(
            $this->makeDefaultRow(
                $this->headings ?: array_keys($firstRow),
                $this->headingStyle
            )
        );
    }

    /**
     * @param array $item
     * @param Style|null $style
     * @return Row
     */
    protected function makeDefaultRow(array $item, ?Style $style = null)
    {
        if ($style) {
            return WriterEntityFactory::createRowFromArray($item, $style);
        }

        return WriterEntityFactory::createRowFromArray($item);
    }

    /**
     * @param WriterInterface $writer
     * @param int|string $index
     * @return bool
     */
    protected function canWriteHeadings(WriterInterface $writer, $index): bool
    {
        if (
            $this->headings === false
            || ! empty($this->writedHeadings[$index])
            || ($this->writedHeadings && $writer instanceof CsvWriter)
        ) {
            return false;
        }

        $this->writedHeadings[$index] = true;

        return true;
    }

    /**
     * @return Generator[]|array
     */
    protected function makeSheetsArray(): array
    {
        $data = $this->data;

        if ($data instanceof ChunkingQuery) {
            return $data->makeGenerators();
        }

        if ($this->data instanceof \Closure) {
            $data = $data($this);
        }

        if ($data instanceof SheetCollection) {
            $data = $data->toArray();
        }

        if (
            (is_array($data) && ! Arr::isAssoc($data))
            || $data instanceof Generator
        ) {
            return [&$data];
        }

        return (array) $data;
    }

    /**
     * @param array $row
     * @return array
     */
    protected function formatRow(array &$row)
    {
        $strings = [];

        foreach ($this->filterAndSortByHeaders($row) as &$value) {
            $value = is_int($value) || is_float($value) || is_null($value) ? (string) $value : $value;

            if (is_string($value)) {
                $strings[] = $value;
            }
        }

        return $strings;
    }

    /**
     * @param array $row
     * @return array
     */
    public function filterAndSortByHeaders(array &$row)
    {
        if (! $this->headings) {
            return $row;
        }

        $newRow = [];
        foreach ($this->headings as $key => &$label) {
            $newRow[$key] = $row[$key] ?? null;
        }

        return $newRow;
    }


}
