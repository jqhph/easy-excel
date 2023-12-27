<?php

namespace Dcat\EasyExcel\Spout;

use OpenSpout\Common\Creator\HelperFactory;
use OpenSpout\Common\Exception\UnsupportedTypeException;
use OpenSpout\Common\Type;
use OpenSpout\Writer\Common\Creator\InternalEntityFactory;
use OpenSpout\Writer\Common\Creator\Style\StyleBuilder;
use OpenSpout\Writer\CSV\Manager\OptionsManager as CSVOptionsManager;
use OpenSpout\Writer\ODS\Creator\HelperFactory as ODSHelperFactory;
use OpenSpout\Writer\ODS\Creator\ManagerFactory as ODSManagerFactory;
use OpenSpout\Writer\ODS\Manager\OptionsManager as ODSOptionsManager;
use OpenSpout\Writer\WriterInterface;
use OpenSpout\Writer\XLSX\Creator\HelperFactory as XLSXHelperFactory;
use OpenSpout\Writer\XLSX\Creator\ManagerFactory as XLSXManagerFactory;
use OpenSpout\Writer\XLSX\Manager\OptionsManager as XLSXOptionsManager;
use Dcat\EasyExcel\Spout\Writers\CSVWriter;
use Dcat\EasyExcel\Spout\Writers\ODSWriter;
use Dcat\EasyExcel\Spout\Writers\XLSXWriter;

class WriterFactory
{
    /**
     * This creates an instance of the appropriate writer, given the extension of the file to be written.
     *
     * @param  string  $path  The path to the spreadsheet file. Supported extensions are .csv,.ods and .xlsx
     * @return WriterInterface
     *
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     */
    public static function createFromFile(string $path)
    {
        $extension = strtolower(pathinfo($path, PATHINFO_EXTENSION));

        return static::createFromType($extension);
    }

    /**
     * This creates an instance of the appropriate writer, given the type of the file to be written.
     *
     * @param  string  $writerType  Type of the writer to instantiate
     * @return WriterInterface
     *
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     */
    public static function createFromType($writerType)
    {
        switch ($writerType) {
            case Type::CSV: return static::createCSVWriter();
            case Type::XLSX: return static::createXLSXWriter();
            case Type::ODS: return static::createODSWriter();
            default:
                throw new UnsupportedTypeException('No writers supporting the given type: '.$writerType);
        }
    }

    /**
     * @return CSVWriter
     */
    protected static function createCSVWriter()
    {
        $optionsManager = new CSVOptionsManager();
        $globalFunctionsHelper = new GlobalFunctionsHelper();

        $helperFactory = new HelperFactory();

        return new CSVWriter($optionsManager, $globalFunctionsHelper, $helperFactory);
    }

    /**
     * @return XLSXWriter
     */
    protected static function createXLSXWriter()
    {
        $styleBuilder = new StyleBuilder();
        $optionsManager = new XLSXOptionsManager($styleBuilder);
        $globalFunctionsHelper = new GlobalFunctionsHelper();

        $helperFactory = new XLSXHelperFactory();
        $managerFactory = new XLSXManagerFactory(new InternalEntityFactory(), $helperFactory);

        return new XLSXWriter($optionsManager, $globalFunctionsHelper, $helperFactory, $managerFactory);
    }

    /**
     * @return ODSWriter
     */
    protected static function createODSWriter()
    {
        $styleBuilder = new StyleBuilder();
        $optionsManager = new ODSOptionsManager($styleBuilder);
        $globalFunctionsHelper = new GlobalFunctionsHelper();

        $helperFactory = new ODSHelperFactory();
        $managerFactory = new ODSManagerFactory(new InternalEntityFactory(), $helperFactory);

        return new ODSWriter($optionsManager, $globalFunctionsHelper, $helperFactory, $managerFactory);
    }
}
