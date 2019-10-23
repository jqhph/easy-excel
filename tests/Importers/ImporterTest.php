<?php

namespace Tests\Importers;

use Dcat\EasyExcel\Excel;
use Tests\TestCase;
use Dcat\EasyExcel\Contracts;

class ImporterTest extends TestCase
{
    /**
     * @group importer
     */
    public function testCsv()
    {
        $file = __DIR__.'/../resources/test.csv';

        $this->assertSheet($file, 0);
    }

    /**
     * @group importer
     */
    public function testXlsx()
    {
        $file = __DIR__.'/../resources/test.xlsx';

        $this->assertSheet($file, 'Sheet1');
    }

    /**
     * @group importer
     */
    public function testWithHeaders()
    {
        $xlsx = __DIR__.'/../resources/test.xlsx';
        $csv  = __DIR__.'/../resources/test.csv';

        $headers = [
            'ID', 'NAME', 'EMAIL',
        ];

        // xlsx
        $xlsxSheetArray = Excel::import($xlsx)
            ->headers($headers)
            ->first()
            ->toArray();

        $this->assertEquals($headers, array_keys(current($xlsxSheetArray)));

        // csv
        $csvSheetArray = Excel::import($csv)
            ->headers($headers)
            ->first()
            ->toArray();

        $this->assertEquals($headers, array_keys(current($csvSheetArray)));
    }

    /**
     * @group importer
     */
    public function testWithoutHeaders()
    {
        $xlsx = __DIR__.'/../resources/test.xlsx';
        $csv  = __DIR__.'/../resources/test.csv';

        // xlsx
        $sheetArray = Excel::import($xlsx)
            ->withoutHeaders()
            ->first()
            ->toArray();

        $this->assertEquals(range(0, 7), array_keys(current($sheetArray)));

        // csv
        $sheetArray = Excel::import($csv)
            ->withoutHeaders()
            ->first()
            ->toArray();

        $this->assertEquals(range(0, 7), array_keys(current($sheetArray)));
    }

    /**
     * @group importer
     */
    public function testWorking()
    {
        // xlsx
        $file = __DIR__.'/../resources/test.xlsx';

        $sheetArray = Excel::import($file)->active()->toArray();
        $this->validateSheetArray($sheetArray);

        // csv
        $file = __DIR__.'/../resources/test.csv';

        $sheetArray = Excel::import($file)->active()->toArray();
        $this->validateSheetArray($sheetArray);
    }

    /**
     * @group importer
     */
    public function testFirst()
    {
        // xlsx
        $file = __DIR__.'/../resources/test.xlsx';

        $sheetArray = Excel::import($file)->first()->toArray();
        $this->validateSheetArray($sheetArray);

        // csv
        $file = __DIR__.'/../resources/test.csv';

        $sheetArray = Excel::import($file)->first()->toArray();
        $this->validateSheetArray($sheetArray);
    }

    /**
     * @group importer
     */
    public function testGetSheet()
    {
        // xlsx
        $xlsx = __DIR__.'/../resources/test.xlsx';

        $sheetArray = Excel::import($xlsx)->sheet('Sheet1')->toArray();
        $this->validateSheetArray($sheetArray);

        $sheetArray = Excel::import($xlsx)->sheet(0)->toArray();
        $this->validateSheetArray($sheetArray);

        // csv
        $csv = __DIR__.'/../resources/test.csv';

        $sheetArray = Excel::import($csv)->sheet(0)->toArray();
        $this->validateSheetArray($sheetArray);
    }

    /**
     * @group importer
     */
    public function testEach()
    {
        $xlsx = __DIR__.'/../resources/test.xlsx';
        $csv  = __DIR__.'/../resources/test.csv';

        Excel::import($xlsx)->each(function (Contracts\Sheet $sheet) {
            $this->validateSheetArray($sheet->toArray());
        });

        Excel::import($csv)->each(function (Contracts\Sheet $sheet) {
            $this->validateSheetArray($sheet->toArray());
        });
    }

    /**
     * @group importer
     * @see  \Tests\Importers::testToArray
     */
    public function testToArray()
    {
        $this->assertTrue(true);
    }

    protected function assertSheet($file, $key)
    {
        $sheetsArray = Excel::import($file)->toArray();

        $this->assertIsArray($sheetsArray);
        $this->assertEquals(count($sheetsArray), 1);

        $this->assertTrue(isset($sheetsArray[$key]));
        $this->assertIsArray($sheetsArray[$key]);

        $this->validateSheetArray($sheetsArray[$key]);
    }

    protected function validateSheetArray(array $sheetArray)
    {
        $this->assertEquals(count($sheetArray), 50);

        $users = include __DIR__.'/../resources/users.php';

        $this->assertEquals(array_values($sheetArray), array_slice($users, 0, 50));
    }

}
