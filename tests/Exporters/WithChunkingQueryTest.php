<?php

namespace Tests\Exporters;

use Dcat\EasyExcel\Excel;
use Dcat\EasyExcel\Support\SheetCollection;

class WithChunkingQueryTest extends WithArrayTest
{
    use Exporter;

    /**
     * @group importer
     */
    public function testCsv()
    {
        $this->assertWithSheet('csv', 0);
    }

    public function testXlsx()
    {
        $this->assertWithSheet('xlsx', 'Sheet1');
        $this->assertWithSheets('xlsx', function ($storePath, $users1, $users2) {

            $chunkSize = 10;

            $users1 = new SheetCollection($users1);
            $users2 = new SheetCollection($users2);

            // 保存
            Excel::export()
                ->chunk([
                    'sheet1' => function (int $times) use ($users1, $chunkSize) {
                        return $users1->forPage($times, $chunkSize);
                    },
                    'sheet2' => function (int $times) use ($users2, $chunkSize) {
                        return $users2->forPage($times, $chunkSize);
                    },
                ])
                ->store($storePath);
        });
    }

    protected function assertWithSheet($type, $key)
    {
        $users = include __DIR__.'/../resources/users.php';

        $storePath = $this->generateTempFilePath($type);

        $collection = (new SheetCollection($users));

        // 保存
        Excel::export()
            ->chunk(function (int $times) use ($collection) {
                $chunkSize = 10;

                return $collection->forPage($times, $chunkSize)->toArray();
            })
            ->store($storePath);

        // 读取
        $this->assertSingleSheet($storePath, $key, $users);
    }

}
