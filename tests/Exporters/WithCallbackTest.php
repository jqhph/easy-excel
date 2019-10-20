<?php

namespace Tests\Exporters;

use Dcat\EasyExcel\Excel;
use Tests\TestCase;

class WithCallbackTest extends WithArrayTest
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
            $callback = function () use ($users1, $users2) {
                return ['sheet1' => $users1, 'sheet2' => $users2];
            };

            // 保存
            Excel::export($callback)->store($storePath);
        });
    }

    protected function assertWithSheet($type, $key)
    {
        $users = include __DIR__.'/../resources/users.php';

        $storePath = $this->generateTempFilePath($type);

        $callback = function () use ($users) {
            return $users;
        };

        // 保存
        Excel::export($callback)->store($storePath);

        // 读取
        $this->assertSingleSheet($storePath, $key, $users);
    }

}
