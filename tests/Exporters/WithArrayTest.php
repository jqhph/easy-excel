<?php

namespace Tests\Exporters;

use Dcat\EasyExcel\Excel;
use Tests\TestCase;

class WithArrayTest extends TestCase
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
        $this->assertWithSheets('xlsx');
    }

    protected function assertWithSheets($type, \Closure $callback = null)
    {
        $users = include __DIR__.'/../resources/users.php';

        $users1 = array_slice($users, 0, 30);
        $users2 = array_values(array_slice($users, 30, 30));

        $storePath = $this->generateTempFilePath($type);

        if ($callback) {
            $callback($storePath, $users1, $users2);
        } else {
            // 保存
            Excel::export(['sheet1' => $users1, 'sheet2' => $users2])->store($storePath);
        }

        // 读取
        $sheetsArray = Excel::import($storePath)->toArray();

        $this->assertIsArray($sheetsArray);
        $this->assertEquals(count($sheetsArray), 2);

        $this->assertTrue(isset($sheetsArray['sheet1']));
        $this->assertTrue(isset($sheetsArray['sheet2']));
        $this->assertEquals(array_values($sheetsArray['sheet1']), $users1);
        $this->assertEquals(array_values($sheetsArray['sheet2']), $users2);
    }

    protected function assertWithSheet($type, $key)
    {
        $users = include __DIR__.'/../resources/users.php';

        $storePath = $this->generateTempFilePath($type);

        // 保存
        Excel::export($users)->store($storePath);

        // 读取
        $this->assertSingleSheet($storePath, $key, $users);
    }

}
