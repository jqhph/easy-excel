<?php

namespace Tests\Exporters;

use Dcat\EasyExcel\Excel;
use Tests\TestCase;

class WithGeneratorTest extends WithArrayTest
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
            // 保存
            $generators = [
                'sheet1' => $this->makeGenerator($users1),
                'sheet2' =>  $this->makeGenerator($users2),
            ];

            Excel::export($generators)->store($storePath);
        });
    }


    protected function assertWithSheet($type, $key)
    {
        $users = include __DIR__.'/../resources/users.php';

        $storePath = $this->generateTempFilePath($type);

        // 保存
        Excel::export($this->makeGenerator($users))->store($storePath);

        // 读取
        $this->assertSingleSheet($storePath, $key, $users);
    }

    protected function makeGenerator(array $values)
    {
        while ($value = array_shift($values)) {
            yield $value;
        }
    }

}
