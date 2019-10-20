<?php

namespace Tests\Exporters;

use Dcat\EasyExcel\Excel;

trait Exporter
{
    protected $tempFiles = [];

    protected function assertSingleSheet(string $file, $key, array $compares)
    {
        // 读取
        $sheetsArray = Excel::import($file)->toArray();

        $this->assertIsArray($sheetsArray);
        $this->assertEquals(count($sheetsArray), 1);
        $this->assertTrue(isset($sheetsArray[$key]));

        $this->assertEquals(array_values($sheetsArray[$key]), $compares);
    }

    protected function generateTempFilePath(string $type)
    {
        return $this->tempFiles[] = __DIR__.'/../resources/'
            .uniqid(microtime(true).mt_rand(0, 1000))
            .'.'.$type;
    }

    public function tearDown(): void
    {
        parent::tearDown();

        // 删除临时文件
        foreach ($this->tempFiles as $file) {
            if (is_file($file)) {
                @unlink($file);
            }
        }
    }
}
