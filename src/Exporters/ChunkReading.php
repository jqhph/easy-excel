<?php

namespace Dcat\EasyExcel\Exporters;

class ChunkReading
{
    /**
     * @var callable[]
     */
    protected $readers;

    /**
     * @param callable|callable[] $reader
     */
    public function __construct($reader)
    {
        $this->readers = (array) $reader;
    }

    /**
     * @return \Generator[]
     */
    public function makeGenerators()
    {
        $generators = [];

        foreach ($this->readers as $key => $reader) {
            $generators[$key] = $this->makeGenerator($reader);
        }

        return $generators;
    }

    /**
     * @param callable $callback
     * @return \Generator
     */
    protected function makeGenerator(callable $callback)
    {
        $callback = $this->resolveReader($callback);

        $times = 1;

        while ($data = call_user_func($callback, $times)) {
            $times++;

            yield $data;
        }
    }

    /**
     *
     * @param callable $reader
     * @return callable
     */
    protected function resolveReader(callable $reader)
    {
        if (is_callable($reader)) {
            return $reader;
        }

        return function ($times) use ($reader) {
            return $reader($times);
        };
    }
}
