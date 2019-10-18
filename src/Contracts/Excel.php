<?php

namespace Dcat\EasyExcel\Contracts;

/**
 * @method $this xlsx()
 * @method $this csv()
 * @method $this ods()
 */
interface Excel
{
    /**
     * @param \Closure $callback
     * @return $this
     */
    public function option(\Closure $callback);

    /**
     * @param array $headers
     * @param \Closure|null $callback
     * @return $this
     */
    public function headers(array $headers, \Closure $callback = null);

    /**
     * @return array|false
     */
    public function getHeaders();

    /**
     * @return $this
     */
    public function withoutHeaders();

    /**
     * @param string $type
     * @return $this
     */
    public function type(string $type);

    /**
     * @return string|null
     */
    public function getType();

}
