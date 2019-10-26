# Easy Excel

<p>
    <a href="https://github.com/jqhph/easy-excel/blob/master/LICENSE"><img src="https://img.shields.io/badge/license-MIT-7389D8.svg?style=flat" ></a>
    <a href="https://github.com/jqhph/easy-excel/releases" ><img src="https://img.shields.io/github/release/jqhph/easy-excel.svg?color=4099DE" /></a> 
    <a href="https://packagist.org/packages/dcat/easy-excel"><img src="https://img.shields.io/packagist/dt/dcat/easy-excel.svg?color=" /></a> 
    <a><img src="https://img.shields.io/badge/php-7.1+-59a9f8.svg?style=flat" /></a> 
</p>

`Easy Excel`是一个基于 <a href="https://github.com/box/spout" target="_blank">box/spout</a> 封装的Excel读写工具，可以帮助开发者更快速更轻松的读写Excel文件，
并且无论读取多大的文件只需占用极少的内存。

> {tip} 由于`box/spout`只支持读写`xlsx`、`csv`、`ods`等类型文件，所以本项目目前也仅支持读写这三种类型的文件。


## 文档

[文档](https://jqhph.github.io/easyexcel/)

## 环境

- PHP >= 7.1
- PHP extension php_zip
- PHP extension php_xmlreader
- box/spout >= 3.0
- league/flysystem >= 1.0


## 安装

```bash
composer require dcat/easy-excel
```

## License
[The MIT License (MIT)](LICENSE).