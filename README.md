重复工作制造的轮子合集
============

工作中是否重复的需要导出报表？生成特定格式的PDF文件？

如果是的话，那你可以尝试下这个轮子看看是否满足你的需求

## 安装

使用composer安装

~~~
composer require mmmcatoo/simple_tools
~~~

## 1、导出Excel

* 支持头部单行和多行设置
* 支持行高、列宽、对齐方式、边框和字体的设定
* 支持对输入值的回调处理，支持对行高、列宽、文字对齐、边框和字体的单独回调设定

```php
// 简单的调用代码如下：
require_once __DIR__ . '/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;use SimpleTools\Export\Excel\ExportAsExcel;

$header = [
    // 第一行头部数据 若有多行 需要存在多个结构
    [
        [
            // 输出值的字段名称 必须
            'field'    => 'id',
            // 输出值的列名称 必须
            'title'    => '序号',
            // 列宽
            'width'    => 12,
            // 行宽
            'height'   => 20,
            /**
             * 边框属性
             * 第一个为边框方位, 第二个为边框的粗细, 第三个为边框的颜色ARGB格式
             */
            'border'   => ['outline', 'thin', 'FF000000'],
            /**
             * 对齐方式
             * 第一个为水平对齐方式, 第二个为垂直对齐的方式
             */
            'align'    => ['center', 'center'],
            /**
             * 字体设定
             * 第一个为字体名称, 第二个字体大小, 第三个为字体的颜色ARGB格式, 第四个为是否加粗
             */
            'font'     => ['SimSun', 10, 'FF000000', false],
            /**
             * 字体设定回调
             * 如果需要完全覆盖 需要返回当前属性设置值完全合法的格式
             * 如果是修改数组中的某个值 返回 属性值@下标位置
             * @var mixed       $value      单元格结果callback之后的值
             * @var mixed       $rawValue   单元格输入数据值
             * @var string      $attribute  属性名称
             * @var array       $options    头部设定的相关数据 若为多行头部 最后一行头部数据
             * @var string      $colIndex   列名
             * @var int         $rowIndex   行号
             * @var array       $records    输入的整行数据
             * @return string|float|int
             */
            'fontCallback'     => function ($value, $rawValue, string $attribute, array $options, string $colIndex, int $rowIndex, array $records) {
                
            },
            /**
             * 单元格合并
             * 第一个为左上角单元格, 第二个为右小角单元格
             */
            'merge'    => ['A1', 'B1'],
            /**
             * 单元格值回调参数
             * 传入的记录值在写入的时候会被此函数的返回值替代 
             * @var mixed       $rawValue 单元格输入数据值
             * @var array       $records  输入的整行数据
             * @var string      $colIndex 列名
             * @var int         $rowIndex 行号
             * @var array       $options  头部设定的相关数据 若为多行头部 最后一行头部数据
             * @return string|float|int
             */
            'callback' => function ($rawValue, array $records, string $colIndex, int $rowIndex, array $options) {

            },
            /**
             * 单元格格式其他回调
             * @var mixed                                         $value    第一个为单元格的计算值
             * @var mixed                                         $rawValue 第一个为单元格的输入原始值
             * @var string                                        $colIndex 列名
             * @var int                                           $rowIndex 行号
             * @var array                                         $options  头部设定的相关数据 若为多行头部 最后一行头部数据
             * @var \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $sheet    数据表对象
             * @retun void
             */
            'attr'     => function ($value, $rawValue, string $colIndex, int $rowIndex, array $options, Worksheet $sheet) {

            },
        ],
    ],
];

$dataSet = [
    [
        'id'          => 233,
        'account'     => 'Mr.Lee',
        'create_time' => 1585190524,
        'passed'      => 14,
    ],
    [
        'id'          => 234,
        'account'     => 'Mr.Wang',
        'create_time' => 1585190524,
        'passed'      => 10,
    ],
];

$instance = new ExportAsExcel();
// 设置导出的头部信息
$instance->setHeader($header);
// 需要写入的行号
$offset = 2;
// 设置单元格的填充值
$instance->setDataSet($dataSet, $offset);
// 如果数据量太大需要使用PHPSpreadSheet提供的缓存需要执行
// @var \Psr\SimpleCache\CacheInterface @cache
// $instance->setCache($cache);
// 保存路径
$savePath = '';
// 保存文件名
$fileName = '';
// 导出类型 需要为如下值 Xls, Xlsx, Ods, Csv, Html, Tcpdf, Dompdf, Mpdf
$type = 'Xlsx';
// 保存选项
$options = [
    // 冻结开始点 其左上和左边的单元格会冻结
    'freeze' => 'A2',
    // Sheet的名字
    'sheetName' => 'ExportSheet',
    // 如果需要更多的自定义设置 需要实现此回调
    'callback' => function(Worksheet $sheet, Spreadsheet $spreadsheet) {
        
    }
];
/** @var bool $exportPath */
$exportResult = $instance->saveFile($savePath, $fileName, $type, $options);

if ($exportResult === false) {
    // 导出失败会返回错误信息
    echo $instance->getError();
}
```



