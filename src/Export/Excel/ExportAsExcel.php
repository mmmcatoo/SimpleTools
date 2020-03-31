<?php
/***
 *
 * Class ExportAsExcel
 *
 * @author      mmmcatoo<mmmcatoo@qq.com>
 * @version     1.0
 * @package     Export\Excel
 * @date        2020-03-28
 */

namespace SimpleTools\Export\Excel;

use Closure;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Settings;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use Psr\SimpleCache\CacheInterface;
use SimpleTools\Exception\Export\HeaderOptionsEmptyException;
use SimpleTools\Exception\Export\SavePathAccessDeniedException;
use SimpleTools\Exception\Export\SavePathCreateFailedException;
use Throwable;

class ExportAsExcel
{
    /**
     * 错误信息
     *
     * @var string
     */
    protected $errorMessage = '';

    /**
     * Excel头部数据
     *
     * @var array
     */
    protected $header = [];

    /**
     * Excel的内容
     *
     * @var array
     */
    protected $dataSet = [];

    /**
     * Excel对象
     *
     * @var \PhpOffice\PhpSpreadsheet\Spreadsheet
     */
    protected $spreedSheet = null;

    /**
     * WorkSheet对象
     *
     * @var \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    protected $workSheet = null;

    /**
     * 缓存接口
     *
     * @var CacheInterface
     */
    protected $cacheInstance = null;

    /**
     * 列的相关信息
     *
     * @var array
     */
    protected $columnsData = [];

    /**
     * 正在绘制的类型
     *
     * @var string
     */
    protected $drawingType = '';

    /**
     * 目前可以支持的属性
     *
     * @var array
     */
    protected $attributes = ['width', 'height', 'border', 'align', 'font', 'background', 'merge'];

    /**
     * 属性的实例对象缓存
     *
     * @var \SimpleTools\Constraints\Attribute[]
     */
    protected $attributesCache = [];

    /**
     * 开始导入数据的行号
     *
     * @var int
     */
    protected $startOffset;

    /**
     * ExportAsExcel constructor.
     */
    public function __construct()
    {

    }

    /**
     * 设置头部信息
     *
     * @param array $header
     * @return \SimpleTools\Export\Excel\ExportAsExcel
     * @throws void
     */
    public function setHeader(array $header): ExportAsExcel
    {
        $this->header = $header;
        return $this;
    }

    /**
     * 导入要写入的单元格数据
     *
     * @param array $dataSet
     * @param int   $startOffset
     * @return \SimpleTools\Export\Excel\ExportAsExcel
     * @throws void
     */
    public function setDataSet(array $dataSet, int $startOffset): ExportAsExcel
    {
        $this->dataSet     = $dataSet;
        $this->startOffset = $startOffset;
        return $this;
    }

    /**
     * @param \Psr\SimpleCache\CacheInterface $cache
     * @return \SimpleTools\Export\Excel\ExportAsExcel
     * @throws void
     */
    public function setCache(CacheInterface $cache): ExportAsExcel
    {
        $this->cacheInstance = $cache;
        return $this;
    }

    public function saveFile(string $savePath, string $fileName, string $type, array $sheetOptions = []): bool
    {
        try {
            if (count($this->header) <= 0) {
                throw new HeaderOptionsEmptyException('生成的Excel的头部不能为空');
            }
            // 创建数据表对象
            $this->createWorkSheet($sheetOptions['sheetName'] ?? '');
            // 检测文件夹是否存在和可读
            if (!is_dir($savePath) && !(mkdir($savePath, 0777, true))) {
                throw new SavePathCreateFailedException($savePath . '不存在并且无法创建');
            }
            if (!is_writable($savePath)) {
                throw new SavePathAccessDeniedException($savePath . '无法写入');
            }

            // 如果需要冻结某些行 提前处理
            if (isset($sheetOptions['freeze'])) {
                $this->workSheet->freezePane($sheetOptions['freeze']);
            }

            // 对导出的Sheet进行自定义设置
            if (isset($sheetOptions['callback']) && ($sheetOptions['callback'] instanceof Closure)) {
                $sheetOptions['callback']($this->workSheet, $this->spreedSheet);
            }
            // 开始绘制头部数据
            $this->drawHeaderCell();
            // 绘制单元格数据
            if (count($this->dataSet)) {
                foreach ($this->dataSet as $offset => $item) {
                    foreach ($this->columnsData as $colIndex => $options) {
                        $this->drawCell($colIndex, $offset + $this->startOffset, $options, $item);
                    }
                }
            }
            // 保存文件
            $fileName = sprintf('%s%s%s.%s', $savePath, DIRECTORY_SEPARATOR, $fileName, strtolower($type));
            $fileName = str_replace(['//', DIRECTORY_SEPARATOR . DIRECTORY_SEPARATOR], DIRECTORY_SEPARATOR, $fileName);
            // 创建导出对象
            $writer = IOFactory::createWriter($this->spreedSheet, $type);
            $writer->save($fileName);
            return true;
        } catch (Throwable $e) {
            $this->errorMessage = $e->getMessage();
            return false;
        }
    }

    /**
     * @return string
     * @return void
     * @throws void
     */
    public function getError(): string
    {
        return $this->errorMessage;
    }

    /**
     * @param string $sheetName
     * @return void
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    protected function createWorkSheet(string $sheetName)
    {
        // 如果设置了缓存实例 先赋值
        if ($this->cacheInstance) {
            Settings::setCache($this->cacheInstance);
        }

        // 创建数据表信息
        $this->spreedSheet = new Spreadsheet();
        $this->workSheet   = $this->spreedSheet->getActiveSheet();
        if ($sheetName) {
            $this->workSheet->setTitle($sheetName);
        }
    }

    /**
     * 绘制头部单元格
     * @param void
     * @return void
     * @throws \Exception
     */
    protected function drawHeaderCell()
    {
        $this->drawingType = 'HEAD';
        // 检查是否只有一行或是多行头部的情况
        $loopHeaderOptions = isset($this->header[0][0]) ? $this->header : [$this->header];
        foreach ($loopHeaderOptions as $offset => $option) {
            foreach ($option as $colIndex => $item) {
                // 转换数字到字母列名
                $colIndex = $this->getPrefix($colIndex);
                // 缓存本列的信息
                $this->columnsData[$colIndex] = $item;
                // 绘制单元格
                $this->drawCell($colIndex, $offset + 1, $item, [$item['field'] => $item['title']]);
            }
        }
        $this->drawingType = 'CELL';
    }

    /**
     * 将头部的数字索引转换为ABC这种Excel索引
     *
     * @param int $colIndex
     * @return string
     * @throws void
     */
    protected function getPrefix(int $colIndex)
    {
        $mod    = floor($colIndex / 26);
        $prefix = $mod >= 26 ? $this->getPrefix($mod) : '';
        return $prefix . chr($colIndex % 26 + 65);
    }

    /**
     * @param string $colIndex
     * @param int    $rowIndex
     * @param array  $options
     * @param array  $records
     * @throws \Exception
     */
    protected function drawCell(string $colIndex, int $rowIndex, array $options, array $records)
    {
        $rawValue = $records[$options['field']] ?? '';
        // 绘制数据单元格需要读取参数回调
        if ($this->drawingType === 'CELL') {
            // 如果回调是个闭包
            if (isset($options['callback']) && ($options['callback'] instanceof Closure)) {
                $rawValue = $options['callback']($rawValue, $records, $colIndex, $rowIndex, $options);
            }
        }

        // 设置单元格文本
        $this->workSheet->setCellValue($colIndex . $rowIndex, $rawValue);

        foreach ($this->attributes as $attribute) {
            if ($this->drawingType === 'CELL') {
                // 绘制单元格的时候 可以设置回调值干预属性
                if (array_key_exists($attribute . 'Callback', $options)) {
                    $modifyData = $options[$attribute . 'Callback']($rawValue, $records[$options['field']] ?? '', $attribute, $options, $rowIndex, $colIndex, $records);
                    if ($modifyData) {
                        if (is_string($modifyData) && strpos($modifyData, '@') !== false) {
                            // 替换部分属性
                            [$value, $pos] = explode('@', $modifyData);
                            $options[$attribute][$pos] = $value;
                        } else {
                            // 替换全部属性
                            $options[$attribute] = $modifyData;
                        }
                    }
                }
            }

            // 绘制单元格属性
            $this->getDrawAttribute($options, $attribute, $rowIndex, $colIndex);
        }

        /**
         * 复杂的格式需要自行处理
         */
        if (isset($options['attr']) && ($options['attr'] instanceof Closure)) {
            $options['attr']($rawValue, $records[$options['field']] ?? '', $colIndex, $rowIndex, $options, $this->workSheet);
        }
    }

    /**
     * @param array  $options
     * @param string $attribute
     * @param int    $rowIndex
     * @param string $colIndex
     * @throws \Exception
     */
    protected function getDrawAttribute(array $options, string $attribute, int $rowIndex, string $colIndex)
    {
        $class = '\\SimpleTools\\Export\\Excel\\Attributes\\' . ucfirst(strtolower($attribute));
        if (!isset($this->attributesCache[$class])) {
            if (!class_exists($class)) {
                return;
            }
            $this->attributesCache[$class] = new $class;
        }
        $instance = $this->attributesCache[$class];
        $instance->drawAttribute($this->workSheet, $options[$attribute] ?? null, $rowIndex, $colIndex, $options);
    }
}