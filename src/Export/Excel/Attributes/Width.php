<?php
/***
 *
 * Class Width
 *
 * @author      mmmcatoo<mmmcatoo@qq.com>
 * @version     1.0
 * @package     Export\Excel\Attributes
 * @date        2020-03-28
 */

namespace SimpleTools\Export\Excel\Attributes;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use SimpleTools\Constraints\Attribute;

class Width extends Attribute
{
    public function drawAttribute(Worksheet $sheet, $value, int $rowIndex, string $colIndex, array $options)
    {
        if ($value > 0) {
            $sheet->getColumnDimension($colIndex)->setWidth($value);
        } else {
            $sheet->getColumnDimension($colIndex)->setAutoSize(true);
        }
    }
}