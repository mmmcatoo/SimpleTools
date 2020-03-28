<?php
/***
 *
 * Class Merge
 *
 * @author      mmmcatoo<mmmcatoo@qq.com>
 * @version     1.0
 * @package     SimpleTools\Export\Excel\Attributes
 * @date        2020-03-28
 */

namespace SimpleTools\Export\Excel\Attributes;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use SimpleTools\Constraints\Attribute;

class Merge extends Attribute
{
    public function drawAttribute(Worksheet $sheet, $value, int $rowIndex, string $colIndex, array $options)
    {
        if ($value) {
            $sheet->mergeCells(implode(':', $value));
        }
    }

}