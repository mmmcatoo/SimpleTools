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

use PhpOffice\PhpSpreadsheet\Style\Border as StyleBorder;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use SimpleTools\Constraints\Attribute;

class Align extends Attribute
{
    public function drawAttribute(Worksheet $sheet, $value, int $rowIndex, string $colIndex, array $options)
    {
        if ($value) {
            if (is_string($value)) {
                $value = (array) $value;
            }

            $style = [];
            if (isset($value[0])) {
                $style['alignment']['horizontal'] = $value[0];
            }
            if (isset($value[1])) {
                $style['alignment']['vertical'] = $value[1];
            }

            if (count($style)) {
                $cellRange = isset($options['merge']) ? implode(':', $options['merge']) : $colIndex . $rowIndex;
                $sheet->getStyle($cellRange)->applyFromArray($style);
            }
        }
    }
}