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

class Border extends Attribute
{
    public function drawAttribute(Worksheet $sheet, $value, int $rowIndex, string $colIndex, array $options)
    {
        if ($value) {
            if (is_scalar($value)) {
                $value = (array)$value;
            }

            if (!isset($value[1])) {
                $value[1] = StyleBorder::BORDER_THIN;
            }
            if (!isset($value[2])) {
                $value[2] = Color::COLOR_BLACK;
            }

            $style     = [
                'borders' => [
                    $value[0] => [
                        'borderStyle' => $value[1],
                        'color'       => [
                            'argb' => $value[2],
                        ],
                    ],
                ],
            ];
            $cellRange = isset($options['merge']) ? implode(':', $options['merge']) : $colIndex . $rowIndex;
            $sheet->getStyle($cellRange)->applyFromArray($style);
        }
    }
}