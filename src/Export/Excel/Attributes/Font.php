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

class Font extends Attribute
{
    public function drawAttribute(Worksheet $sheet, $value, int $rowIndex, string $colIndex, array $options)
    {
        if ($value) {
            if (is_scalar($value)) {
                $value = (array) $value;
            }

            $font = $sheet->getStyle($colIndex . $rowIndex)->getFont();
            $font->setName($value[0] ?? 'SimSun');
            $font->setSize($value[1] ?? 12);
            $font->getColor()->setARGB($value[2] ?? Color::COLOR_BLACK);
            $font->setBold($value[3] ?? false);
        }
    }
}