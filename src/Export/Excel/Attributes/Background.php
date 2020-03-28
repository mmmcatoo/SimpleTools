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
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use SimpleTools\Constraints\Attribute;

class Background extends Attribute
{
    public function drawAttribute(Worksheet $sheet, $value, int $rowIndex, string $colIndex, array $options)
    {
        if ($value) {
            if ($value !== 'transparent') {
                $sheet
                    ->getStyle($colIndex . $rowIndex)
                    ->getFill()
                    ->setFillType(Fill::FILL_SOLID)
                    ->getStartColor()
                    ->setARGB($value);
            }
        }
    }
}