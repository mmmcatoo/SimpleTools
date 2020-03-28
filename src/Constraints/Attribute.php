<?php
/***
 *
 * Class Attribute
 *
 * @author      mmmcatoo<mmmcatoo@qq.com>
 * @version     1.0
 * @package     Constraints
 * @date        2020-03-28
 */

namespace SimpleTools\Constraints;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

abstract class Attribute
{
    /**
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $sheet
     * @param                                               $value
     * @param int                                           $rowIndex
     * @param string                                        $colIndex
     * @param array                                         $options
     * @return mixed
     * @throws \Exception
     */
    public abstract function drawAttribute(Worksheet $sheet, $value, int $rowIndex, string $colIndex, array $options);
}