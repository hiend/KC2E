<?php
namespace KC2E;

use PDO;
use PHPExcel;
use PHPExcel_Exception;
use PHPExcel_Worksheet;
use PHPExcel_Writer_Excel2007;

class Export
{
    /**
     * @var PDO
     */
    private $connect;

    /**
     * @var PHPExcel
     */
    private $excel;

    /**
     * @var PHPExcel_Worksheet
     */
    private $worksheet;

    /**
     * @var PHPExcel_Writer_Excel2007
     */
    private $writer;

    /**
     * @param string $kcdb
     */
    function __construct($kcdb)
    {
        $this->connect = new PDO("sqlite:{$kcdb}");
        $this->excel = new PHPExcel();
        $this->writer = new PHPExcel_Writer_Excel2007($this->excel);
    }

    /**
     * @param string $xsl
     */
    public function save($xsl)
    {
        $this->create_worksheet();

        $this->writer->save($xsl);
    }

    /**
     * @param int $index
     * @param string $keyword_col
     *
     * @throws PHPExcel_Exception
     */
    private function create_worksheet($index = 0, $keyword_col = 'A')
    {
        $this->worksheet = $this->excel->setActiveSheetIndex($index);

        $this->worksheet->getColumnDimension($keyword_col)
            ->setAutoSize(true);

        $this->write_groups();
    }

    /**
     * @param int $row_index
     */
    private function write_groups($row_index = 1)
    {
        foreach ($this->get_groups() as $id => $name) {
            $this->worksheet->setCellValue("A{$row_index}", $name);
            $this->worksheet->getStyle("A{$row_index}")->applyFromArray(array('font' => array('bold' => true)));

            $row_index = $this->write_keywords(++$row_index, $id);
        }
    }

    /**
     * @param int $row_index
     * @param int $group_id
     *
     * @throws PHPExcel_Exception
     *
     * @return mixed
     */
    private function write_keywords($row_index, $group_id)
    {
        foreach ($this->get_keywords($group_id) as $keyword) {
            $this->worksheet->setCellValue("A{$row_index}", $keyword);
            $this->worksheet->getRowDimension($row_index)->setOutlineLevel(1)->setVisible(false);

            $row_index++;
        }

        $this->worksheet->getRowDimension($row_index)->setCollapsed(true);

        return $row_index;
    }

    /**
     * @param int $group_id
     *
     * @return array
     */
    private function get_keywords($group_id)
    {
        $statement = $this->connect->query("SELECT `ID`, `KeyText` FROM `KeyCollector_Keys` WHERE `Tab_ID`={$group_id} ORDER BY `YandexWordstatBaseFreq` DESC");

        return $statement->fetchAll(PDO::FETCH_KEY_PAIR);
    }

    /**
     * @return array
     */
    private function get_groups()
    {
        $statement = $this->connect->query("SELECT `ID`, `TabHeader`, `OwnerId` FROM `KeyCollector_UserTabs` ORDER BY `ID`, `OwnerId`");
        $rows = $statement->fetchAll(PDO::FETCH_ASSOC);
        $groups = [];

        foreach ($rows as $row) {
            $groups[$row['ID']] = $this->get_group_name($rows, $row['OwnerId'], $row['TabHeader']);
        }

        return $groups;
    }

    /**
     * @param array $rows
     * @param int $owner_id
     * @param string $name
     * @param string $delimiter
     *
     * @return string
     */
    private function get_group_name(&$rows, $owner_id, $name, $delimiter = "\\")
    {
        if (is_null($owner_id) or $owner_id == 10) {
            return $name;
        }

        foreach ($rows as $row) {
            if ($row['ID'] == $owner_id) {
                return "{$name} {$delimiter} " . $this->get_group_name($rows, $row['OwnerId'], $row['TabHeader'], $delimiter);
            }
        }

        return $name;
    }
}
