<?php
/**
 * User: @gabidj
 * Date: 10/26/2017
 * Time: 7:06 PM
 */

namespace GSRO\SpreadsheetHandler;


use PhpOffice\PhpSpreadsheet\Worksheet;

class SpreadsheetHandler
{
    /**
     * @var Worksheet
     */
    protected $worksheet;
    
    /**
     * @return Worksheet
     */
    public function getWorksheet(): Worksheet
    {
        return $this->worksheet;
    }
    
    /**
     * @param Worksheet $worksheet
     */
    public function setWorksheet(Worksheet $worksheet)
    {
        $this->worksheet = $worksheet;
    }
    
    /**
     * @var string
     */
    protected $currentPosition = 'A1';
    
    /**
     * @return string
     */
    public function getCurrentPosition(): string
    {
        return $this->currentPosition;
    }
    
    /**
     * @param string $currentPosition
     */
    public function setCurrentPosition(string $currentPosition)
    {
        $this->currentPosition = $currentPosition;
    }
    
    public function __construct(Worksheet $worksheet)
    {
        $this->worksheet = $worksheet;
        $pos = explode(':', $this->worksheet->getSelectedCells());
        $this->currentPosition = $pos[-1+count($pos)];
    }
    
    public function writeRow(array $row, $offset = 'A', array $style = [])
    {
        $currentRow = filter_var($this->currentPosition, FILTER_SANITIZE_NUMBER_INT);
        $currentCol = $offset;
        foreach ($row as $value) {
            $cell = $this->worksheet->getCell($currentCol.$currentRow);
            $cell->setValue($value);
            $cell->getStyle()->applyFromArray($style);
            $currentCol++;
        }
        $currentRow++;
        $this->currentPosition = $currentCol.$currentRow;
        return $this;
    }
    
    public function writeAssoc($array, $hasHeader = true, array $style = [], $offset ='A', array $headerStyle = [])
    {
        if ($hasHeader) {
            $header = array_keys($headerStyle);
            $this->writeRow($header, $offset, $headerStyle);
        }
        foreach ($array as $row) {
            $this->writeRow($row, $offset, $style);
        }
        return $this;
    }
    
    /**
     * Read row without affecting current position
     * @param $rowIndex
     * @return array
     */
    public function readRow($rowIndex, $offset = 'A')
    {
        $currentRow = $rowIndex;
        $currentCol = $offset;
        while ($value = $this->worksheet->getCell($currentCol.$currentRow)->getValue()) {
            $currentCol++;
            $values[] = $value;
        };
        return $values;
    }
    
    public function readRowWithHeaders($headers, $rowIndex, $offset = 'A')
    {
        $values = [];
        $currentRow = $rowIndex;
        $currentCol = $offset;
    
        foreach ($headers as $header) {
            $value = $this->worksheet->getCell($currentCol.$currentRow)->getValue();
            $currentCol++;
            $values[$header] = $value;
        };
        return $values;
    
    }
    
    
    
    public function readNextRow(string $position = null, $offset = 'A')
    {
        $values = [];
        $currentRow = filter_var($this->currentPosition, FILTER_SANITIZE_NUMBER_INT);
        $currentCol = $offset;
    
        $values = $this->readRow($currentRow, $currentCol);
    
        $currentRow++;
        $this->setCurrentPosition($currentCol.$currentRow);
    
        return $values;
    }
    
    public function readNextRowByHeaders(array $headers = [], $offset = 'A')
    {
        $values = [];
        $currentRow = filter_var($this->currentPosition, FILTER_SANITIZE_NUMBER_INT);
        $currentCol = $offset;
    
        $values = $this->readRowWithHeaders($headers, $currentRow, $currentCol);
        
        $currentRow++;
        $this->setCurrentPosition($currentCol.$currentRow);
        
        return $values;
    }   
}
