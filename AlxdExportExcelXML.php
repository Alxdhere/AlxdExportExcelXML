<?php
class AlxdExportExcelXML
{
    private $colCount;
    private $rowCount;
    private $file;
    private $baseFullFilename;
    private $zipFullFilename;
    private $currentRow = array();

    public $exportDir = '/var/www/exports/xml';
    public $exportFile = 'export.xml';


    public function __construct($exportFile, $colCount, $rowCount)
    {
        if (!is_dir($this->exportDir))
            throw new CHttpException(500, 'Invalid export directory \''.$this->exportDir.'\' to export in .xml!');

        if (!preg_match('/^[a-zA-Z\p{Cyrillic}0-9\-\.\[\]\(\)_]+\.xml$/u', $exportFile))
            throw new CHttpException(500, 'Invalid export file name \''.$exportFile.'\' to export in .xml!');

        $this->exportFile = $exportFile;

        $this->baseFullFilename = $this->exportDir.DIRECTORY_SEPARATOR.$this->exportFile;
        if (file_exists($this->baseFullFilename))
            unlink($this->baseFullFilename);

        $this->colCount = $colCount;
        $this->rowCount = $rowCount;
    }

    function __destruct()
    {
        unset($this->exportDir, $this->exportFile, $this->baseFullFilename, $this->currentRow, $this->colCount, $this->rowCount, $this->file);
    }

    public function getColCount()
    {
        return $this->colCount;
    }

    public function getRowCount()
    {
        return $this->rowCount;
    }

    public function getBaseFullFileName()
    {
        return $this->baseFullFilename;
    }

    public function getZipFullFileName()
    {
        return $this->zipFullFilename;
    }

    public function openWriter()
    {
        $this->file = fopen($this->baseFullFilename, 'w+');
        fwrite($this->file, '<?xml version="1.0"?>');
        fwrite($this->file, '<?mso-application progid="Excel.Sheet"?>');

    }

    public function openWorkbook()
    {
        fwrite($this->file, '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:html="http://www.w3.org/TR/REC-html40">');
    }

    public function writeDocumentProperties($organization = null, $user = null)
    {
        fwrite($this->file, '<DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">');
        if (!is_null($user))
        {
            fwrite($this->file, '<Author>'.$user->description.'</Author>');
            fwrite($this->file, '<LastAuthor>'.$user->description.'</LastAuthor>');
        }

        $dt = new Datetime();
        $dt_string = $dt->format('Y-m-d\TH:i:s\Z');
        fwrite($this->file, '<Created>'.$dt_string.'</Created>');
        fwrite($this->file, '<LastSaved>'.$dt_string.'</LastSaved>');

        if (!is_null($organization))
            fwrite($this->file, '<Company>'.$organization->name.'</Company>');

        fwrite($this->file, '<Version>12.00</Version>');
        fwrite($this->file, '</DocumentProperties>');
    }

    public function writeStyles()
    {
        fwrite($this->file, '<Styles>');
        //default style
        fwrite($this->file, '<Style ss:ID="Default" ss:Name="Normal"><Font ss:Color="#000000"/></Style>');
        //Datetime style
        fwrite($this->file, '<Style ss:ID="DateTime"><NumberFormat ss:Format="General Date"/></Style>');
        fwrite($this->file, '<Style ss:ID="Date"><NumberFormat ss:Format="Short Date"/></Style>');
        fwrite($this->file, '<Style ss:ID="Time"><NumberFormat ss:Format="h:mm:ss"/></Style>');
        //Hyperlink style
        fwrite($this->file, '<Style ss:ID="Hyperlink" ss:Name="Hyperlink"><Font ss:Color="#0000FF" ss:Underline="Single"/></Style>');
        //Bold
        fwrite($this->file, '<Style ss:ID="Bold"><Font ss:Bold="1"/></Style>');


    fwrite($this->file, '</Styles>');
    }

    public function openWorksheet()
    {
        fwrite($this->file, '<Worksheet ss:Name="Export">');
        fwrite($this->file, strtr('<Table ss:ExpandedColumnCount="{col_count}" ss:ExpandedRowCount="{row_count}" x:FullColumns="1" x:FullRows="1" ss:DefaultRowHeight="15">', array('{col_count}'=>$this->colCount, '{row_count}'=>$this->rowCount)));
    }

    public function resetRow()
    {
        $this->currentRow = array();
    }

    public function openRow($isBold = false)
    {
        $this->currentRow[] = '<Row ss:AutoFitHeight="0"'.($isBold ? ' ss:StyleID="Bold"' : '').'>';
    }

    public function closeRow()
    {
        $this->currentRow[] = '</Row>';
    }

    public function flushRow()
    {
        fwrite($this->file, implode('', $this->currentRow));
        unset($this->currentRow);
    }

    public function appendCellNum($value)
    {
        $this->currentRow[] = '<Cell><Data ss:Type="Number">'.$value.'</Data></Cell>';
    }

    public function appendCellString($value)
    {
        $this->currentRow[] = '<Cell><Data ss:Type="String">'.htmlspecialchars($value).'</Data></Cell>';
    }

    public function appendCellReal($value)
    {
        return $this->appendCellNum($value);
    }

    public function appendCellDateTime($value)
    {
        if (empty($value))
            $this->appendCellString('');
        else
            $this->currentRow[] = '<Cell ss:StyleID="DateTime"><Data ss:Type="DateTime">'.$value.'</Data></Cell>';
    }

    public function appendCellDate($value)
    {
        if (empty($value))
            $this->appendCellString('');
        else
            $this->currentRow[] = '<Cell ss:StyleID="Date"><Data ss:Type="DateTime">'.$value.'</Data></Cell>';
    }

    public function appendCellTime($value)
    {
        if (empty($value))
            $this->appendCellString('');
        else
            $this->currentRow[] = '<Cell ss:StyleID="Time"><Data ss:Type="DateTime">'.$value.'</Data></Cell>';
    }

    public function appendCellLink($value,$link)
    {
        if (empty($value))
            $this->appendCellString('');
        else
            $this->currentRow[] = '<Cell ss:StyleID="Hyperlink" ss:HRef="'.$link.'"><Data ss:Type="String">'.$value.'</Data></Cell>';
    }

    public function closeWorksheet()
    {
        fwrite($this->file, '</Table>');
        fwrite($this->file, '<AutoFilter x:Range="R1C1:R'.$this->rowCount.'C'.$this->colCount.'" xmlns="urn:schemas-microsoft-com:office:excel"></AutoFilter>');
        fwrite($this->file, '</Worksheet>');
    }

    public function closeWorkbook()
    {
        fwrite($this->file, '</Workbook>');
    }

    public function closeWriter()
    {
        fclose($this->file);
    }

    public function zip()
    {
        $zipfile = trim(pathinfo('"'.$this->exportFile.'"', PATHINFO_FILENAME).'.zip','"');
        $curDir = getcwd();
        chdir($this->exportDir);
        exec('zip -m "'.$zipfile.'" "'.$this->exportFile.'"');
        chdir($curDir);

        $this->zipFullFilename = $this->exportDir.DIRECTORY_SEPARATOR.$zipfile;
    }
}
?>
