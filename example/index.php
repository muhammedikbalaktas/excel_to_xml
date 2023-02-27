<?php
require '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
// ini_set('display_errors', 1);
// error_reporting(E_ALL);

$reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader('Xlsx');
$reader->setReadDataOnly(TRUE);
$spreadsheet = $reader->load("../excels/ex.xlsx");
$dom = new DOMDocument('1.0');
$dom->encoding = 'UTF-8';
$worksheet = $spreadsheet->getActiveSheet();
$firstRow = $worksheet->toArray()[0];

$highestRow = $worksheet->getHighestRow(); 
$highestColumn = $worksheet->getHighestColumn(); 
$highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn); // e.g. 5

$root = $dom->createElement('books');
$dom->appendChild($root);



for ($row = 2; $row <= $highestRow; ++$row) {
    $element = $dom->createElement('book');
    for ($col = 1,$index=0; $col <= $highestColumnIndex; $col++,$index++) {
        $value = $worksheet->getCellByColumnAndRow($col, $row)->getValue();
        $key=$firstRow[$index];
        
        $final_key = str_replace(' ', '', $key);
        $final_key = str_replace('(', '-', $final_key);
        $final_key = str_replace(')', '-', $final_key);
        
        
        $book =$dom->createElement($final_key,$value);
        $element->appendChild($book);
        
        
    }
     $root->appendChild($element);
}

$dom->appendChild($root);



$output_file = 'output.xml';


$dom->preserveWhiteSpace = false;
$dom->formatOutput = true;


try {
    $dom->save($output_file);
    echo 'XML file generated successfully!';
} catch (Exception $e) {
    echo 'Error: ' . $e->getMessage();
}
