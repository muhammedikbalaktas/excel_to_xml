<?php
require '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


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

echo $firstRow[0];

for ($row = 2; $row <= $highestRow; ++$row) {
    $element = $dom->createElement('book');
    for ($col = 1,$index=0; $col <= $highestColumnIndex; $col++,$index++) {
        $value = $worksheet->getCellByColumnAndRow($col, $row)->getValue();
        $key=$firstRow[$index];
        $words = explode(" ", $key); // Split string by space

        $first_word = $words[0];
        $book =$dom->createElement("".$first_word,$value);
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
