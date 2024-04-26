<?php
//call the autoload
require '../../vendor/autoload.php';
//load phpspreadsheet class
use PhpOffice\PhpSpreadsheet\Spreadsheet;
//call xlsx writer class to make an xlsx file
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
//call IOFactory to be able to load existing file
use PhpOffice\PhpSpreadsheet\IOFactory;

$excelfile = 'schnupperabo.xlsx';
//load excelfile with IOFactory
$spreadsheet = IOFactory::load($excelfile);
//get current active sheet (first sheet)
$sheet = $spreadsheet->getActiveSheet();

$bestellwunsch = $_POST["bestellwunsch"];
$anrede = $_POST["anrede"];
$titel = $_POST["titel"];
$vorname = $_POST["vorname"];
$nachname = $_POST["nachname"];
$strasseHausnr = $_POST["strassehausnr"];
$plz = $_POST["plz"];
$ort = $_POST["ort"];
$land = $_POST["land"];
$telefon = $_POST["telefon"];
$email = $_POST["email"];
$hundebesitzerOderUnternehmen = $_POST["hundebesitzeroderunternehmen"];
$andereHaustiere = $_POST["anderehaustiere"];
$agbsZustimmung = $_POST["agbscheckbox"];


$highestRow = $sheet->getHighestRow();
$row = $highestRow +1;
//set values:
$sheet
    ->setCellValue('A'.$row, $bestellwunsch)
    ->setCellValue('B'.$row, $anrede)
    ->setCellValue('C'.$row, $titel)
    ->setCellValue('D'.$row, $vorname)
    ->setCellValue('E'.$row, $nachname)
    ->setCellValue('F'.$row, $strasseHausnr)
    ->setCellValue('G'.$row, $plz)
    ->setCellValue('H'.$row, $ort)
    ->setCellValue('I'.$row, $land)
    ->setCellValue('J'.$row, $telefon)
    ->setCellValue('K'.$row, $email)
    ->setCellValue('L'.$row, $hundebesitzerOderUnternehmen)
    ->setCellValue('M'.$row, $andereHaustiere)
    ->setCellValue('N'.$row, $agbsZustimmung);

//make an xlsx writer object using above spreadsheet
$writer = new Xlsx($spreadsheet);
//write the file in current directory
$writer->save($excelfile);
//phpspreadsheet documentation: https://phpspreadsheet.readthedocs.io/
?>