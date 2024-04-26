<?php
//call the autoload
require '../../vendor/autoload.php';
//load phpspreadsheet class
use PhpOffice\PhpSpreadsheet\Spreadsheet;
//call xlsx writer class to make an xlsx file
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
//call IOFactory to be able to load existing file
use PhpOffice\PhpSpreadsheet\IOFactory;


$excelfile = 'jahresabo.xlsx';
//load excelfile with IOFactory
$spreadsheet = IOFactory::load($excelfile);
//get current active sheet (first sheet)
$sheet = $spreadsheet->getActiveSheet();

$aboform = $_POST["printoronline"];
$praemie = $_POST["praemie"];
$anrede = $_POST["anrede"];
$titel = $_POST["titel"];
$vorname = $_POST["vorname"];
$nachname = $_POST["nachname"];
$strasseHausnr = $_POST["strassehausnr"];
$plz = $_POST["plz"];
$ort = $_POST["ort"];
$land = $_POST["land"];
$telefon = $_POST["telefon"];
$fax = $_POST["fax"];
$email = $_POST["email"];
$hundebesitzerOderUnternehmen = $_POST["hundebesitzeroderunternehmen"];
$artkundenerwerb = $_POST["wieaufmerksamgeworden"];
$agbsZustimmung = $_POST["agbscheckbox"];
$newsletter = !empty($_POST["newslettercheckbox"]);


$highestRow = $sheet->getHighestRow();
$row = $highestRow +1;
//set values:
$sheet
    ->setCellValue('A'.$row, $aboform)
    ->setCellValue('B'.$row, $praemie)
    ->setCellValue('C'.$row, $anrede)
    ->setCellValue('D'.$row, $titel)
    ->setCellValue('E'.$row, $vorname)
    ->setCellValue('F'.$row, $nachname)
    ->setCellValue('G'.$row, $strasseHausnr)
    ->setCellValue('H'.$row, $plz)
    ->setCellValue('I'.$row, $ort)
    ->setCellValue('J'.$row, $land)
    ->setCellValue('K'.$row, $telefon)
    ->setCellValue('L'.$row, $fax)
    ->setCellValue('M'.$row, $email)
    ->setCellValue('N'.$row, $hundebesitzerOderUnternehmen)
    ->setCellValue('O'.$row, $artkundenerwerb)
    ->setCellValue('P'.$row, $agbsZustimmung)
    ->setCellValue('Q'.$row, $newsletter);

//make an xlsx writer object using above spreadsheet
$writer = new Xlsx($spreadsheet);
//save the file
$writer->save($excelfile);
//phpspreadsheet documentation: https://phpspreadsheet.readthedocs.io/
?>