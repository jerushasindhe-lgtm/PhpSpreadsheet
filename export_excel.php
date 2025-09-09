<?php
require 'vendor/autoload.php';
include 'db_connect.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Fetch data
$result = $conn->query("SELECT u.name, r.rating, r.feedback, r.created_at 
                        FROM reviews r 
                        JOIN users u ON r.user_id=u.user_id");

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

// Header
$sheet->setCellValue('A1', 'Name');
$sheet->setCellValue('B1', 'Rating');
$sheet->setCellValue('C1', 'Feedback');
$sheet->setCellValue('D1', 'Submitted On');

// Fill data
$rowNum = 2;
while ($row = $result->fetch_assoc()) {
    $sheet->setCellValue("A$rowNum", $row['name']);
    $sheet->setCellValue("B$rowNum", $row['rating']);
    $sheet->setCellValue("C$rowNum", $row['feedback']);
    $sheet->setCellValue("D$rowNum", $row['created_at']);
    $rowNum++;
}

// Export
$writer = new Xlsx($spreadsheet);
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="reviews.xlsx"');
header('Cache-Control: max-age=0');
$writer->save('php://output');
exit;
?>
