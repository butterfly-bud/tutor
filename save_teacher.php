<?php
require 'vendor/autoload.php'; // Load PhpSpreadsheet

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Excel file path
$file = "teachers_data.xlsx";

if ($_SERVER["REQUEST_METHOD"] == "POST") {
    $name = $_POST["teacherName"];
    $age = $_POST["teacherAge"];
    $notes = $_POST["notes"];

    // Check if file exists
    if (file_exists($file)) {
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($file);
        $sheet = $spreadsheet->getActiveSheet();
        $row = $sheet->getHighestRow() + 1;
    } else {
        // Create new file if it doesn't exist
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setCellValue("A1", "Teacher Name");
        $sheet->setCellValue("B1", "Age");
        $sheet->setCellValue("C1", "Notes");
        $row = 2;
    }

    // Insert new data
    $sheet->setCellValue("A$row", $name);
    $sheet->setCellValue("B$row", $age);
    $sheet->setCellValue("C$row", $notes);

    // Save the file
    $writer = new Xlsx($spreadsheet);
    $writer->save($file);

    echo "✅ Data saved successfully in Excel!";
} else {
    echo "❌ Invalid request!";
}
?>
