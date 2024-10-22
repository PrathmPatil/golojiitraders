<?php

// require 'vendor/autoload.php'; 
// Include PhpSpreadsheet library

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Check if the form is submitted
if ($_SERVER["REQUEST_METHOD"] == "POST") {
    
    // Get form data
    $name = $_POST['name'];
    $email = $_POST['email'];
    $message = $_POST['message'];

    // Create a new Spreadsheet
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    // Set spreadsheet column headers
    $sheet->setCellValue('A1', 'Name');
    $sheet->setCellValue('B1', 'Email');
    $sheet->setCellValue('C1', 'Message');

    // Insert form data into the Excel sheet
    $sheet->setCellValue('A2', $name);
    $sheet->setCellValue('B2', $email);
    $sheet->setCellValue('C2', $message);

    // Set the path where the Excel file will be saved
    $filePath = 'form_data.xlsx';

    // Write the Excel file to the given path
    $writer = new Xlsx($spreadsheet);

    try {
        $writer->save($filePath);
        echo "Data saved to Excel file successfully!";
    } catch (Exception $e) {
        echo "Error saving Excel file: ", $e->getMessage();
    }
} else {
    echo "No data received.";
}