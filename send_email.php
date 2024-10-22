<?php
// Import necessary classes from PHPMailer
use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;
use PHPMailer\PHPMailer\SMTP;

// Include PHPExcel library
require_once 'assets/vendor/PHPExcel.php';
require_once 'assets/vendor/PHPExcel/IOFactory.php';

// Include PHPMailer library files
require_once __DIR__ . '/assets/vendor/phpmailer/src/Exception.php';
require_once __DIR__ . '/assets/vendor/phpmailer/src/PHPMailer.php';
require_once __DIR__ . '/assets/vendor/phpmailer/src/SMTP.php';

// Initialize PHPMailer and enable exception handling
$mail = new PHPMailer(true);

// Check if required form fields are set and not empty
if (
    isset($_POST['name']) && !empty($_POST['name']) &&
    isset($_POST['email']) && !empty($_POST['email']) &&
    isset($_POST['subject']) && !empty($_POST['subject']) &&
    isset($_POST['message']) && !empty($_POST['message'])
) {
    try {
        // Server settings for PHPMailer
        $mail->SMTPDebug = false; // Disable debug output (set to SMTP::DEBUG_SERVER for debugging)
        $mail->isSMTP(); // Set mailer to use SMTP
        $mail->Host = 'smtp.gmail.com'; // Specify main and backup SMTP servers
        $mail->SMTPAuth = true; // Enable SMTP authentication
        $mail->SMTPSecure = PHPMailer::ENCRYPTION_STARTTLS; // Enable TLS encryption
        $mail->Port = 587; // TCP port for Gmail SMTP

        // SMTP credentials for Gmail account
        $mail->Username = 'prathmpatil2818@gmail.com'; // Your Gmail address
        $mail->Password = 'pnxcgsugmethcdzj'; // Your Gmail app password

        // Build email body content from form data
        $mail_body = 'Name: ' . $_POST['name'] . '<br>Email: ' . $_POST['email'] .
                     '<br>Subject: ' . $_POST['subject'] . '<br>Message: ' . $_POST['message'];

        // Sender and recipient settings
        $mail->setFrom('prathmpatil2818@gmail.com', 'Prathmesh Patil'); // Sender's email and name
        $mail->addAddress('prathmpatil2818@gmail.com', 'Prathmesh Patil'); // First recipient
        $mail->addAddress('sanket@gotojiitraders.com', 'Sanket Chougule'); // Second recipient

        // Set email format to HTML
        $mail->isHTML(true);
        $mail->Subject = "Contact Form Details"; // Subject of the email
        $mail->Body = $mail_body; // Email body content in HTML format

        // Path to the existing Excel file
        $inputFileName = 'upload/contact_list.xls';

        // Load the existing Excel file
        $objPHPExcel = PHPExcel_IOFactory::load($inputFileName);

        // Select the active sheet or specify the sheet to read
        $sheet = $objPHPExcel->getActiveSheet();

        // Find the next empty row
        $highestRow = $sheet->getHighestRow();
        $nextRow = $highestRow + 1;

        // Prepare new data to append
        $newData = [
            [$nextRow, $_POST['name'], $_POST['email'], $_POST['subject'], $_POST['message']]
        ];

        // Write new data to the next available row
        foreach ($newData as $rowIndex => $rowData) {
            foreach ($rowData as $colIndex => $cellValue) {
                $sheet->setCellValueByColumnAndRow($colIndex, $nextRow + $rowIndex, $cellValue);
            }
        }

        // Save the updated file
        $outputFileName = 'upload/contact_list.xls';
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->save($outputFileName);

        // Send the email and check if successful
        if ($mail->send()) {
            // Success response
            $data['status_code'] = 200;
            $data['message'] = 'Your message has been sent. Thank you!';
            echo json_encode($data);
        } else {
            // Failure response
            $data['status_code'] = 400;
            $data['message'] = 'Something went wrong. Please try again later.';
            echo json_encode($data);
        }
    } catch (Exception $e) {
        // Exception handling for email sending failure
        $data['status_code'] = 400;
        $data['message'] = "Error in sending email. Mailer Error: {$mail->ErrorInfo}";
        echo json_encode($data);
    }
} else {
    // If form fields are missing or empty, send error response
    $data['status_code'] = 400;
    $data['message'] = 'Please enter all the required details.';
    echo json_encode($data);
}
?>
