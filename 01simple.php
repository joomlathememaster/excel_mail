<?php
/**
 * PHPExcel
 *
 * Copyright (C) 2006 - 2014 PHPExcel
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPExcel
 * @package    PHPExcel
 * @copyright  Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    1.8.0, 2014-03-02
 */

/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

/** Include PHPExcel */
require_once dirname(__FILE__) . '/Classes/PHPExcel.php';
require __DIR__ .'/vendor/autoload.php';

$faker = Faker\Factory::create();

// Create new PHPExcel object
$objPHPExcel = new PHPExcel();

$alphas=array('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z');
// Set document properties (todo last)
/**
$objPHPExcel->getProperties()->setCreator("Maarten Balliauw")
							 ->setLastModifiedBy("Maarten Balliauw")
							 ->setTitle("PHPExcel Test Document")
							 ->setSubject("PHPExcel Test Document")
							 ->setDescription("Test document for PHPExcel, generated using PHP classes.")
							 ->setKeywords("office PHPExcel php")
							 ->setCategory("Test result file");
*/
//fake data product
$products=array();
for($i=0; $i<50; $i++){
    $product=array(
        'id'=>$faker->randomNumber(2),
        'title'=>$faker->word,
        'image'=>$faker->imageUrl($width = 640, $height = 480),
        'description'=>$faker->text,
        'number'=>$faker->randomNumber(1),
        'price'=>number_format($faker->randomNumber(2),2) ,
    );
    $products[]=(object)$product;
}
//fake data order
$orders=array();
for($i=0; $i<50; $i++){
    $order=array(
        'id'=>$faker->randomNumber(2),
        'product_id'=>$faker->randomNumber,
        'customer_name'=>$faker->name,
        'total'=>number_format($faker->randomNumber(2),2) ,
        'date_created'=>date('Y-m-d'),
    );
    $orders[]=(object)$order;
}
// add tables as table key in array
$fake_array=array(
    'product' => $products,
    'order' => $orders
);
$sheet_count = 0; 

/** Add data for product (id,title,image,description,number,price)
    Add data for order (id,product_id,customer_name,total,date_created)
 */
// Rename worksheet
foreach($fake_array as $key => $value){
    if($sheet_count>0){  
        $objPHPExcel->createSheet();
        $objPHPExcel->setActiveSheetIndex($sheet_count);
    }
    $objPHPExcel->getActiveSheet()->setTitle($key);
    // Sheet heading in the first row
    $key_array = (array)$value[0]; 
    $k=0;
    foreach($key_array as $field_name=>$field_value){ 
        $objPHPExcel->getActiveSheet()->setCellValue($alphas[$k].'1', $field_name);
        $k++;
    }
    
    // sheet data in the next rows from first row 
    $row = 1;
    foreach($value as $key1 => $value1){
        $row++;
        $data_array = (array)$value1;
        $t=0;
        foreach($data_array as $field_data=>$field_data_value){ 
            $objPHPExcel->getActiveSheet()->setCellValue($alphas[$t].$row, $field_data_value);
            $t++;
        }
    
    }
    $sheet_count++;
}  
 

// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$objPHPExcel->setActiveSheetIndex(0);
         
// Save Excel 95 file
$callStartTime = microtime(true);

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save(str_replace('.php', '.xls', __FILE__));
$callEndTime = microtime(true);
$callTime = $callEndTime - $callStartTime;

//---------------------------- //

class mailer {

    function mail_attachment($filename, $path, $mailto, $from_mail, $from_name, $replyto, $subject, $message) {
         $file = $path.$filename;
        $file_size = filesize($file);
        $handle = fopen($file, "r");
        $content = fread($handle, $file_size);
        fclose($handle);
        $content = chunk_split(base64_encode($content));
        $uid = md5(uniqid(time()));
        $header = "From: ".$from_name." <".$from_mail.">\r\n";
        $header .= "Reply-To: ".$replyto."\r\n";
        $header .= "MIME-Version: 1.0\r\n";
        $header .= "Content-Type: multipart/mixed; boundary=\"".$uid."\"\r\n\r\n";
        $header .= "This is a multi-part message in MIME format.\r\n";
        $header .= "--".$uid."\r\n";
        $header .= "Content-type:text/plain; charset=iso-8859-1\r\n";
        $header .= "Content-Transfer-Encoding: 7bit\r\n\r\n";
        $header .= $message."\r\n\r\n";
        $header .= "--".$uid."\r\n";
        $header .= "Content-Type: application/octet-stream; name=\"".$filename."\"\r\n"; // use different content types here
        $header .= "Content-Transfer-Encoding: base64\r\n";
        $header .= "Content-Disposition: attachment; filename=\"".$filename."\"\r\n\r\n";
        $header .= $content."\r\n\r\n";
        $header .= "--".$uid."--";
        if (mail($mailto, $subject, "", $header)) {
            echo "mail send ... OK"; // or use booleans here
        } else {
            echo "mail send ... ERROR!";
        } 
    }

}
$mailer= new mailer;
 
$my_file = "01simple.xls";
$my_path = "/home/devjoom/public_html/excel_email/";
$my_name = "admin";
$my_mail = "huanquangchu@gmail.com";
$my_replyto = "huanquangchu@gmail.com";
$my_subject = "This is a mail with attachment.";
$my_message = "Helo,\r\ndo you like this script? I hope it will help.\r\n\r\ngr. Olaf";
 
$mailer->mail_attachment($my_file, $my_path, "knightdev86@gmail.com,expertertn@gmail.com,nguyenhieptn@yahoo.com", $my_mail, $my_name, $my_replyto, $my_subject, $my_message);
 