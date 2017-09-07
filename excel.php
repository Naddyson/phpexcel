<?php
use app\models\Provider;

// Подключаем класс для работы с excel
require_once('../vendor/phpoffice/phpexcel/Classes/PHPExcel.php');
// Подключаем класс для вывода данных в формате excel
require_once('../vendor/phpoffice/phpexcel/Classes/PHPExcel/Writer/Excel5.php');



// Создаем объект класса PHPExcel
$xls = new PHPExcel();
// Устанавливаем индекс активного листа
$xls->setActiveSheetIndex(0);
// Получаем активный лист
$sheet = $xls->getActiveSheet();
// Подписываем лист
$sheet->setTitle('Отчет по поиску поставщика ');


$sheet->mergeCells('B1:E8'); //logo
$sheet->mergeCells('F1:S8'); //name

$sheet->mergeCells('B9:S9'); //red
setRedColor($sheet,"B9");

$sheet->mergeCells('B10:E10'); $sheet->mergeCells('F10:S10'); //клиент
$sheet->mergeCells('B11:E11'); $sheet->mergeCells('F11:S11'); // номер заказа
$sheet->mergeCells('B12:E12'); $sheet->mergeCells('F12:S12'); //Наименование услуги
$sheet->mergeCells('B13:E13'); $sheet->mergeCells('F13:S13'); // Дата составления отчета
$sheet->setCellValue("B10",'Клиент');
$sheet->setCellValue("B11",'Номер заказа');
$sheet->setCellValue("B12",'Наименование услуги');
$sheet->setCellValue("B13",'Дата составления отчета');

function setRedColor($sheet, $cell){
    $sheet->getStyle($cell)->applyFromArray(
        array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => 'FF0000')
            )
        )
    );
}


function textCenter($sheet, $cell){
    $sheet->getStyle($cell)->getAlignment()->setHorizontal(
        PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $sheet->getStyle($cell)->getAlignment()->setHorizontal(
        PHPExcel_Style_Alignment::VERTICAL_CENTER);
}



$sheet->mergeCells('B14:S14'); //red
setRedColor($sheet,"B14");

$sheet->mergeCells('B15:S15');

$sheet->setCellValue("B15",'Техническое задание');



$sheet->mergeCells('B16:S18');//техническое задание
textCenter($sheet, "B16");

$sheet->mergeCells('B19:S19');
setRedColor($sheet,"B19");
$sheet->mergeCells('B20:S20');
$sheet->setCellValue("B20",'Результаты поиска');

textCenter($sheet, "B20");

$provider1 = [

    "Местонахождение" => "Шеньчжень",
    "Выпускаемая продукция" => "Болты, гайки, шурупы",
    "Характеристики товара" => "Отвечают указанным требованиям",
    "Фото товара" => "Изображение",
    "Минимальный заказ" => "Одна коробка ( 500 шт )",
    "Срок производства" => "Уточнять по наличию",
    "Оценка благонадежности" => [
        "Уставный капитал" => "500 000 RMB",
        "Основной вид деятельности" => "Мелкие металлические изделия",
        "Дата основания" => "10/2/2015",
        "Срок действия лицензии" => "На неопределенный срок",
        "Наличие сайта" => " "],
    "Вес" => " ",
    "Прайс-лист" => " ",
    "Комментарии" => " Комментарий какой-то"


];


$providersCount = 1; //итератор поставщиков
//РЕЗУЛЬТАТЫ ПОИСКА
$startrow = makeProvider($sheet,21,$provider1);


function makeProvider($sheet,$startrow,$provider){
 $sheet->mergeCellsByColumnAndRow(1,$startrow,5,$startrow); //Номер поставщика
 $sheet->mergeCellsByColumnAndRow(5,$startrow,18,$startrow); //empty
 $sheet->setCellValueByColumnAndRow(1,$startrow,"Поставщик №".$GLOBALS['providersCount']); //set value
 $startrow++;
 $GLOBALS['providersCount']++;

 foreach ($provider as $key => $value) {


      if (is_array($value)) { // для оценки благонадежности

       $sheet->mergeCellsByColumnAndRow(1,$startrow,5,$startrow+count($value)-1);
       $sheet->mergeCellsByColumnAndRow(1,$startrow,5,$startrow+count($value)-1);
       $sheet->setCellValueByColumnAndRow(1, $startrow, $key);
       foreach ($value as $key => $val){
        $sheet->mergeCellsByColumnAndRow(6,$startrow,8,$startrow);
        $sheet->setCellValueByColumnAndRow(6, $startrow, $key);
        $sheet->mergeCellsByColumnAndRow(9, $startrow, 18, $startrow);
        $sheet->setCellValueByColumnAndRow(9, $startrow, $val);
        $startrow++;
       }
      $startrow++;
      } else {
       $sheet->mergeCellsByColumnAndRow(1, $startrow, 5, $startrow);
       $sheet->setCellValueByColumnAndRow(1, $startrow, $key);
       $sheet->mergeCellsByColumnAndRow(6, $startrow, 18, $startrow);

       $sheet->setCellValueByColumnAndRow(6, $startrow, $value);


       $startrow++;
      }


 }
 return $startrow;
}



$providerModel = Provider::model();
$providers = $providerModel->findAll();
var_dump($providers);


//$startrow=makeCalc($sheet,$calculation,$startrow);
function makeCalc($sheet,$calculation,$startrow){
    $sheet->mergeCellsByColumnAndRow(1, $startrow, 18, $startrow);
    $sheet->setCellValueByColumnAndRow(1,$startrow,"КАЛЬКУЛЯЦИЯ");

    function setCalcRow($sheet,$startrow,$arraydata)
    {
        $sheet-> mergeCellsByColumnAndRow (1,$startrow,5,$startrow);//первая колонка всегда длинаня
        $sheet-> setCellValueByColumnAndRow (1,$startrow,$arraydata[0]);
        $startcolumn = 6;
        for ($i = 1; $i < count($arraydata); $i++) {
            $sheet-> setCellValueByColumnAndRow ($startcolumn,$startrow,$arraydata[$i]);
            $startcolumn++;

        }
        return $startrow+1;
    }

}














    /*for($i=0; $i<count($calculation); $i++){ //для каждого типа калькуляции
        $iter = 0;
        $sheet->mergeCellsByColumnAndRow(1, $startrow, 5, $startrow);
        $sheet->mergeCellsByColumnAndRow(6, $startrow, 18, $startrow);

        $sheet->setCellValueByColumnAndRow(6, $startrow, "TEST");
        $startrow++;

        $isFirstColumn = true;
        $isFirstRow = true;
        //$sheet->setCellValueByColumnAndRow();
        for ($j=0; $j<count($calculation[$i]); $j++) {//для каждого товара в типе

            foreach ($calculation[$i][$j] as $key => $value) { //для каждого значения калькуляции
                /*if (!is_string($value)) {*/

                /*} else {
                    $sheet->mergeCellsByColumnAndRow(1, $startrow, 5, $startrow);
                    $sheet->mergeCellsByColumnAndRow(6, $startrow, 18, $startrow);
                    $sheet->setCellValueByColumnAndRow(1, $startrow, $key);
                    $sheet->setCellValueByColumnAndRow(6, $startrow, $value);
                    $startrow++;
                }*/
       /*     }
            $startrow++;

        }
        $startrow++;
    }
    return $startrow;*/



// Выравнивание текста



/*for ($i = 2; $i < 10; $i++) {
	for ($j = 2; $j < 10; $j++) {
        // Выводим таблицу умножения
        $sheet->setCellValueByColumnAndRow(
                                          $i - 2,
                                          $j,
                                          $i . "x" .$j . "=" . ($i*$j));
	    // Применяем выравнивание
	    $sheet->getStyleByColumnAndRow($i - 2, $j)->getAlignment()->
                setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	}
}*/
// Выводим HTTP-заголовки
 header ( "Expires: Mon, 1 Apr 1974 05:00:00 GMT" );
 header ( "Last-Modified: " . gmdate("D,d M YH:i:s") . " GMT" );
 header ( "Cache-Control: no-cache, must-revalidate" );
 header ( "Pragma: no-cache" );
 header ( "Content-type: application/vnd.ms-excel" );
 header ( "Content-Disposition: attachment; filename=matrix.xls" );

// Выводим содержимое файла
 $objWriter = new PHPExcel_Writer_Excel5($xls);
 $objWriter->save('php://output');



