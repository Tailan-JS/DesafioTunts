<?php
require __DIR__.'/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


$ch = curl_init( "https://restcountries.com/v2/all?fields=name,capital,currencies,area");
curl_setopt($ch,CURLOPT_RETURNTRANSFER,true);
curl_setopt($ch,CURLOPT_SSL_VERIFYPEER, false);
$result = json_decode(curl_exec($ch));


$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1','Countries List');
$cells = [['Name','Capital','Area','Currencies']];

//setting styles
$stylesTitle = [
    'font' => [
        'bold' => true,
        'size' => 16,
        'color' => [
            'rgb' => '4F4F4F'
        ]
    ],
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
    ],
];
$stylesHeader = [
    'font' => [
        'size' => 12,
        'color' => [
            'rgb' => '808080'
        ]
    ]
];
$sheet->getStyle('A1')->applyFromArray($stylesTitle);
$sheet->getStyle('A2:D2')->applyFromArray($stylesHeader);
$sheet->getStyle('C3:C252')->getNumberFormat()
->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
$sheet->mergeCells('A1:D1');
foreach($result as $country){
    //generating array with currency codes
    $codes = '';
    if(isset($country->currencies)){
        foreach($country->currencies as $index => $code)
           $codes .= $index==0?$code->code:','.$code->code;
    }else{
        $codes .= ' - ';
    }
    //verifying if there are null fields
    $capital=isset($country->capital)?$country->capital:' - ';
    $area= isset($country->area)?$country->area:' - ';
    //arranging data in an array
    array_push($cells,[$country->name,$capital,$area,$codes       
        ]);
    }

$sheet->fromArray($cells,'-','A2');
$writer = new Xlsx($spreadsheet);
$writer->save('countries.xlsx');


?>