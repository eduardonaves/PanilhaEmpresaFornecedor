<?php

error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');

if (PHP_SAPI == 'cli')
	die('This example should only be run from a Web Browser');

/** Include PHPExcel */
require_once '../../libs/PHPExcel_1.7.9/Classes/PHPExcel.php';

// Create new PHPExcel object
$objPHPExcel = new PHPExcel();

$empresa = array('empresa' => 'JS MARITUBA 2',
				 'end'	=> 'RUA RODOVIA BR 316 KM 10 4A MARITUBA',	
				 'bairro'	=> 'SAO JOAO',
				 'cidade' => 'MARITUBA',
				 'uf' => 'PA',
				 'cep' =>	'67200000',
				 'telefone' => '(91) 3073-7390',
				 'fax' => '',
				 'cnpj' =>	'04.185.877/0004-06',
				 'ie' => '152507922',
				 'mail'	=> 'js.belem2.nfe@jspecas.com.br');

$pedido = array('pedido' => '261517',
				'fornecedor' => 'KNORR MAI/2012',
				'emissao' => '13/03/2012');

$fornecedor = array( 'fornecedor' => 'JS MARITUBA 2',
				 	 'end'	=> 'RUA RODOVIA BR 316 KM 10 4A MARITUBA',	
					 'bairro'	=> 'SAO JOAO',
					 'cidade' => 'MARITUBA',
					 'uf' => 'PA',
					 'cep' =>	'67200000',
					 'telefone' => '(91) 3073-7390',
					 'fax' => '',
					 'cnpj' =>	'04.185.877/0004-06',
					 'ie' => '152507922',
					 'mail'	=> 'js.belem2.nfe@jspecas.com.br');

//alterado para index por causa da manupulação do for la em baixo: linha 359
$items = array(0 => '001662413R ',
			   1 => '1',
			   2 => 'REPARO VALVULA TRANSFERENCIA CAIXA	',
			   3 => 'KNORR /',
			   4 => 'SP70560 /',
			   5 => '5',
			   6 => '1',
			   7 => '5'
);

// Set document properties
$objPHPExcel->getProperties()->setCreator("JS BackOfice")
							 ->setLastModifiedBy("JS BackOfice")
							 ->setTitle("Office 2007 XLSX Test Document")
							 ->setSubject("Office 2007 XLSX Test Document")
							 ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
							 ->setKeywords("office 2007 openxml php")
							 ->setCategory("Test result file");

//set aligment de todo o sheet pra esquerda
$objPHPExcel->getDefaultStyle()
    ->getAlignment()
    ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);

//ajustando os larguras das celulas
$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(20.00);  
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(13.00);  
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(40.00);  
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(20.00);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(20.00);

//seta bordas
$objPHPExcel->getActiveSheet()->getStyle('A1:A19')->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle('F1:F18')->getFont()->setBold(true);
//$objPHPExcel->getActiveSheet()->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); 

//items para cada coluna e setado o estilo de fonte e borders
$objPHPExcel->getActiveSheet()->getStyle('A22')->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle('B22')->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle('C22')->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle('D22')->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle('E22')->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle('F22')->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle('G22')->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle('H22')->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle('D10')->getFont()->setBold(true);

//empresa
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B1:H1');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B2:H2');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B3:H3');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B4:E4');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('G4:H4');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('C5:H5');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B6:E6');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('G6:H6');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B7:E7');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('G7:H7');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B8:H8');

//pedido
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A9:H9');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('G10:H10');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A11:H11');

//fornecedor
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B12:H12');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B13:H13');	
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B14:H14');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B15:E15');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('G15:H15');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('C16:H16');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B17:E17');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('G17:H17');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B18:E18');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('G18:H18');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B19:H19');

//items
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A20:H20');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B21:H21');

//Empresa
$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A1', 'Empresa')
            ->setCellValue('B1', $empresa['empresa'])
            ->setCellValue('A2', 'End.')
            ->setCellValue('B2', $empresa['end'])
            ->setCellValue('A3', 'Bairro')
            ->setCellValue('B3', $empresa['bairro'])
            ->setCellValue('A4', 'Cidade')
            ->setCellValue('B4', $empresa['cidade'])
            ->setCellValue('F4', 'UF')
            ->setCellValue('G4', $empresa['uf'])
            ->setCellValue('A5', 'CEP')
            ->setCellValue('B5', $empresa['cep'])
            ->setCellValue('A6', 'Telefone')
            ->setCellValue('B6', $empresa['telefone'])
            ->setCellValue('F6', 'Fax')
            ->setCellValue('G6', $empresa['fax'])
            ->setCellValue('A7', 'CNPJ')
            ->setCellValue('B7', $empresa['cnpj'])
            ->setCellValue('F7', 'I.E.')
            ->setCellValue('G7', $empresa['ie'])
            ->setCellValue('A8', 'E-Mail')
            ->setCellValue('B8', $empresa['mail']);

//pedido de compra
$objPHPExcel->setActiveSheetIndex(0)
			->setCellValue('A10', 'Pedido de Compra')
            ->setCellValue('B10', $pedido['pedido'])
            ->setCellValue('D10', 'Num. Fornecedor')
            ->setCellValue('E10', $pedido['fornecedor'])
            ->setCellValue('F10', 'Emissao')
            ->setCellValue('G10', $pedido['emissao']);

//fornecedor
$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A12', 'Fornecedor')
            ->setCellValue('B12', $fornecedor['fornecedor'])
            ->setCellValue('A13', 'End.')
            ->setCellValue('B13', $fornecedor['end'])
            ->setCellValue('A14', 'Bairro')
            ->setCellValue('B14', $fornecedor['bairro'])
            ->setCellValue('A15', 'Cidade')
            ->setCellValue('B15', $fornecedor['cidade'])
            ->setCellValue('F15', 'UF')
            ->setCellValue('G15', $fornecedor['uf'])
            ->setCellValue('A16', 'CEP')
            ->setCellValue('B16', $fornecedor['cep'])
            ->setCellValue('A17', 'Telefone')
            ->setCellValue('B17', $fornecedor['telefone'])
            ->setCellValue('F17', 'Fax')
            ->setCellValue('G17', $fornecedor['fax'])
            ->setCellValue('A18', 'CNPJ')
            ->setCellValue('B18', $fornecedor['cnpj'])
            ->setCellValue('F18', 'I.E.')
            ->setCellValue('G18', $fornecedor['ie'])            
            ->setCellValue('A19', 'E-Mail')
            ->setCellValue('B19', $fornecedor['mail']);

//items
$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A22', 'Codigo de Item')
            ->setCellValue('B22', 'Categoria')
            ->setCellValue('C22', 'Descricao do Item')
            ->setCellValue('D22', 'Fabricante')
            ->setCellValue('E22', 'Cod Fabricante')
            ->setCellValue('F22', 'Qtde')
            ->setCellValue('G22', 'Vlr Unit.')
            ->setCellValue('H22', 'Total');

//loop para integrar os items na panilha
//altera as linhas
for($i = 23; $i < 30; $i++){

	//converte da tabela ascii para char
	$j = 41;
	$num  = base_convert($j, 16, 10);
	$number = chr($num) . $i;
	$inicial = $number;

	$c = 0;
	//altera as colunas
	while($j < 49){

		$num  = base_convert($j, 16, 10);
		$number = chr($num) . $i;
		$final = $number;

		$objPHPExcel->setActiveSheetIndex(0)
   			->setCellValue((string)$number, $items[$c].' - '.$i );

		$j++;
		$c++;
	}	
}

//seta todas as bordar de todo excel utilizado
$objPHPExcel->getActiveSheet()->getStyle('A1:'.$final)->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);

//inserindo row
//$objPHPExcel->getActiveSheet()->insertNewRowBefore($objPHPExcel->getActiveSheet()->getHighestRow()+1);
//$objPHPExcel->getActiveSheet()->insertNewRowBefore($objPHPExcel->getActiveSheet()->getHighestRow()+1);
//pegando maior row e coluna utilizada
//$num_rows = $objPHPExcel->getActiveSheet()->getHighestRow();
//die($num_rows);
//$num_cols = $objPHPExcel->getActiveSheet()->getHighestColumn();
//die($num_cols);

//Rename worksheet
$objPHPExcel->getActiveSheet()->setTitle('Simple');

//Set active sheet index to the first sheet, so Excel opens this as the first sheet
$objPHPExcel->setActiveSheetIndex(0);

//Redirect output to a client’s web browser (Excel5)
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="01simple.xls"');
header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
header('Cache-Control: max-age=1');

// If you're serving to IE over SSL, then the following may be needed
header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
header ('Pragma: public'); // HTTP/1.0

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');
exit;

?>