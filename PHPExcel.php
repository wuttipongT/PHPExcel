<?php
/*
offical site : https://phpexcel.codeplex.com/
source : https://github.com/PHPOffice/PHPExcel
 */
require_once '../classes/PHPExcel/Classes/PHPExcel.php';

$SRNO 	= isset($_REQUEST['SRNO']) ? $_REQUEST['SRNO'] : 'WHDD16-713';
$PKLIST = new PackingList($SRNO);
$objPHPExcel = new PHPExcel();

$PK_QSINO 		= $PKLIST->getQSINOS();
$PK_WETORDERS = $PKLIST->getWETORDERS();
$PK_PALLETS 	= $PKLIST->getPALLETS();
$PK_SMP 			= $PKLIST->getPALLETSummary();

  $num_row = 1;
  $npage = 0;
  $height = 30;

	$styleArray = array(
	 'font'  => array(
			 'color' => array('rgb' => '000000'),
			 'size'  => 16,
			 'name'  => 'Tahoma'
	 ));

	 $objPHPExcel->getProperties()->setCreator("[C000067]")
                                							 ->setLastModifiedBy("[C000067]")
                                							 ->setTitle("xxxxxx")
                                							 ->setSubject("xxxx")
                                							 ->setDescription("xxx")
                                							 ->setCategory("xxxx");
  $objPHPExcel->getDefaultStyle()->applyFromArray($styleArray);
  // Create a first sheet
  $objPHPExcel->setActiveSheetIndex(0);
  $sheet = $objPHPExcel->getActiveSheet();

	$sheet->getPageSetup()
	 ->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4)
	 ->setFitToPage(true)
	 ->setFitToWidth(1)
	 ->setFitToHeight(0);

	 $sheet->getPageMargins()
	 ->setTop(0.25)
	 ->setRight(0.25)
	 ->setLeft(0.25)
	 ->setBottom(0.25);
	
$sheet->getColumnDimension('A')->setWidth(17);
  $sheet->getColumnDimension('B')->setWidth(40);
  $sheet->getColumnDimension('C')->setWidth(20);
  $sheet->getColumnDimension('D')->setWidth(20);
	$sheet->getColumnDimension('E')->setWidth(20);
	$sheet->getColumnDimension('F')->setWidth(21);

  $QSINO = sizeof($PK_QSINO['QSISONO']);
  $WETNO = sizeof($PK_WETORDERS['WETORDER']);
  $WDMOR = sizeof($PK_WETORDERS['ORDERNO']);
  $CUSPO = sizeof($PK_WETORDERS['CUSPO']);

  $a = $QSINO > $WETNO ? $QSINO : $WETNO;
  $b = $WDMOR > $CUSPO ? $WDMOR : $CUSPO;
  $c = $a > $b ? $a : $b;
  $mod = 22;
	foreach($PK_PALLETS as $index => $PALLET){
		$SRNO = $PALLET['SRNO'];
		$MODEL = $PALLET['MODEL'];
		$WETORDER = $PALLET['WETORDER'];
		$WETMODEL = $PALLET['WETMODEL'];
		$WDPO = $PALLET['WDPO'];

		$PNQ_PLT = intval($PALLET['PLLQTY']);
		$PNQ_MTY = intval($PALLET['MCQTY']);
		$PNQ_QTY = intval($PALLET['QTY']);
		$PNQ_PTY = intval($PALLET['PPLQTY']);
		$PNQ_NWG = round(floatval($PALLET['NETWEIGHT']),2);
		$PNQ_GWG = round(floatval($PALLET['GROSSWEIGHT']),2);
		$PNQ_MST = round(floatval($PALLET['MESSUREMENT']),2);

		if ($PNQ_NWG == 0){
			$PNQ_NWG = $PNQ_GWG;
			if ($PNQ_PTY>=65){
				$PNQ_NWG -= 2;
			}
			$PNQ_NWG -= 18;
		}

		$ALL_PLT += $PNQ_PLT;
		$ALL_QTY += $PNQ_QTY;
		$ALL_NWG += $PNQ_NWG;
		$ALL_GWG += $PNQ_GWG;
		$ALL_MST += $PNQ_MST;

		$PNQ_PLT = number_format($PNQ_PLT,0);
		$PNQ_QTY = number_format($PNQ_QTY,0);
		$PNQ_NWG = number_format($PNQ_NWG,2);
		$PNQ_GWG = number_format($PNQ_GWG,2);
		$PNQ_MST = number_format($PNQ_MST,2);

		$MESNT = $PNQ_MST. PHP_EOL .$PALLET['DIMENSION'];
		$CAPA  = $PALLET['CAPACITY'] == '0TB' ? '(NO HARD DISK DRIVE)' : '';
		$MSGTXT = "EXTERNAL HARD DISK DRIVE $CAPA
					WET Model : $WETMODEL
					MODEL : $MODEL
					WD PO.NO. : $WDPO
					$PNQ_PTY CARTONS PER PALLET
					SUB TOTAL
					[ $PNQ_PLT PALLET ]";

		if ($index == 0){
			$npage++;
			setHeader(null, $npage);
			$start = $num_row + 1;
		}

		if( $num_row % $mod + $c == ($mod + $c) - 1 ){
      $npage++;
      $sheet->setBreak( cell(array('A', $num_row) ), PHPExcel_Worksheet::BREAK_ROW );
      $num_row = $num_row + 1;
      setHeader(null, $npage);
			$start = $num_row + 1;
      $c = 0;
    }

		setBody($MSGTXT,$PNQ_QTY,$PNQ_NWG,$PNQ_GWG,$MESNT);

	}
  $sheet->getStyle( cell(array(array('A', $start),array('F', $num_row))) )
        ->getAlignment()
        ->setWrapText(true);

  $sheet->getStyle( cell(array(array('A', $start),array('F', $num_row))) )->applyFromArray(array(
    'alignment' => array(
      //  'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
        'vertical'=> PHPExcel_Style_Alignment::VERTICAL_TOP
    )
  ));
  $sheet->getStyle( cell(array(array('C', $start),array('F', $num_row))) )->applyFromArray(array(
    'alignment' => array(
      'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_RIGHT,
    )
  ));
	// Packing List Summary Part 1
	$ALL_PLT = number_format($ALL_PLT,0);
	$ALL_QTY = number_format($ALL_QTY,0);
	$ALL_NWG = number_format($ALL_NWG,2);
	$ALL_GWG = number_format($ALL_GWG,2);
	$ALL_MST = number_format($ALL_MST,2);
  setFooter($ALL_PLT,$ALL_QTY,$ALL_NWG,$ALL_GWG,$ALL_MST);

	// foreach($sheet->getRowDimensions() as $rd) {
  //   $rd->setRowHeight($height);
	// }
	// Redirect output to a client’s web browser (Excel5)
	header('Content-Type: application/vnd.ms-excel');
	header('Content-Disposition: attachment;filename="SFCML601C.xls"');
	header('Cache-Control: max-age=0');
	// If you're serving to IE 9, then the following may be needed
	header('Cache-Control: max-age=1');
	// If you're serving to IE over SSL, then the following may be needed
	header ('Expires: '. date("D, d M Y H:i:s")); // Date in the past
	header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
	header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
	header ('Pragma: public'); // HTTP/1.0
	$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
	$objWriter->save('php://output');

	function setBody($MSGTXT,$PNQ_QTY,$PNQ_NWG,$PNQ_GWG,$MESNT){
		try {
			$objPHPExcel = $GLOBALS['objPHPExcel'];
			$sheet = $objPHPExcel->getActiveSheet();
			$num_row     = $GLOBALS['num_row'];

			$num_row +=1;
			$sheet->setCellValue( cell(array('A', $num_row)), "NO MARK");
			$sheet->setCellValue( cell(array('B', $num_row)), $MSGTXT);
			$sheet->setCellValue( cell(array('C', $num_row)), $PNQ_QTY);
			$sheet->setCellValue( cell(array('D', $num_row)), $PNQ_NWG);
			$sheet->setCellValue( cell(array('E', $num_row)), $PNQ_GWG);
			$sheet->setCellValue( cell(array('F', $num_row)), $MESNT);

			$GLOBALS['num_row'] = $num_row;
		} catch (Exception $e) {
				throw $e;
		}
	}
	function setFooter($ALL_PLT,$ALL_QTY,$ALL_NWG,$ALL_GWG,$ALL_MST){
    $objPHPExcel = $GLOBALS['objPHPExcel'];
    $num_row     = $GLOBALS['num_row'];
    $height      = $GLOBALS['height'];
    $sheet 			 = $objPHPExcel->getActiveSheet();
		$PK_SMP 		 = $GLOBALS['PK_SMP'];
		$style = array(
			'alignment' => array(
					'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
			),
		);
		$style2 = array(
			'alignment' => array(
				//  'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
					'vertical'=> PHPExcel_Style_Alignment::VERTICAL_CENTER
			),
		);

    $styleArray = array(
     'font'  => array(
         'color' => array('rgb' => 'FF0000'),
    ));
		$style3 = array(
			'alignment' => array_merge($style['alignment'], $style2['alignment'])
		);
			//allborders
		$style4 = array(
			'borders' => array(
				'top' => array('style' => PHPExcel_Style_Border::BORDER_THIN),
				'bottom' => array('style' => PHPExcel_Style_Border::BORDER_THIN)
			)
		);
    try {
			$num_row += 2;
			$sheet
						->setCellValue( cell(array('A', $num_row)), "TOTAL $ALL_PLT PALLETS")
						->mergeCells( cell(array(array('A', $num_row),array('B', $num_row))) )
						->setCellValue( cell(array('C', $num_row)), $ALL_QTY)
						->setCellValue( cell(array('D', $num_row)), $ALL_NWG)
						->setCellValue( cell(array('E', $num_row)), $ALL_GWG)
						->setCellValue( cell(array('F', $num_row)), $ALL_MST)
			 			->getStyle( cell(array(array('A', $num_row),array('F', $num_row ))) )->applyFromArray($style3 + $style4);
			$sheet->getRowDimension($num_row)->setRowHeight(30);
      $sheet->setBreak( cell(array('A', $num_row) ), PHPExcel_Worksheet::BREAK_ROW );
      //$num_row += 2;
			$index = 1;
			$col = array("A","C","E");
			$i =0;
			$num_row += 1;
			foreach ($PK_SMP as $value) {
				$QTY = $value['QTY'];
				$AMOUNT = $value['AMOUNT'];
				$SHIPTO = $value['SHIPTO'];
				$SUM = number_format($value['SUM'],0);
				$PALIST = implode(",", $value['PALLETNO']) ;
				$QSINOLIST = implode("<br>", $value['QSINO']) ;
				if($index == 1) $start = $num_row;

				$sheet->setCellValue( cell(array($col[$i], $num_row+1)), "Say Total : 1 Pallet Including $QTY Cartons only");
	      // $num_row += 1;
	      $sheet->setCellValue( cell(array($col[$i], $num_row+2)), "Shipping Mark");
	      // $num_row += 1;
	      $sheet->setCellValue( cell(array($col[$i], $num_row+3)), "WD");
	      // $num_row += 1;
	      $sheet->setCellValue( cell(array($col[$i], $num_row+4)), $SHIPTO);
	      // $num_row += 1;
	      $sheet->setCellValue( cell(array($col[$i], $num_row+5)), "P/L : $PALIST");
	      // $num_row += 1;
	      $sheet->setCellValue( cell(array($col[$i], $num_row+6)), "( $SUM CARTONS)");
	      // $num_row += 1;
	      $sheet->setCellValue( cell(array($col[$i], $num_row+7)), "MADE IN THAILAND");
	      // $num_row += 1;
	      $sheet->setCellValue( cell(array($col[$i], $num_row+8)), "QSI SN NO.");
				// $num_row += 1;
	      $sheet->setCellValue( cell(array($col[$i], $num_row+9)), $QSINOLIST);

				$index ++;
				$i++;
				if($i > 2){
					$i=0;
					$num_row +=9;
				}
			}
			$num_row +=9;
      $sheet->getStyle( cell(array(array('A', $start + 1),array('F', $num_row ))) )->applyFromArray($styleArray);

      $GLOBALS['num_row'] = $num_row;
    } catch (Exception $e) {
      throw $e;
    }
  }
	function setHeader($value, $npage){
		$objPHPExcel = $GLOBALS['objPHPExcel'];
		$num_row     = $GLOBALS['num_row'];
		$height      = $GLOBALS['height'];
		$INVNC 			 = $GLOBALS['SRNO'];

		$style = array(
			'alignment' => array(
					'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
			),
		);
		$style2 = array(
			'alignment' => array(
				//  'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
					'vertical'=> PHPExcel_Style_Alignment::VERTICAL_CENTER
			),
		);
		$style3 = array(
			'alignment' => array_merge($style['alignment'], $style2['alignment'])
		);
		$style4 = array(
			'borders' => array(
				'top' => array('style' => PHPExcel_Style_Border::BORDER_THIN),
				'bottom' => array('style' => PHPExcel_Style_Border::BORDER_THIN)
			)
		);
		$style5 = array(
		 'font'  => array(
				 'size'  => 10,
		 ));
		$sheet = $objPHPExcel->getActiveSheet();
		try {
			$sheet->setCellValue( cell(array('A', $num_row)), "WORLD ELECTRIC (THAILAND)  LTD.")
						->mergeCells( cell(array(array('A', $num_row),array('B', $num_row))) )
            ->getStyle( cell(array('A', $num_row)) )->applyFromArray(array('font'=>array('bold'=>true)));
			$sheet->setCellValue( cell(array('F', $num_row)), "Page $npage")
						->getStyle( cell(array('F', $num_row)) )->applyFromArray(array(
							'alignment' => array(
									'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_RIGHT,
							),
              'font'  => array(
           			 'bold'  => true,
           	  )
						));
			$num_row += 1;
			$sheet->setCellValue( cell(array('A', $num_row)), "HEAD OFFICE : 236 MOO 2 NONGCHARK BANBUNG CHONBURI 20170,THAILAND.")
						->mergeCells( cell(array(array('A', $num_row),array('B', $num_row))) );
			$num_row += 1;
			$sheet->setCellValue( cell(array('A', $num_row)), "TEL. (038) 4444000 (AUTOMATIC LINES),FAX. (038) 443915-7")
						->mergeCells( cell(array(array('A', $num_row),array('B', $num_row))) );
			$num_row += 1;
			$sheet->setCellValue( cell(array('A', $num_row)), "BKK OFFICE    : 2ND.FLOOR C.C.T. BLDG., 109 SURAWONG ROAD,BANGKOK 10500,THAILAND")
						->mergeCells( cell(array(array('A', $num_row),array('B', $num_row))) );
      $num_row += 1;
			$sheet->setCellValue( cell(array('A', $num_row)), "TEL. (02) 2372700-2,2377000-1,2339353 FAX: (02) 2362905,2354705,2382305")
						->mergeCells( cell(array(array('A', $num_row),array('B', $num_row))) );
			$sheet->getStyle( cell(array(array('A', $num_row - 4),array('F', $num_row ))) )->applyFromArray($style5);
      $sheet->getStyle( cell(array(array('A', $num_row - 3),array('F', $num_row ))) )->applyFromArray(array(  'font'  => array( 'bold'  => false, )));
			$num_row += 1;
	    $sheet->setCellValue( cell(array('A', $num_row)), $INVNC)
						->mergeCells( cell(array(array('A', $num_row),array('B', $num_row))) )
						->getStyle( cell(array(array('A', $num_row),array('A', $num_row ))) )->applyFromArray(array('font'=>array('name'=>'Code 128','size'=>72,'bold'=>false) + $style));

			$num_row += 1;
			$sheet->setCellValue( cell(array('A', $num_row)), "                       ".$INVNC)
						->mergeCells( cell(array(array('A', $num_row),array('B', $num_row))) )
						->getStyle( cell(array(array('A', $num_row),array('A', $num_row ))) )->applyFromArray($style5);

						$sheet->getStyle( cell(array('A', $num_row)) )->getAlignment()->setWrapText(true);
	    $sheet
		        ->setCellValue( cell(array('E', $num_row)), "Export Date :")
						->setCellValue( cell(array('F', $num_row)), date('Y-m-d'));
			$num_row += 1;
		  $sheet->setCellValue( cell(array('E', $num_row)), "Invoice No. :")
						->setCellValue( cell(array('F', $num_row)), $INVNC);
			$num_row += 1;
			$sheet->setCellValue( cell(array('E', $num_row)), "BOI NO. :")
					->setCellValue( cell(array('F', $num_row)), "1393");
			$sheet->getStyle( cell(array(array('E', $num_row - 3),array('E', $num_row ))) )->applyFromArray(array(
				'alignment' => array(
						'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_RIGHT,
				),
        'font'=>array('bold'=>true)
			));
			$sheet->getStyle( cell(array(array('F', $num_row - 3),array('F', $num_row ))) )->applyFromArray(array(
				'alignment' => array(
		        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT,
		    ),
			));
      if($npage == 1){

				$QSINO = implode(PHP_EOL, $GLOBALS['PK_QSINO']['QSISONO']);
				$WETNO = implode(PHP_EOL, $GLOBALS['PK_WETORDERS']['WETORDER']);
				$WDMOR = implode(PHP_EOL, $GLOBALS['PK_WETORDERS']['ORDERNO']);
				$CUSPO = implode(PHP_EOL, $GLOBALS['PK_WETORDERS']['CUSPO']);
        $num_row += 2;
        $sheet
              ->setCellValue( cell(array('C', $num_row)), "QSI SN NO.")
              ->setCellValue( cell(array('D', $num_row)), "WET Order No.")
              ->setCellValue( cell(array('E', $num_row)), "MO NO.")
              ->setCellValue( cell(array('F', $num_row)), "Customer Order.");
        $sheet->getStyle( cell(array(array('C', $num_row),array('F', $num_row ))) )->applyFromArray(array(
					'font'  => array(
							'bold'  => true,
					)
				));
				$num_row += 1;
				$sheet
							->setCellValue( cell(array('C', $num_row)), $QSINO)
							->setCellValue( cell(array('D', $num_row)), $WETNO)
							->setCellValue( cell(array('E', $num_row)), $WDMOR)
							->setCellValue( cell(array('F', $num_row)), $CUSPO);
				$sheet->getStyle( cell(array(array('C', $num_row),array('F', $num_row))) )
							->getAlignment()
							->setWrapText(true);
      }
			$num_row += 2;
		  $sheet
	          ->setCellValue( cell(array('A', $num_row)), "Packing List")
						->mergeCells( cell(array(array('A', $num_row),array('F', $num_row))) )
						->getStyle( cell(array(array('A', $num_row),array('F', $num_row ))) )->applyFromArray($style3 + array('font'=>array('bold'=>true)));
			$sheet->getRowDimension($num_row)->setRowHeight(30);
			$num_row += 1;
			$sheet
	          ->setCellValue( cell(array('A', $num_row)), "CASE MAKING")
	          ->setCellValue( cell(array('B', $num_row)), "DESCRIPTION")
	          ->setCellValue( cell(array('C', $num_row)), "QTY")
	          ->setCellValue( cell(array('D', $num_row)), "NET WEIGHT")
	          ->setCellValue( cell(array('E', $num_row)), "GROSS WEIGHT")
	          ->setCellValue( cell(array('F', $num_row)), "MEASUREMENT");
      $num_row += 1;
      $sheet
            ->setCellValue( cell(array('C', $num_row)), "(PCS)")
            ->setCellValue( cell(array('D', $num_row)), "(KGS)")
            ->setCellValue( cell(array('E', $num_row)), "(KGS)")
            ->setCellValue( cell(array('F', $num_row)), "(M3)")
						->getStyle( cell(array(array('A', $num_row - 1),array('F', $num_row ))) )->applyFromArray($style3 + $style4 + array('font'=>array('bold'=>true)));
      $sheet
            ->mergeCells( cell(array(array('A', $num_row-1),array('A', $num_row))) )
            ->mergeCells( cell(array(array('B', $num_row-1),array('B', $num_row))) );

			// 			$objPHPExcel->getDefaultStyle()
			//     ->getBorders()
			//     ->getTop()
			//         ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
			// $objPHPExcel->getDefaultStyle()
			//     ->getBorders()
			//     ->getBottom()
			//         ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
//			$sheet->getRowDimension($num_row)->setRowHeight(30);
		  $GLOBALS['num_row'] = $num_row;
		} catch (Exception $e) {
			throw $e;
		}
	}
	function cell($arr , $glue = ''){
		try {
			if(is_array($arr[0])){
		    $cells = array();
		    foreach ($arr as $key => $value) {
		      # code...
		      array_push($cells, implode('', $value));
		    }
		    return implode(':', $cells);
		  }
		  return implode('', $arr);
		} catch (Exception $e) {
			throw $e;
		}
	}
