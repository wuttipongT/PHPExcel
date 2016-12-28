<?php
/* **
https://github.com/pear/Spreadsheet_Excel_Writer
https://pear.php.net/manual/en/package.fileformats.spreadsheet-excel-writer.php
*/
include("../include.me.php");
require_once '../classes/StringBuilder.php';
include_once "Spreadsheet/Excel/Writer.php";

$SHELF_ID = isset($_REQUEST['SHELF_ID']) ? $_REQUEST['SHELF_ID'] : null;
$PARTCODE = isset($_REQUEST['PARTCODE']) ? $_REQUEST['PARTCODE'] : null;
$factory = isset($_REQUEST['factory']) ? $_REQUEST['factory'] : null;
$date = isset($_REQUEST['date']) ? $_REQUEST['date'] : null;

// if($date != ''){
//   $date_ex = explode('/', $date);
//   $date = implode('', array_reverse($date_ex));
// }

//$PAGE = $_REQUEST['PAGE'];
$workbook =& new Spreadsheet_Excel_Writer();
// $fHead =& $workbook->addFormat(array('Size' => 9,'Align' => 'center','vAlign'=>'vcenter','Bold' => 1,'FontFamily' =>'Tahoma','Border' => 0));
// $fHead2 =& $workbook->addFormat(array('Size' => 9,'Align' => 'left','Bold' => 1,'FontFamily' =>'Tahoma','Border' => 0, 'width'=>150));
// $fHead3 =& $workbook->addFormat(array('Size' => 9,'Align' => 'center','Bold' => 1,'FontFamily' =>'Tahoma','Border' => 0));
// $fDetial1 =& $workbook->addFormat(array('Size' => 9,'Align' => 'left','Bold' => 0,'FontFamily' =>'Tahoma','Border' => 0));
// $fDetial2 =& $workbook->addFormat(array('Size' => 9,'Align' => 'right','Bold' => 0,'FontFamily' =>'Tahoma','Border' => 0));
// $fDetial3 =& $workbook->addFormat(array('Size' => 9,'Align' => 'center','Bold' => 0,'FontFamily' =>'Tahoma','Border' => 0));
// $fDetial4 =& $workbook->addFormat(array('Size' => 9,'Align' => 'left','Bold' => 0,'FontFamily' =>'Tahoma'));
// $fDetial5 =& $workbook->addFormat(array('Size' => 9,'Align' => 'right','Bold' => 0,'FontFamily' =>'Tahoma','Border' => 0,'NumFormat' =>'#,##0'));
// $fDetial6 =& $workbook->addFormat(array('Size' => 9,'Align' => 'right','Bold' => 0,'FontFamily' =>'Tahoma','Border' => 0,'NumFormat' =>'#,##0','BorderTop' =>1,'Top' =>1));

$format_header =& $workbook->addFormat();
$format_header->setAlign('left');
$format_header->setVAlign('vcenter');
$format_header->setBold(1);
$format_header->setFontFamily('Tahoma');
//$format_header->setBorder(1);
//$format_header->setMerge();
$format_head_center =& $workbook->addFormat();
$format_head_center->setVAlign('vcenter');
$format_head_center->setBold(1);
$format_head_center->setFontFamily('Tahoma');
$format_head_center->setBorder(1);
$format_head_center->setMerge();

$worksheet =& $workbook->addWorksheet('SFCML7015');
$worksheet->setColumn(0,0,22);
$worksheet->setColumn(1,1,105);

$worksheet->setColumn(2,2,18);
$worksheet->setColumn(3,3,18);

//$worksheet->setColumn(3,7,15);
//setSheet($worksheet);

$format_header_Inven =& $workbook->addFormat();
$format_header_Inven->setAlign('center');
$format_header_Inven->setVAlign('vcenter');
$format_header_Inven->setBold(1);
$format_header_Inven->setFontFamily('Tahoma');

$worksheet->hideGridlines();
//$worksheet->setLandscape ();  //setPortrait();
$worksheet->setPrintScale (60); //$scale=40
$worksheet->setMarginLeft (0.25);
$worksheet->setMarginRight(0.15);
$worksheet->setMarginBottom(0.25);
$worksheet->setMarginTop(0.25);
$worksheet->setPaper(9); //9 = A4 , 11 =A5\

$format_body_right =& $workbook->addFormat();
$format_body_right->setAlign('right');
$format_body_right->setVAlign('vcenter');
$format_body_right->setFontFamily('Tahoma');
//$format_body_right->setBorder(1);
//$format_body_right->setLeft(1);
//$format_body_right->setRight(1);
//$format_body_right->setBottom(1);

$format_body =& $workbook->addFormat();
$format_body->setVAlign('vcenter');
$format_body->setFontFamily('Tahoma');
//$format_body->setBorder(1);Left
//$format_body->setLeft(1);
//$format_body->setRight(1);
//$format_body->setBottom(1);

$format_body_Approve =& $workbook->addFormat();
$format_body_Approve->setAlign('right');
$format_body_Approve->setVAlign('vcenter');
$format_body_Approve->setFontFamily('Tahoma');


$response = GET_SFCMONTHINVENTORY($SHELF_ID, $PARTCODE,$factory, $date);
$num_row = 0 ;
$couter = 1;
$npage = 0;
$aaDatas = $response['aaData'];
$height = 20;

foreach ( $aaDatas as $key => $value ) {
  # code...
$couter ++;
  if($num_row==0){
    $npage++;
    $num_row = 0;
    $worksheet->setRow($num_row, $height);
    $worksheet->write($num_row, 0, 'PROGRAM : SFCML601C', $format_header);
    $worksheet->write($num_row, 1, 'Inventory Report', $format_header_Inven);
    $worksheet->setMerge($num_row, 1, $num_row, 2);
    $worksheet->write($num_row, 3, sprintf('DATE : %s', date("d-M-y")), $format_header);

    $num_row = 1;
    $worksheet->setRow($num_row, $height);
    $worksheet->write($num_row, 0, sprintf('Shelf No : %s', $value['PALLETID']), $format_header);
    $worksheet->write($num_row, 1, sprintf('Year-Month : %s', $value['YYYYMM']), $format_header_Inven);
    $worksheet->setMerge($num_row, 1, $num_row, 2);
    $worksheet->write($num_row, 3, sprintf('PAGE : %s',$npage), $format_header);

    $num_row = 3;
    $worksheet->setRow($num_row, $height);
    $worksheet->write($num_row,0,'Part code',$format_head_center);
    $worksheet->write($num_row + 1,0,'',$format_head_center);
    $worksheet->write($num_row + 1,0,'',$format_head_center);

    $worksheet->setMerge($num_row, 0, $num_row + 1, 0);
    $worksheet->write($num_row,1,'Part name',$format_head_center);
    $worksheet->write($num_row+1,1,'',$format_head_center);

    $worksheet->setMerge($num_row, 1, $num_row + 1, 1);
    $worksheet->write($num_row,2,'Actual',$format_head_center);
    $worksheet->write($num_row,3,'',$format_head_center);
    $worksheet->setMerge($num_row, 2, $num_row, 3);
    $worksheet->write($num_row + 1,2,'Mid Quantity',$format_head_center);
    //$worksheet->setMerge($num_row + 1, 2, $num_row + 1, 3);
    $worksheet->write($num_row + 1,3,'Quantity',$format_head_center);
    //$worksheet->setMerge($num_row + 1, 4, $num_row + 1, 5);
    $worksheet->setRow($num_row + 1, $height);

    $num_row = 5;
  }
  //
  if($key > 0){
      $aaData = $aaDatas[$key - 1];
  }

  //$value['QST60_PALLET']
  $PALLET = isset($aaData) ? $aaData['PALLETID'] : $value['PALLETID'];
//  echo $QST60_PALLET .'!='. $value['QST60_PALLET'] . ' <br/>';
  if ($PALLET != $value['PALLETID']){
    $npage++;

    $num_row = $num_row + 3;
    $worksheet->setRow($num_row, $height);
    $worksheet->write($num_row,1,"PIC                                  Checker                                    Approve       " ,$format_body_Approve);
    $worksheet->write($num_row,2,"" ,$format_body_Approve);
  //  $worksheet->setMerge($num_row, 1, $num_row, 2);
    $worksheet->mergeCells($num_row, 1, $num_row, 2);
    $num_row = $num_row + 2;
    $worksheet->setRow($num_row, $height);
    $worksheet->write($num_row,1,"_______________________  _______________________  _______________________" ,$format_body_Approve);
    //$worksheet->write($num_row,2,"" ,$format_body_Approve);
    //$worksheet->setMerge($num_row, 1, $num_row, 2);
    $worksheet->mergeCells($num_row, 1, $num_row, 2);
  //  $num_row = $num_row + 1;
    // $body = 82;
    // $header = 5;
    // $footer = 6;
    // $ln     = ($body - $header) - $footer;
    // $i = 0;
    // for(;$i < $ln - $couter ;$i++)
    // {
    //
    //   // $worksheet->write($num_row,0, "", $format_body);
    //   // $worksheet->write($num_row,1, "", $format_body);
    //   // $worksheet->write($num_row,2, "",$format_body_right);
    //   // $worksheet->write($num_row,3, "", $format_body_right);
    //   $num_row = $num_row + 1;
    // }

    $num_row = $num_row + 1; //1
    $worksheet->setRow($num_row, $height);
    $worksheet->setHPagebreaks (array($num_row));
    $worksheet->write($num_row, 0, 'PROGRAM : SFCML601C', $format_header);
    $worksheet->write($num_row, 1, 'Inventory Report', $format_header_Inven);
    //$worksheet->setMerge($num_row, 1, $num_row, 2);
    $worksheet->write($num_row, 3, sprintf('DATE : %s', date("d-M-y")), $format_header);

    $num_row = $num_row + 1; //2
    $worksheet->setRow($num_row, $height);
    $worksheet->write($num_row, 0, sprintf('Shelf No : %s',$value['PALLETID']), $format_header);
    $worksheet->write($num_row, 1, sprintf('Year-Month : %s', $value['YYYYMM']), $format_header_Inven);
  //  $worksheet->setMerge($num_row, 1, $num_row, 2);
    $worksheet->write($num_row, 3, sprintf('PAGE : %s',$npage), $format_header);

    $num_row = $num_row + 2;
    $worksheet->setRow($num_row, $height);
    $worksheet->write($num_row,0,'Part code',$format_head_center);
    $worksheet->write($num_row + 1,0,'',$format_head_center);
    //$worksheet->setMerge($num_row, 0, $num_row + 1, 0);
    $worksheet->setRow($num_row + 1, $height);
    $worksheet->mergeCells($num_row, 0, $num_row + 1, 0);

    $worksheet->write($num_row,1,'Part name',$format_head_center);
    $worksheet->write($num_row + 1,1,'',$format_head_center);

    //$worksheet->setMerge($num_row, 1, $num_row + 1, 1);
    $worksheet->mergeCells($num_row, 1, $num_row + 1, 1);

    $worksheet->write($num_row,2,'Actual',$format_head_center);
    $worksheet->write($num_row,3,'',$format_head_center);
    //$worksheet->setMerge($num_row, 2, $num_row, 3);
    $worksheet->write($num_row + 1,2,'Mid Quantity',$format_head_center);
    //$worksheet->setMerge($num_row + 1, 2, $num_row + 1, 3);
    $worksheet->write($num_row + 1,3,'Quantity',$format_head_center);
    //$worksheet->setMerge($num_row + 1, 4, $num_row + 1, 5);

    $num_row =  $num_row + 2;
    $couter = 1;
  }

  else if ($num_row % 83 == 0)
  {
    $npage++;

    $num_row = $num_row + 3;
    $worksheet->setRow($num_row, $height);
    $worksheet->write($num_row,1,"PIC                                  Checker                                    Approve       " ,$format_body_Approve);
    $worksheet->write($num_row,2,"" ,$format_body_Approve);
  //  $worksheet->setMerge($num_row, 1, $num_row, 2);
    $worksheet->mergeCells($num_row, 1, $num_row, 2);
    $num_row = $num_row + 2;
    $worksheet->setRow($num_row, $height);
    $worksheet->write($num_row,1,"_______________________  _______________________  _______________________" ,$format_body_Approve);
    //$worksheet->write($num_row,2,"" ,$format_body_Approve);
    //$worksheet->setMerge($num_row, 1, $num_row, 2);
    $worksheet->mergeCells($num_row, 1, $num_row, 2);
  //  $num_row = $num_row + 1;
    // $body   = 82;
    // $header = 5;
    // $footer = 6;
    // $ln     = (($body - $header) - $footer) - $couter;
    // $i = 0;
    // for(;$i < $ln - $couter ;$i++)
    // {
    //
    //   // $worksheet->write($num_row,0, "", $format_body);
    //   // $worksheet->write($num_row,1, "", $format_body);
    //   // $worksheet->write($num_row,2, "",$format_body_right);
    //   // $worksheet->write($num_row,3, "", $format_body_right);
    //   $num_row = $num_row + 1;
    // }

    $num_row = $num_row + 1; //1
    $worksheet->setRow($num_row, $height);
    $worksheet->setHPagebreaks (array(12));
    // $worksheet->setHPagebreaks (array($num_row));
    $worksheet->write($num_row, 0, 'PROGRAM : XXXX', $format_header);
    $worksheet->write($num_row, 1, 'Inventory Report', $format_header_Inven);
    //$worksheet->setMerge($num_row, 1, $num_row, 2);
    $worksheet->write($num_row, 3, sprintf('DATE : %s', date("d-M-y")), $format_header);

    $num_row = $num_row + 1; //2
    $worksheet->setRow($num_row, $height);
    $worksheet->write($num_row, 0, sprintf('Shelf No : %s',$value['PALLETID']), $format_header);
    $worksheet->write($num_row, 1, sprintf('Year-Month : %s', $value['YYYYMM']), $format_header_Inven);
  //  $worksheet->setMerge($num_row, 1, $num_row, 2);
    $worksheet->write($num_row, 3, sprintf('PAGE : %s',$npage), $format_header);

    $num_row = $num_row + 2;
    $worksheet->setRow($num_row, $height);
    $worksheet->write($num_row,0,'Part code',$format_head_center);
    $worksheet->write($num_row + 1,0,'',$format_head_center);
    //$worksheet->setMerge($num_row, 0, $num_row + 1, 0);
    $worksheet->setRow($num_row + 1, $height);
    $worksheet->mergeCells($num_row, 0, $num_row + 1, 0);

    $worksheet->write($num_row,1,'Part name',$format_head_center);
    $worksheet->write($num_row + 1,1,'',$format_head_center);

    //$worksheet->setMerge($num_row, 1, $num_row + 1, 1);
    $worksheet->mergeCells($num_row, 1, $num_row + 1, 1);

    $worksheet->write($num_row,2,'Actual',$format_head_center);
    $worksheet->write($num_row,3,'',$format_head_center);
    //$worksheet->setMerge($num_row, 2, $num_row, 3);
    $worksheet->write($num_row + 1,2,'Mid Quantity',$format_head_center);
    //$worksheet->setMerge($num_row + 1, 2, $num_row + 1, 3);
    $worksheet->write($num_row + 1,3,'Quantity',$format_head_center);
    //$worksheet->setMerge($num_row + 1, 4, $num_row + 1, 5);

    $num_row =  $num_row + 2;
    $couter = 1;
  }
  $worksheet->write($num_row,0, $value['PARTCD'], $format_body);
  $worksheet->write($num_row,1, $value['QST22_PARTNM'], $format_body);
  $worksheet->write($num_row,2, $value['MID_QUANTITY'],$format_body_right);
  $worksheet->write($num_row,3, $value['INVENQTY'], $format_body_right);

  $worksheet->setRow($num_row, $height);
  $num_row += 1;
//  $couter ++;
}

$num_row = $num_row + 3;
$worksheet->write($num_row,1,"PIC                               Checker                                    Approve       " ,$format_body_Approve);
$worksheet->write($num_row,2,"" ,$format_body_Approve);
$worksheet->write($num_row,3,"" ,$format_body_Approve);
$worksheet->mergeCells($num_row, 1, $num_row, 2);
$num_row = $num_row + 2;
$worksheet->write($num_row,1,"_______________________  _______________________  _______________________" ,$format_body_Approve);
$worksheet->setMerge($num_row, 1, $num_row, 2);

$workbook->send("SFCML601C.xls");
$workbook->close();

function GET_SFCMONTHINVENTORY($SHELF_ID, $PARTCODE, $PLACE, $date){
  $oracle_user = $GLOBALS['oracle_user'];
  $oracle_pwd  = $GLOBALS['oracle_pwd'];
  $db = $GLOBALS['db'];

  $sb = new StringBuilder();
	$conn = oci_connect($oracle_user, $oracle_pwd, $db) or die('Connect Error');
	$response['e']	= false;
	$response['msg'] = array();

	$sb->append(" SELECT COUNT(PID) AS MID_QUANTITY,PARTCD,PALLETID,QST22_PARTNM,YYYYMM, SUM(INVENQTY) AS INVENQTY ");
  // $sb->append(" ( ");
  // $sb->append(" SELECT SUM(INVENQTY) AS INVENQTY ");
  // $sb->append(" FROM SFCMONTHINVENTORY ");
  // $sb->append(" WHERE PARTCD=QST60_PARTCD and PALLETID=QST60_PALLET ");
  // $sb->append(" )AS INVENQTY ");
  // $sb->append(" ( ");
  // $sb->append(" SELECT YYYYMM ");
  // $sb->append(" FROM SFCMONTHINVENTORY ");
  // $sb->append(" WHERE PARTCD=QST60_PARTCD and PALLETID=QST60_PALLET and rownum <=1 ");
  // $sb->append(" GROUP BY YYYYMM, PARTCD, PALLETID)AS YYYYMM ");
	$sb->append(" FROM SFCMONTHINVENTORY  ");
	$sb->append(" LEFT JOIN QSTBOMPC ");
	$sb->append(" ON PARTCD=QST22_PARTCD  ");
  $sb->append(" WHERE 1=1 ");

  if($SHELF_ID != ''){
    $sb->append(sprintf(" AND PALLETID='%s' ", $SHELF_ID));
  }
  if($PARTCODE != ''){
    $sb->append(sprintf(" AND PARTCD='%s' ", $PARTCODE));
  }

  if($PLACE != ''){
    $sb->append(sprintf(" AND WHCODE='%s' ", $PLACE));
  }

  if($date != ''){
    $sb->append(sprintf(" AND YYYYMM='%s' ", $date));
  }

  //$sb->append(" AND rownum <=2000 ");

  $sb->append(" GROUP BY PARTCD,PALLETID,QST22_PARTNM,YYYYMM  ");
  $sb->append(" ORDER BY PALLETID, PARTCD ");



	$strSQL = $sb->toString();
  //print $strSQL;
	$stmt = oci_parse($conn, $strSQL);

	if(!oci_execute($stmt)){
		$e = oci_error($stmt);
		$response['e']	= true;
		$response['msg'] = $e;

		echo json_encode($response);
		exit(0);
	}

	while (($row = oci_fetch_object($stmt)) != false) {
		// Use upper case attribute names for each standard Oracle column
		$item['PARTCD'] = $row->PARTCD;
    $item['QST22_PARTNM'] = $row->QST22_PARTNM;
		$item['MID_QUANTITY'] = $row->MID_QUANTITY;
    $item['PALLETID']     = $row->PALLETID;
    $item['YYYYMM']       = $row->YYYYMM;
		$item['INVENQTY'] 	  = number_format($row->INVENQTY);

		$items[] = $item;
		$item = array();
	}
  $response['aaData'] = $items;

  return $response;
}
?>
