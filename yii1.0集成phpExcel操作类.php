<?php
/**
 * excel处理类
 * @author Administrator
 *
 */
class OperateExcel{
	
	
	/**
	 * 导入excel
	 * @param unknown $files
	 * @return unknown
	 */
	public function importNormalExcel($files){
		// 关闭YII的自动加载功能，改用手动加载，否则会出错，PHPExcel有自己的自动加载功能
		Yii::$enableIncludePath=false;
		/*引入PHPExcel.php文件*/
		Yii::import('application.extensions.PHPExcel.PHPExcel', 1);
		//下面是Excel数据导出处理逻辑
		$extend = pathinfo($files['post_file']['name']);
		$extend = strtolower($extend["extension"]);
		$filePath = 'tmp/excel/';
		$extend2 = strrchr ($files['post_file']['name'],'.');
		//上传后的文件名
		$name = "amazon_import_".time().$extend2;
		$uploadfile=$filePath.$name;//上传后的文件名地址
		$moveRes = move_uploaded_file($files['post_file']['tmp_name'] , $uploadfile);//假如上传到当前目录下
		$extend=='xlsx' ? $reader_type='Excel2007' : $reader_type='Excel5';
		$objReader = PHPExcel_IOFactory::createReader($reader_type);//use excel2007 for 2007 format
		$objPHPExcel = $objReader->load($uploadfile);
		$sheet = $objPHPExcel->getSheet(0); 
        $highestRow = $sheet->getHighestRow();           //取得总行数 
        $highestColumn = $sheet->getHighestColumn(); //取得总列数
        $dataArr = array();
        for($j=1;$j<=$highestRow;$j++)                        //从第一行开始读取数据
        {
        	$valArr = array();
        	for($k='A';$k<=$highestColumn;$k++)            //从A列读取数据
        	{
        		$valArr[] = $objPHPExcel->getActiveSheet()->getCell("$k$j")->getValue();
        	}
        	$dataArr[] = $valArr;
        }
		return $dataArr;
	}
	
	/**
	 * 导入excel
	 * @param unknown $files
	 * @return unknown
	 */
	public function importExcel($files){
        // 关闭YII的自动加载功能，改用手动加载，否则会出错，PHPExcel有自己的自动加载功能
	    Yii::$enableIncludePath=false;                
	     /*引入PHPExcel.php文件*/        
	    Yii::import('application.extensions.PHPExcel.PHPExcel', 1);  
        //下面是Excel数据导出处理逻辑
		$extend = pathinfo($files['post_file']['name']);
		$extend = strtolower($extend["extension"]);
		$filePath = 'tmp/excel/';
		$extend2 = strrchr ($files['post_file']['name'],'.');
		//上传后的文件名
		$name = "amazon_import_".time().$extend2;
		$uploadfile=$filePath.$name;//上传后的文件名地址
		$moveRes = move_uploaded_file($files['post_file']['tmp_name'] , $uploadfile);//假如上传到当前目录下
		$extend=='xlsx' ? $reader_type='Excel2007' : $reader_type='Excel5';
		$objReader = PHPExcel_IOFactory::createReader($reader_type);//use excel2007 for 2007 format
		$objPHPExcel = $objReader->load($uploadfile);
		$sheetCount = $objPHPExcel->getSheetCount();
		$dataArr = array();
		for ($i = 0; $i < $sheetCount; $i++) {
			$inData = $objPHPExcel->getSheet($i)->toArray();
			if(empty($dataArr)){
				$dataArr = $inData;
			}else{
				array_shift($inData);
				$dataArr = array_merge($dataArr , $inData);
			}
		}
		return $dataArr;
	}
	
        public function importExcel2($files,$num){
        // 关闭YII的自动加载功能，改用手动加载，否则会出错，PHPExcel有自己的自动加载功能
	    Yii::$enableIncludePath=false;                
	     /*引入PHPExcel.php文件*/        
	    Yii::import('application.extensions.PHPExcel.PHPExcel', 1);  
        //下面是Excel数据导出处理逻辑
		$extend = pathinfo($files['post_file']['name']);
		$extend = strtolower($extend["extension"]);
		$filePath = 'tmp/excel/';
		$extend2 = strrchr ($files['post_file']['name'],'.');
		//上传后的文件名
		$name = "amazon_import_".time().$extend2;
		$uploadfile=$filePath.$name;//上传后的文件名地址
		$moveRes = move_uploaded_file($files['post_file']['tmp_name'] , $uploadfile);//假如上传到当前目录下
		$extend=='xlsx' ? $reader_type='Excel2007' : $reader_type='Excel5';
		$objReader = PHPExcel_IOFactory::createReader($reader_type);//use excel2007 for 2007 format
		$objPHPExcel = $objReader->load($uploadfile);
               
                $inData1 = $objPHPExcel->getSheet(0)->toArray();
                $inData2 = $objPHPExcel->getSheet(1)->toArray();
                $inData = array('dataa'=>$inData1,'datab'=>$inData2);                
		return $inData;
	}
        
	/**
	 * 导出excel
	 */
	public function exportExcel($listArr , $headArr){
		//if(!$listArr || !$headArr) return false;
		set_time_limit(0);
		// 关闭YII的自动加载功能，改用手动加载，否则会出错，PHPExcel有自己的自动加载功能
		Yii::$enableIncludePath=false;
		/*引入PHPExcel.php文件*/
		Yii::import('application.extensions.PHPExcel.PHPExcel', 1);
		
		$obj_phpexcel = new PHPExcel();
		
		$objActSheet = $obj_phpexcel->getActiveSheet();
		//设置字体
		$objActSheet->getDefaultStyle()->getFont()->setName('微软雅黑');
		//设置字体大小
		$objActSheet->getDefaultStyle()->getFont()->setSize(9);
		//设置自动行高
		$objActSheet->getDefaultRowDimension()->setRowHeight(15);
		//设置自动列宽
		$objActSheet->getDefaultColumnDimension()->setWidth(20);
		$key = 0;
		foreach($headArr as $v){
			$index = PHPExcel_Cell::stringFromColumnIndex($key);
			//设置页头填充色
			$objStyle = $objActSheet ->getStyle($index.'1');
			$objFill = $objStyle->getFill();
			$objFill->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
			$objFill->getStartColor()->setARGB('#FFFF00');
				
			//设置页头边框色
			$objBorder = $objStyle->getBorders();
			$objBorder->getTop()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
			$objBorder->getBottom()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
			$objBorder->getLeft()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
			$objBorder->getRight()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
			
			//设置对其方式
			$objAlign = $objStyle->getAlignment();
			$objAlign->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objAlign->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
			
			if($key==5){
				$to_index = PHPExcel_Cell::stringFromColumnIndex($key+7);
				$obj_phpexcel->setActiveSheetIndex(0)->setCellValue($index.'1' , '商品描述');
				$objActSheet->mergeCells($index.'1:'.$to_index.'1');
			}elseif($key==19){
				$to_index = PHPExcel_Cell::stringFromColumnIndex($key+7);
				$obj_phpexcel->setActiveSheetIndex(0)->setCellValue($index.'1' , '自贸专区商品完成区内备案后必填');
				$objActSheet->mergeCells($index.'1:'.$to_index.'1');
			}else{
				$obj_phpexcel->setActiveSheetIndex(0)->setCellValue($index.'1' , $v);
			}
			$key++;
		}
		$key = 0;
		$keyArr = array();
		foreach($headArr as $v){
			if($key <6  || ($key>13 && $key <20) || $key>27){
				$objActSheet->mergeCells($index.'1:'.$index.'2');
			}
			$index = PHPExcel_Cell::stringFromColumnIndex($key);
			$keyArr[] = $index;
			//设置页头填充色
			$objStyle = $objActSheet ->getStyle($index.'2');
			$objFill = $objStyle->getFill();
			$objFill->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
			$objFill->getStartColor()->setARGB('#FFFF00');
			
			//设置页头边框色
			$objBorder = $objStyle->getBorders();
			$objBorder->getTop()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
			$objBorder->getBottom()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
			$objBorder->getLeft()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
			$objBorder->getRight()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
			
			//设置对其方式
			$objAlign = $objStyle->getAlignment();
			$objAlign->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objAlign->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
			
			$obj_phpexcel->setActiveSheetIndex(0)->setCellValue($index.'2' , $v);
			$key++;
	 	}
	 	$keyWordArr = CDict::$title_key_word;
		$key2 = 3;
		foreach($listArr as $l_key =>$info){ //行写入
			$key3 = 0;
			foreach($info as $c_key=>$value){// 列写入
				$span = $keyArr[$key3];
				//设置对其方式
				//设置关键字高亮
				if($c_key == 'NAMEZH'){
					foreach ($keyWordArr as $word){
						if(strpos($value, $word)!==false){
							$objStyle = $objActSheet ->getStyle($span.$key2);
							$objStyle->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_RED);
							$objStyle->getFont()->setSize(10);
						}
					}
				}
// 				$objAlign = $objStyle->getAlignment();
// 				$objAlign->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
// 				$objAlign->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
 				//设置单元格避免大数字科学计数显示
// 				$objStyle = $objActSheet ->getStyle($span.$key2);
// 				$objStyle->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
				if(is_numeric($value) && $c_key != 'GOODNO' && $c_key != 'HSCODE' && $c_key != 'GENERHSCODE' && $c_key != 'MODEL' && $c_key != 'RECEIVEZIPCODE' && $c_key != 'RECZIP'){
					$objStyle = $objActSheet ->getStyle($span.$key2);
					if(strpos($value,'.')){
						$objStyle->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00);
					}else{
						$objStyle->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
					}
					$objActSheet->setCellValue($span.$key2, $value);
				}else{
					$objActSheet->setCellValueExplicit($span.$key2,$value,PHPExcel_Cell_DataType::TYPE_STRING);
				}
				//$objActSheet->setCellValueExplicit($span.$key2,$value,PHPExcel_Cell_DataType::TYPE_STRING);
				//$objActSheet->setCellValue($span.$key2, $value);
				$key3++;
			}
			$key2++;
		}
		//$obj_Writer = PHPExcel_IOFactory::createWriter($obj_phpexcel,'Excel5');
		$obj_Writer=new PHPExcel_Writer_Excel5($obj_phpexcel);//用于其他版本格式
		$date = date("YmdHis");
		$filename = "export_record_".$date.".xls";
		ob_end_clean();//清除缓冲区,避免乱码
		header("Content-Type: application/force-download");
		header("Content-Type: application/octet-stream");
		header("Content-Type: application/download");
		header('Content-Disposition:inline;filename="'.$filename.'"');
		header("Content-Transfer-Encoding: binary");
		header("Last-Modified: " . gmdate("D, d M Y H:i:s") . " GMT");
		header("Cache-Control: must-revalidate, post-check=0, pre-check=0");
		header("Pragma: no-cache");
		$obj_Writer->save('php://output');
	}
	
	/**
	 * 生成一般excel
	 */
	public function exportNormalExcel($listArr , $headArr , $file_tag = '' , $path=''){
		//if(!$listArr || !$headArr) return false;
		set_time_limit(0);
		// 关闭YII的自动加载功能，改用手动加载，否则会出错，PHPExcel有自己的自动加载功能
		Yii::$enableIncludePath=false;
		/*引入PHPExcel.php文件*/
		Yii::import('application.extensions.PHPExcel.PHPExcel', 1);
		
		$obj_phpexcel = new PHPExcel();
		
		$objActSheet = $obj_phpexcel->getActiveSheet();
		//设置字体
		$objActSheet->getDefaultStyle()->getFont()->setName('微软雅黑');
		//设置字体大小
		$objActSheet->getDefaultStyle()->getFont()->setSize(9);
		//设置自动行高
		//$objActSheet->getDefaultRowDimension()->setRowHeight(15);
		$key = 0;
		$keyArr = array();
		foreach($headArr as $v){
			$index = PHPExcel_Cell::stringFromColumnIndex($key);
			$keyArr[] = $index;
			//设置页头填充色
			$objStyle = $objActSheet ->getStyle($index.'1');
			$objStyle->getAlignment()->setWrapText(true);
			$objFill = $objStyle->getFill();
			$objFill->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
			$objFill->getStartColor()->setARGB('#FFFF00');
			//设置页头边框色
			$objBorder = $objStyle->getBorders();
			$objBorder->getTop()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
			$objBorder->getBottom()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
			$objBorder->getLeft()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
			$objBorder->getRight()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
			//设置自动列宽
			$objActSheet->getColumnDimension($index)->setAutoSize(true);
			//设置对其方式
			$objAlign = $objStyle->getAlignment();
			$objAlign->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objAlign->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
			$obj_phpexcel->setActiveSheetIndex(0)->setCellValue($index.'1' , $v);
			$key++;
		}
		$keyWordArr = CDict::$title_key_word;
		$key2 = 2;
		foreach($listArr as $l_key =>$info){ //行写入
			$key3 = 0;
			foreach($info as $c_key=>$value){// 列写入
				$span = $keyArr[$key3];
				if($c_key == 'NAMEZH'){
					//设置关键字高亮
					foreach ($keyWordArr as $word){
						if(strpos($value, $word)!==false){
							$objStyle = $objActSheet ->getStyle($span.$key2);
							$objStyle->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_RED);
							$objStyle->getFont()->setSize(10);
						}
					}
				}
				$objActSheet->setCellValueExplicit($span.$key2,$value,PHPExcel_Cell_DataType::TYPE_STRING);
				//$objActSheet->setCellValue($span.$key2, $value);
				$key3++;
			}
			$key2++;
		}
		$obj_Writer=new PHPExcel_Writer_Excel5($obj_phpexcel);//用于其他版本格式
		$date = date("YmdHis");
		$filename = $file_tag."export_".$date.".xls";
		ob_end_clean();//清除缓冲区,避免乱码
		if($path){
		    $obj_Writer->save($path.$filename);
		    return $path.$filename;
		}else{
		    header("Content-Type: application/force-download");
		    header("Content-Type: application/octet-stream");
		    header("Content-Type: application/download");
		    header('Content-Disposition:inline;filename="'.$filename.'"');
		    header("Content-Transfer-Encoding: binary");
		    header("Last-Modified: " . gmdate("D, d M Y H:i:s") . " GMT");
		    header("Cache-Control: must-revalidate, post-check=0, pre-check=0");
		    header("Pragma: no-cache");
		    $obj_Writer->save('php://output');
		}
	}
	
	
	/**
	 * 生成一般excel
	 */
	public function exportFileNormalExcel($listArr , $headArr , $dirName , $file_tag = ''){
	    //if(!$listArr || !$headArr) return false;
	    set_time_limit(0);
	    // 关闭YII的自动加载功能，改用手动加载，否则会出错，PHPExcel有自己的自动加载功能
	    Yii::$enableIncludePath=false;
	    /*引入PHPExcel.php文件*/
	    Yii::import('application.extensions.PHPExcel.PHPExcel', 1);
	
	    $obj_phpexcel = new PHPExcel();
	
	    $objActSheet = $obj_phpexcel->getActiveSheet();
	    //设置字体
	    $objActSheet->getDefaultStyle()->getFont()->setName('微软雅黑');
	    //设置字体大小
	    $objActSheet->getDefaultStyle()->getFont()->setSize(9);
	    //设置自动行高
	    $objActSheet->getDefaultRowDimension()->setRowHeight(15);
	    $key = 0;
	    $keyArr = array();
	    foreach($headArr as $v){
	        $index = PHPExcel_Cell::stringFromColumnIndex($key);
	        $keyArr[] = $index;
	        //设置页头填充色
	        $objStyle = $objActSheet ->getStyle($index.'1');
	        $objFill = $objStyle->getFill();
	        $objFill->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
	        $objFill->getStartColor()->setARGB('#FFFF00');
	        //设置页头边框色
	        $objBorder = $objStyle->getBorders();
	        $objBorder->getTop()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
	        $objBorder->getBottom()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
	        $objBorder->getLeft()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
	        $objBorder->getRight()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
	        //设置自动列宽
	        $objActSheet->getColumnDimension($index)->setAutoSize(true);
	        //设置对其方式
	        $objAlign = $objStyle->getAlignment();
	        $objAlign->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	        $objAlign->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	        $obj_phpexcel->setActiveSheetIndex(0)->setCellValue($index.'1' , $v);
	        $key++;
	    }
	    $keyWordArr = CDict::$title_key_word;
	    $key2 = 2;
	    foreach($listArr as $l_key =>$info){ //行写入
	        $key3 = 0;
	        foreach($info as $c_key=>$value){// 列写入
	            $span = $keyArr[$key3];
	            if($c_key == 'NAMEZH'){
	                //设置关键字高亮
	                foreach ($keyWordArr as $word){
	                    if(strpos($value, $word)!==false){
	                        $objStyle = $objActSheet ->getStyle($span.$key2);
	                        $objStyle->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_RED);
	                        $objStyle->getFont()->setSize(10);
	                    }
	                }
	            }
	            if(is_numeric($value) && $c_key != 'GOODNO' && $c_key != 'HSCODE' && $c_key != 'RECEIVEZIPCODE' && $c_key != 'RECZIP'){
	                $objStyle = $objActSheet ->getStyle($span.$key2);
	                if(strpos($value,'.')){
	                    $objStyle->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00);
	                }else{
	                    $objStyle->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
	                }
	                $objActSheet->setCellValue($span.$key2, $value);
	            }else{
	                $objActSheet->setCellValueExplicit($span.$key2,$value,PHPExcel_Cell_DataType::TYPE_STRING);
	            }
	            //$objActSheet->setCellValue($span.$key2, $value);
	            $key3++;
	        }
	        $key2++;
	    }
	    $obj_Writer=new PHPExcel_Writer_Excel5($obj_phpexcel);//用于其他版本格式
	    $filename = $file_tag.".xls";
	    ob_end_clean();//清除缓冲区,避免乱码
	    
	    $obj_Writer->save($dirName."/".$filename);
	}
	
	
	
	/**
	 * 生成一般excel
	 */
	public function exportAbnormalOrderExcel($listArr ,$listArr2 , $headArr , $headArr2 , $file_tag = ''){
	    //if(!$listArr || !$headArr) return false;
	    set_time_limit(0);
	    // 关闭YII的自动加载功能，改用手动加载，否则会出错，PHPExcel有自己的自动加载功能
	    Yii::$enableIncludePath=false;
	    /*引入PHPExcel.php文件*/
	    Yii::import('application.extensions.PHPExcel.PHPExcel', 1);
	
	    $obj_phpexcel = new PHPExcel();
	
	    $objActSheet = $obj_phpexcel->getActiveSheet();
	    //设置字体
	    $objActSheet->getDefaultStyle()->getFont()->setName('微软雅黑');
	    //设置字体大小
	    $objActSheet->getDefaultStyle()->getFont()->setSize(9);
	    //设置自动行高
	    $objActSheet->getDefaultRowDimension()->setRowHeight(15);
	    $key = 0;
	    $keyArr = array();
	    foreach($headArr as $v){
	        $index = PHPExcel_Cell::stringFromColumnIndex($key);
	        $keyArr[] = $index;
	        //设置页头填充色
	        $objStyle = $objActSheet ->getStyle($index.'1');
	        $objFill = $objStyle->getFill();
	        $objFill->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
	        $objFill->getStartColor()->setARGB('#FFFF00');
	        //设置页头边框色
	        $objBorder = $objStyle->getBorders();
	        $objBorder->getTop()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
	        $objBorder->getBottom()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
	        $objBorder->getLeft()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
	        $objBorder->getRight()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
	        //设置自动列宽
	        $objActSheet->getColumnDimension($index)->setAutoSize(true);
	        //设置对其方式
	        $objAlign = $objStyle->getAlignment();
	        $objAlign->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	        $objAlign->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	        $obj_phpexcel->setActiveSheetIndex(0)->setCellValue($index.'1' , $v);
	        $key++;
	    }
	    $keyWordArr = CDict::$title_key_word;
	    $key2 = 2;
	    foreach($listArr as $l_key =>$info){ //行写入
	        $key3 = 0;
	        foreach($info as $c_key=>$value){// 列写入
	            $span = $keyArr[$key3];
	            $objStyle = $objActSheet ->getStyle($span.$key2);
	            //设置边框色
	            $objBorder = $objStyle->getBorders();
	            $objBorder->getTop()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
	            $objBorder->getBottom()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
	            $objBorder->getLeft()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
	            $objBorder->getRight()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
	            
	            $objAlign = $objStyle->getAlignment();
	            $objAlign->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	            $objAlign->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	            if($c_key == 'NAMEZH'){
	                //设置关键字高亮
	                foreach ($keyWordArr as $word){
	                    if(strpos($value, $word)!==false){
	                        $objStyle = $objActSheet ->getStyle($span.$key2);
	                        $objStyle->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_RED);
	                        $objStyle->getFont()->setSize(10);
	                    }
	                }
	            }
	            if($info['ASSBILLNO'] == '' && in_array($c_key, array('BATCHNO' , 'GATHERBillNO','ASSBILLNO','PAYTOTALPRICE' ,'ABNORMAL'))){
	                $objActSheet->mergeCells($span.($key2-1).':'.$span.$key2);
	            }
	            if(in_array($c_key, array('GATHERBillNO','ASSBILLNO','ABNORMAL' , 'NAMEZH' , 'GOODNO'))){
	                $objActSheet->setCellValueExplicit($span.$key2,$value,PHPExcel_Cell_DataType::TYPE_STRING);
	            }elseif($c_key == 'PAYTOTALPRICE' || $c_key == 'CNYTOTALMONEY'){
	                $objStyle->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00);
	                $objActSheet->setCellValue($span.$key2, $value);
	            }elseif($c_key == 'AMOUNT' || $c_key == 'BATCHNO'){
	                $objStyle->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
	                $objActSheet->setCellValue($span.$key2, $value);
	            }
	            $key3++;
	        }
	        $key2++;
	    }

	    //创建第二个工作表
		$newWorkSheet = new PHPExcel_Worksheet($obj_phpexcel, 'orderlist'); //创建一个工作表
		$obj_phpexcel->addSheet($newWorkSheet); //插入工作表
		$obj_phpexcel->setActiveSheetIndex(1); //切换到新创建的工作表
		
		$objActSheet1 = $obj_phpexcel->getActiveSheet();
		//设置字体
		$objActSheet1->getDefaultStyle()->getFont()->setName('微软雅黑');
		//设置字体大小
		$objActSheet1->getDefaultStyle()->getFont()->setSize(9);
		//设置自动行高
		$objActSheet1->getDefaultRowDimension()->setRowHeight(15);
		$key = 0;
		$keyArr = array();
		foreach($headArr2 as $v){
			$index = PHPExcel_Cell::stringFromColumnIndex($key);
			$keyArr[] = $index;
			//设置页头填充色
			$objStyle = $objActSheet1 ->getStyle($index.'1');
			$objFill = $objStyle->getFill();
			$objFill->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
			$objFill->getStartColor()->setARGB('#FFFF00');
			//设置页头边框色
			$objBorder = $objStyle->getBorders();
			$objBorder->getTop()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
			$objBorder->getBottom()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
			$objBorder->getLeft()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
			$objBorder->getRight()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
			//设置自动列宽
			$objActSheet1->getColumnDimension($index)->setAutoSize(true);
			//设置对其方式
			$objAlign = $objStyle->getAlignment();
			$objAlign->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objAlign->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
			$obj_phpexcel->setActiveSheetIndex(1)->setCellValue($index.'1' , $v);
			$key++;
		}
		$key2 = 2;
		foreach($listArr2 as $l_key =>$info){ //行写入
			$key3 = 0;
			foreach($info as $c_key=>$value){// 列写入
				$span = $keyArr[$key3];
				//设置关键字高亮
				if($c_key == 'NAMEZH'){
					foreach ($keyWordArr as $word){
						if(strpos($value, $word)!==false){
							$objStyle = $objActSheet1 ->getStyle($span.$key2);
							$objStyle->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_RED);
							$objStyle->getFont()->setSize(10);
						}
					}
				}
				// 				$objStyle = $objActSheet ->getStyle($span.$key2);
				// 				$objStyle->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
				//设置单元格避免大数字科学计数显示
// 				$objStyle = $objActSheet1 ->getStyle($span.$key2);
// 				$objStyle->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
				//$objActSheet1->setCellValue($span.$key2, $value);
				if(is_numeric($value) && $c_key != 'GOODNO' && $c_key != 'HSCODE' && $c_key != 'RECEIVEZIPCODE' && $c_key != 'RECZIP'){
					$objStyle = $objActSheet1 ->getStyle($span.$key2);
					if(strpos($value,'.')){
						$objStyle->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00);
					}else{
						$objStyle->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
					}
					$objActSheet1->setCellValue($span.$key2,$value);
				}else{
					$objActSheet1->setCellValueExplicit($span.$key2,$value,PHPExcel_Cell_DataType::TYPE_STRING);
				}
				//$objActSheet1->setCellValueExplicit($span.$key2,$value,PHPExcel_Cell_DataType::TYPE_STRING);
				$key3++;
			}
			$key2++;
		}
		$obj_phpexcel->setActiveSheetIndex(0); //设置第一个工作表为活动工作表

	    $obj_Writer=new PHPExcel_Writer_Excel5($obj_phpexcel);//用于其他版本格式
	    $date = date("YmdHis");
	    $filename = $file_tag."export_".$date.".xls";
	    ob_end_clean();//清除缓冲区,避免乱码
	    header("Content-Type: application/force-download");
	    header("Content-Type: application/octet-stream");
	    header("Content-Type: application/download");
	    header('Content-Disposition:inline;filename="'.$filename.'"');
	    header("Content-Transfer-Encoding: binary");
	    header("Last-Modified: " . gmdate("D, d M Y H:i:s") . " GMT");
	    header("Cache-Control: must-revalidate, post-check=0, pre-check=0");
	    header("Pragma: no-cache");
	    $obj_Writer->save('php://output');
	}
	
	
	/**
	 * 生成多个sheet的excel
	 */
	public function exportSheetExcel($listArr1 , $listArr2 , $headArr1 , $headArr2){
		//if(!$listArr || !$headArr) return false;
		set_time_limit(0);
		// 关闭YII的自动加载功能，改用手动加载，否则会出错，PHPExcel有自己的自动加载功能
		Yii::$enableIncludePath=false;
		/*引入PHPExcel.php文件*/
		Yii::import('application.extensions.PHPExcel.PHPExcel', 1);
	
		$obj_phpexcel = new PHPExcel();
		$obj_phpexcel->setActiveSheetIndex(0); //设置第一个工作表为活动工作表
		$obj_phpexcel->getActiveSheet()->setTitle('orderhead'); //设置工作表名称
		$objActSheet = $obj_phpexcel->getActiveSheet();
		//设置字体
		$objActSheet->getDefaultStyle()->getFont()->setName('微软雅黑');
		//设置字体大小
		$objActSheet->getDefaultStyle()->getFont()->setSize(9);
		//设置自动行高
		$objActSheet->getDefaultRowDimension()->setRowHeight(15);
		$key = 0;
		$keyArr = array();
		foreach($headArr1 as $v){
			$index = PHPExcel_Cell::stringFromColumnIndex($key);
			$keyArr[] = $index;
			//设置页头填充色
			$objStyle = $objActSheet ->getStyle($index.'1');
			$objFill = $objStyle->getFill();
			$objFill->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
			$objFill->getStartColor()->setARGB('#FFFF00');
			//设置页头边框色
			$objBorder = $objStyle->getBorders();
			$objBorder->getTop()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
			$objBorder->getBottom()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
			$objBorder->getLeft()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
			$objBorder->getRight()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
			//设置自动列宽
			$objActSheet->getColumnDimension($index)->setAutoSize(true);
			//设置对其方式
			$objAlign = $objStyle->getAlignment();
			$objAlign->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objAlign->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
			$obj_phpexcel->setActiveSheetIndex(0)->setCellValue($index.'1' , $v);
			$key++;
		}
		$keyWordArr = CDict::$title_key_word;
		$keyWordArr2 = CDict::$title_hscode_word;
		$key2 = 2;
		foreach($listArr1 as $l_key =>$info){ //行写入
			$key3 = 0;
			foreach($info as $c_key=>$value){// 列写入
				$span = $keyArr[$key3];
				// 				$objStyle = $objActSheet ->getStyle($span.$key2);
				// 				$objStyle->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
				//设置单元格避免大数字科学计数显示
// 				if(in_array($c_key, array('ORDERCOMMITTIME','PAY_TIME','USER_ID','IDENTIFY_CODE'))){
// 					$objStyle = $objActSheet ->getStyle($span.$key2);
// 	 				$objStyle->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
// 				}
				//$objActSheet->setCellValue($span.$key2, $value);
				if(is_numeric($value) && $c_key != 'TRX_SERIAL_NO' && $c_key!='IDENTIFY_CODE' && $c_key!='USER_ID' && $c_key != 'GOODNO' && $c_key != 'HSCODE' && $c_key != 'RECEIVEZIPCODE' && $c_key != 'RECZIP' && $c_key != 'MERCHANTORDERID'){
					$objStyle = $objActSheet ->getStyle($span.$key2);
					if(strpos($value,'.')){
						$objStyle->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00);
					}else{
						$objStyle->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
					}
					$objActSheet->setCellValue($span.$key2,$value);
				}else{
					$objActSheet->setCellValueExplicit($span.$key2,$value,PHPExcel_Cell_DataType::TYPE_STRING);
				}
				//$objActSheet->setCellValueExplicit($span.$key2,$value,PHPExcel_Cell_DataType::TYPE_STRING);
				$key3++;
			}
			$key2++;
		}
		
		//创建第二个工作表
		$newWorkSheet = new PHPExcel_Worksheet($obj_phpexcel, 'orderlist'); //创建一个工作表
		$obj_phpexcel->addSheet($newWorkSheet); //插入工作表
		$obj_phpexcel->setActiveSheetIndex(1); //切换到新创建的工作表
		
		$objActSheet1 = $obj_phpexcel->getActiveSheet();
		//设置字体
		$objActSheet1->getDefaultStyle()->getFont()->setName('微软雅黑');
		//设置字体大小
		$objActSheet1->getDefaultStyle()->getFont()->setSize(9);
		//设置自动行高
		$objActSheet1->getDefaultRowDimension()->setRowHeight(15);
		$key = 0;
		$keyArr = array();
		foreach($headArr2 as $v){
			$index = PHPExcel_Cell::stringFromColumnIndex($key);
			$keyArr[] = $index;
			//设置页头填充色
			$objStyle = $objActSheet1 ->getStyle($index.'1');
			$objFill = $objStyle->getFill();
			$objFill->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
			$objFill->getStartColor()->setARGB('#FFFF00');
			//设置页头边框色
			$objBorder = $objStyle->getBorders();
			$objBorder->getTop()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
			$objBorder->getBottom()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
			$objBorder->getLeft()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
			$objBorder->getRight()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
			//设置自动列宽
			$objActSheet1->getColumnDimension($index)->setAutoSize(true);
			//设置对其方式
			$objAlign = $objStyle->getAlignment();
			$objAlign->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objAlign->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
			$obj_phpexcel->setActiveSheetIndex(1)->setCellValue($index.'1' , $v);
			$key++;
		}
		$key2 = 2;
		foreach($listArr2 as $l_key =>$info){ //行写入
			$key3 = 0;
			foreach($info as $c_key=>$value){// 列写入
				$span = $keyArr[$key3];
				//设置关键字高亮
				if($c_key == 'NAMEZH'){
					foreach ($keyWordArr as $word){
						if(strpos($value, $word)!==false){
							$objStyle = $objActSheet1 ->getStyle($span.$key2);
							$objStyle->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_RED);
							$objStyle->getFont()->setSize(10);
						}
					}
				}
				if($c_key == 'HSCODE'){
				    //设置关键字高亮
				    foreach ($keyWordArr2 as $word2){
				        if(strpos($value, $word2)!==false){
				            $objStyle = $objActSheet1 ->getStyle($span.$key2);
							$objStyle->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_RED);
							$objStyle->getFont()->setSize(10);
				        }
				    }
				}
				// 				$objStyle = $objActSheet ->getStyle($span.$key2);
				// 				$objStyle->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
				//设置单元格避免大数字科学计数显示
// 				$objStyle = $objActSheet1 ->getStyle($span.$key2);
// 				$objStyle->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
				//$objActSheet1->setCellValue($span.$key2, $value);
				if(is_numeric($value) && $c_key != 'GOODNO' && $c_key != 'HSCODE' && $c_key != 'RECEIVEZIPCODE' && $c_key != 'RECZIP'){
					$objStyle = $objActSheet1 ->getStyle($span.$key2);
					if(strpos($value,'.')){
						$objStyle->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00);
					}else{
						$objStyle->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
					}
					$objActSheet1->setCellValue($span.$key2,$value);
				}else{
					$objActSheet1->setCellValueExplicit($span.$key2,$value,PHPExcel_Cell_DataType::TYPE_STRING);
				}
				//$objActSheet1->setCellValueExplicit($span.$key2,$value,PHPExcel_Cell_DataType::TYPE_STRING);
				$key3++;
			}
			$key2++;
		}
		$obj_phpexcel->setActiveSheetIndex(0); //设置第一个工作表为活动工作表
		
		$obj_Writer=new PHPExcel_Writer_Excel5($obj_phpexcel);//用于其他版本格式
		$date = date("YmdHis");
		$filename = "export_".$date.".xls";
		ob_end_clean();//清除缓冲区,避免乱码
		header("Content-Type: application/force-download");
		header("Content-Type: application/octet-stream");
		header("Content-Type: application/download");
		header('Content-Disposition:inline;filename="'.$filename.'"');
		header("Content-Transfer-Encoding: binary");
		header("Last-Modified: " . gmdate("D, d M Y H:i:s") . " GMT");
		header("Cache-Control: must-revalidate, post-check=0, pre-check=0");
		header("Pragma: no-cache");
		$obj_Writer->save('php://output');
	}
}
?>
