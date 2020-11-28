<?php
namespace tp5er;

use PHPExcel_IOFactory;
use PHPExcel;

class easyExcel{

   
    private $excelExt='.xlsx';//保存文件后缀
    private $excelPath;//文件保存的绝对位置
    private $letter=["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"];
    //Excel5和Excel2007
    private $ExcelVersion=['Excel2007','Excel5'];

    private $sheetNum=0;
    //phpExcel实例化对象
    private $phpExcel;
    private $phpWriter;
    private $xlsReader;
    private $phpSheet;

    public function __construct()
    {
        //  实例化PHPExcel类
        $this->phpExcel = new PHPExcel();
       
    }

    /**
     * 创建新的Sheet 支持链式操作
     * @param string $sheet_title
     * @param array  $data       导出数据内容
     * @param array  $excelHeader导出表头
     * @return $this
     * @throws \Exception
     * @throws \PHPExcel_Exception
     */
    public function createSheet($sheet_title='Sheet1',$data=[],$excelHeader=[])
    {
        if ( empty($excelHeader)||!is_array($excelHeader)){
            throw new \Exception("Parameter is incorrect");
            return $this;
        }
        $sheet_num = $this->getNewSheetNum();
        $objPHPExcel=$this->phpExcel;
        $objPHPExcel->createSheet($sheet_num);
        //设置当前的sheet
        $objPHPExcel->setActiveSheetIndex($sheet_num);
        //设置sheet的name
        $objPHPExcel->getActiveSheet()->setTitle($sheet_title);
        $sheet=$objPHPExcel->getActiveSheet();
       //表头设置
        $excelHeader=array_values($excelHeader);
        foreach($excelHeader as $item=>$value){
            $sheet->setCellValue($this->letter[$item]."1",$value);
        }
       //表内容设置
        foreach($data as $item=>$value ){
            $value=array_values($value);
            foreach($value as $i=>$v)
            $sheet->setCellValue($this->letter[$i].($item+2),$value[$i]);
        }
        return $this;
    }


    /**
     * 导出下载
     * @return output  
     */ 
    public function downFile($excelName='')
    {
        ob_start();
        if(empty($excelName)){
            $excelName = 'JQExcel'.date("Ymdhis");
        }
        try{
            try{
              $this->phpWriter = PHPExcel_IOFactory::createWriter($this->phpExcel,$this->ExcelVersion[0]);
            }catch(Exception $e){
              $this->phpWriter = PHPExcel_IOFactory::createWriter($this->phpExcel,$this->ExcelVersion[1]);
            }   
        }catch(Exception $e){
            throw new \Exception("Export failed");
        }
        header('Content-Type: application/vnd.ms-excel; charset=utf-8');
        header("Content-Disposition: attachment;filename=\"$excelName.$this->excelExt\"");
        header('Cache-Control: max-age=0');
        $this->phpWriter->save('php://output');  
        ob_end_flush();
        die();
    }
    /**
     * 导出保存服务器上
     * @param  String  $path 保存路径
     * @param  boolean $activate 自定义保存路径需要将此处设置为true
     * @return Object  
     */ 
    public function saveFile($excelName='',$filepath='',$activate=false)
    {
        
        if (empty($filepath) || $activate==false) {
            $filepath  =  $_SERVER['DOCUMENT_ROOT'].'/jqexcel/';
        }
        if(empty($excelName)){
            $excelName = 'JQExcel'.date("Ymdhis");
        }
        if(!$this->checkPath($filepath)){
            throw new \Exception("The current directory is not writable");
        } else{
            try{
                try{
                  $this->phpWriter = PHPExcel_IOFactory::createWriter($this->phpExcel,$this->ExcelVersion[0]);
                }catch(Exception $e){
                  $this->phpWriter = PHPExcel_IOFactory::createWriter($this->phpExcel,$this->ExcelVersion[1]);
                }   
            }catch(Exception $e){
                throw new \Exception("Export failed");
            }
          
            $this->excelPath=$filepath.$excelName.$this->excelExt;
            $this->phpWriter->save($this->excelPath); 
            return $this;
        }

    }
    /**
     * 导入基本设置
     * @param  String  $path 保存路径
     * @param  boolean $activate 自定义保存路径需要将此处设置为true
     * @return Object  
     * @throws Exception
     * @throws \PHPExcel_Exception
     */ 
    public function loadExcel($filepath)
    {
      if(!is_file($filepath)){
        throw new \Exception("File does not exist");
      }
       try{
            try{
                $xlsReader =  PHPExcel_IOFactory::createReader($this->ExcelVersion[0]);
                $xlsReader->setReadDataOnly(true); 
                $xlsReader->setLoadSheetsOnly(true);
                $this->xlsReader=$xlsReader->load($filepath);
            }catch(Exception $e){
                $xlsReader =  PHPExcel_IOFactory::createReader($this->ExcelVersion[1]);
                $xlsReader->setReadDataOnly(true); 
                $xlsReader->setLoadSheetsOnly(true);
                $this->xlsReader=$xlsReader->load($filepath);
            }
       }catch(Exception $e){
            throw new \Exception("Reading failed");
       }
      return $this;
    }
    /**
     * 导入从对象中获取Sheet数组
     * @return Array  
     */ 
    public function getSheetNames(){
        if (isset($this->xlsReader)){
            return $this->xlsReader->getSheetNames();
        }else{
            return false;
        }
    }
    /**
     * 导入从Sheet中获取数组
     * @param  name  sheet中的key值
     * @return Array 
     */ 
    public function getSheetByName($name){
        if (isset($this->xlsReader)){
            return $this->xlsReader->getSheetByName($name);
        }else{
            return false;
        }

    }
    /**
     * 获取新的Sheet编号
     * @return int
     */
    protected function getNewSheetNum(){
        $sheet_num=$this->sheetNum;
        $this->sheetNum=$sheet_num+1;
        return $sheet_num;
    }

    /**
     * 检查目录是否可写
     * @param  string   $path    目录
     * @return boolean
     */
    protected function checkPath($path)
    {
        if (is_dir($path)) {
            return true;
        }
        if (mkdir($path, 0755, true)) {
            return true;
        } else {
            return false;
        }
    }
    /**
     * 返回数组的维度
     * @param  Array   $arr 任意数组
     * @return number  数组维度
     */
    protected function array_depth($arr)
    {
        if(!is_array($arr)) return 0;
        $max_depth = 0;
        foreach($arr as $item1)
        {
            $t1 = $this->array_depth($item1);
            if( $t1 > $max_depth) $max_depth = $t1;
        }
        return $max_depth + 1;
    }
}
