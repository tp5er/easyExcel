## 预备知识
PHPExcel是一个使用纯PHP编写的库，并提供一组类，允许您编写和读取不同的电子表格文件格式，如Excel（BIFF）.xls，Excel 2007（OfficeOpenXML）.xlsx，CSV，Libre / OpenOffice Calc .ods，Gnumeric，PDF，HTML，...这个项目是围绕Microsoft的OpenXML标准和PHP构建的。


### 文件格式支持（读）

*   BIFF 5-8（.xls）Excel 95及以上
*   Office Open XML（.xlsx）Excel 2007及以上版本
*   SpreadsheetML（.xml）Excel 2003
*   开放文件格式/ OASIS（.ods）
*   Gnumeric的
*   HTML
*   SYLK
*   CSV

### 文件格式支持（写）

*   BIFF 8（.xls）Excel 95及以上
*   Office Open XML（.xlsx）Excel 2007及以上版本
*   HTML
*   CSV
*   PDF（使用tcPDF，DomPDF或mPDF库，需要单独安装）

### 要求

*   PHP版本5.2.0以上
*   PHP扩展名php_zip启用（如果您需要PHPExcel来处理.xlsx .ods或.gnumeric文件，则需要）
*   PHP扩展名php_xml启用
*   PHP扩展名php_gd2启用（可选，但精确的列宽自动计算所需）
*   注意： PHP 5.6.29有一个阻止SQLite3缓存正常工作的错误。如果您需要SQLite3缓存，请使用较新的（或更旧版本）的PHP。

### 使用composer进行安装
~~~
     composer require tp5er/easyexcel dev-master
~~~
### 使用composer update进行安装
~~~
    "require": {
        "tp5er/easyexcel": "dev-master"
    },
~~~

## 使用easyExcel类进行导出导入使用方法
实例化对象->导出->下载到本地/保存到服务器上。
实例化对象->导入->获取参数之后自行处理。

## 引入类文件
~~~
    use \tp5er\easyExcel;
~~~

### 导出保存在服务器上
~~~
    //表的数据设置
    $arr=db('user')->field('id,name,sex')->limit(10)->select();
    $arr=Array
    (
        '0' => Array('id' => 1,'name' => 'tp5er','sex'=>'男'),
        '1' => Array('id' => 1,'name' => 'thinkphp','sex'=>'男')
    );
    表头数据
    $excelHeader=array_keys($list[0]);
    Array ( [0] => id [1] => name [2] => mobile)
    实例化
    $easyExcel=new easyExcel();
    $easyExcel->createSheet('Sheet1',$list,$fileheader)->createSheet('Sheet2',$list,$fileheader)->saveFile();
~~~

### 导出下载到本地

~~~
    $easyExcel->createSheet('Sheet1',$list,$fileheader)->createSheet('Sheet2',$list,$fileheader)->downFile();
~~~

### 导入使用手册

~~~
    获取有多少sheet
    $Sheet=$easyExcel->loadExcel($filepath)->getSheetNames();
    传入sheet获取对应数据
    $arr=$easyExcel->getSheetByName($Sheet[0])->toArray();
~~~



