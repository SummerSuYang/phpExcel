<?php
/**
 * Created by PhpStorm.
 * User: Administrator
 * Date: 2017/9/29 0029
 * Time: 9:51
 */

include_once './vendor/autoload.php';
include_once './toolbox.php';
set_time_limit(0);
class input
{
    public $phpExcel;
    public $activeSheet;
    public $reader;
    public $uploadFileName = 'myfile';
    public $uploadFilePath;
    public $acceptExtension=['xls','xlsx'];
    public $uploadFileExtension;
    public $dbserver;
    public function __construct()
    {
        $this->checkFile();
        $this->getFileInfo();
        $reader = PHPExcel_IOFactory::createReaderForFile($this->uploadFilePath);
        $this->phpExcel = $reader->load($this->uploadFilePath);
        $this->activeSheet = $this->phpExcel->getActiveSheet();
        $this->dbserver = dbServer::getInstance();
    }

    public function handle()
    {
        $data = $this->getData();

        $result = $this->insertData($data);

        $this->jsonReturn(1,"$result items hava been inserted.");
    }
    public function getData()
    {
        $row = $this->activeSheet->getHighestRow();
        $column =  PHPExcel_Cell::columnIndexFromString($this->activeSheet->getHighestDataColumn());

        $data = [];
        for($i=2;$i<=$row;$i++)
        {
            for($j=0;$j<$column;$j++)
            {
                $position = PHPExcel_Cell::stringFromColumnIndex($j).$i;
                $value = $this->activeSheet->getCell($position)->getValue();
                $data[$i][] = $this->filtValue($value);
            }
        }

        return $data;
    }

    public function filtValue($value)
    {
        if(is_null($value)) return "";

        return htmlspecialchars($value);
    }
    public function checkFile()
    {
        $file = $_FILES[$this->uploadFileName];

        if(is_null($file))
            $this->jsonReturn(0,'need a upload excel file');

        if(count($file) != count($file,1))
            $this->jsonReturn(0,'please upload one file');
    }

    public function getFileInfo()
    {
        $file = $_FILES[$this->uploadFileName];

        $this->uploadFilePath = $file['tmp_name'];

        $this->uploadFileExtension = $this->getFileExtension($file['name']);

        if( ! in_array($this->uploadFileExtension,$this->acceptExtension))
            $this->jsonReturn(0,'please upload the file with the proper extension');
    }

    public function insertData($data)
    {
        $value = 'values';
        foreach ($data as $item)
        {
            $attribute = '(';
            for($i=0;$i<count($item);$i++)
            {
                $attribute.="'$item[$i]',";
            }
            $attribute = rtrim($attribute,',');
            $attribute.=')';

            $value.="$attribute,";
        }

        $value = rtrim($value,',');

        try
        {
            $sql = "insert into excel(fullname,gender,region,phone,email,industry_id,position,company,work_age,education) $value";
            $result = $this->dbserver->exec($sql);
        }
        catch (Exception $e)
        {
            $this->jsonReturn(0,$e->getMessage());
        }

        return $result;
    }

    public function getFileExtension($str)
    {
        return substr($str,strrpos($str,'.')+1);
    }
    public function jsonReturn($code,$msg,$data='')
    {
        $data = [
            'code' => $code,
            'msg' => $msg,
            'data' => $data
        ];

        header('Content-type: application/json');
        echo json_encode($data);
        exit;
    }

    public function arrayReturn($code,$msg,$data='')
    {
        $data = [
            'code' => $code,
            'msg' => $msg,
            'data' => $data
        ];

        return $data;
    }
}


$obj = new input();
$obj->handle();