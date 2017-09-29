
<?php
/**
 * Created by PhpStorm.
 * User: Administrator
 * Date: 2017/9/28 0028
 * Time: 15:47
 */

require_once './vendor/autoload.php';

class outPut
{
    public $data;
    public $phpExcel;
    public $activeSheet;
    public $column=[];
    public $key;
    public function __construct($data)
    {
        $this->data = $data;
        $this->phpExcel = new PHPExcel();
        $this->activeSheet = $this->phpExcel->getActiveSheet();
    }

    public function setSheetTitle($title)
    {
        $this->activeSheet->setTitle($title);
    }

    public function handle()
    {
        $this->writeField();
        $this->writeRow();
    }

    public function writeField()
    {
        //获取字段名
        $this->key = array_keys($this->data[0]);
        for($i=0;$i<count($this->key);$i++)
        {
            $column = PHPExcel_Cell::stringFromColumnIndex($i);
            array_push($this->column,$column);
            $position = $column.'1';

            //写入数据
            $this->activeSheet->setCellValue($position,$this->key[$i]);

            $this->fieldStyle($column);
        }
    }

    public function fieldStyle($column)
    {
        //设置宽度自动
        $this->activeSheet->getColumnDimension($column)->setAutoSize(true);

        $position = $column.'1';
        $this->activeSheet->getStyle($position)->getFont()->setSize(20);
        $this->activeSheet->getStyle($position)->getFont()->setBold(true);
        $this->activeSheet->getStyle($position)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $this->activeSheet->getStyle($position)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
    }

    public function writeRow()
    {
        $k = 1;
        foreach ($this->data as $item)
        {
            $k++;
            for($i=0;$i<count($this->key);$i++)
            {
                $position = $this->column[$i].$k;
                $value = $item[$this->key[$i]];
                $this->activeSheet->setCellValueExplicit($position,$value,\PHPExcel_Cell_DataType::TYPE_STRING);
                $this->rowStyle($position);
            }
        }
    }

    public function rowStyle($position)
    {
        $this->activeSheet->getStyle($position)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $this->activeSheet->getStyle($position)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
    }

    public function setName()
    {
        return uniqid();
    }
    public function outPut($name=null)
    {
        if(is_null($name)) $name = $this->setName();

        ob_end_clean();
        header("Content-Type:application/force-download");
        header("Content-Type:application/octet-stream");
        header("Content-Type:application/download");
        header('Content-Type:application/vnd.ms-excel');
        header('Content-Disposition:attachment;filename="'.$name.'.xlsx"');
        header('Cache-Control: no-cache, no-store, max-age=0, must-revalidate');
        $objWriter = \PHPExcel_IOFactory::createWriter($this->phpExcel, 'Excel2007');
        $objWriter->save('php://output');
    }
}

$data = [
    [
        'name' => '王晨',
        'sex' => '女',
        'age' => 18,
        'birth' =>  '1997-03-13',
        'phone' => '18948971224'
    ],
    [
        'name' => '李绯红',
        'sex' => '男',
        'age' => 21,
        'birth' =>  '1987-06-15',
        'phone' => '18891428924'
    ],
    [
        'name' => '王云',
        'sex' => '女',
        'age' => 19,
        'birth' =>  '1989-07-13',
        'phone' => '17936448924'
    ],
    [
        'name' => '郭瑞',
        'sex' => '女',
        'age' => 35,
        'birth' =>  '1977-05-13',
        'phone' => '18947896924'
    ]
];

$obj = new outPut($data);
$obj->setSheetTitle('员工信息');
$obj->handle();
$obj->outPut();