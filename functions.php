<?php
/**
 * Created by PhpStorm.
 * User: Administrator
 * Date: 2017/9/28 0028
 * Time: 15:01
 */

require_once './vendor/autoload.php';

$obj = new PHPExcel();

//创建者
$obj->getProperties()->setCreator('Summer Su');
//标题
$obj->getProperties()->setTitle('learn PHPExcel');
//题目
$obj->getProperties()->setSubject('second');

$obj->setActiveSheetIndex(0);
//设置sheet标题
$obj->getActiveSheet()->setTitle('test sheet');
//设置值
$obj->getActiveSheet()->setCellValue('A1','hello word');
$obj->getActiveSheet()->setCellValue('A2', 12);
$obj->getActiveSheet()->setCellValue('A3', true);

//设置单元格的宽度
$obj->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
$obj->getActiveSheet()->getColumnDimension('D')->setWidth(12);

// 所有单元格默认高度
$obj->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);
// 第一行的默认高度
$obj->getActiveSheet()->getRowDimension('1')->setRowHeight(30);

//设置字体（getStyle必须接受一个位置，不然会默认为A1）
$obj->getActiveSheet()->getStyle('B1')->getFont()->setName('Candara');
//设置B1的字体大小
$obj->getActiveSheet()->getStyle('B1')->getFont()->setSize(20);
//设置B1字体加粗
$obj->getActiveSheet()->getStyle('B1')->getFont()->setBold(true);
//设置B1字体颜色
$obj->getActiveSheet()->getStyle('B1')->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_BLUE);

//垂直居中
$obj->getActiveSheet()->getStyle('A18')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
//水平居中
$obj->getActiveSheet()->getStyle('A18')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

//设置单元格上边的大小
$obj->getActiveSheet()->getStyle('A18')->getBorders()->getTop()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
//设置单元格下边的大小
$obj->getActiveSheet()->getStyle('A18')->getBorders()->getBottom()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);

//设置填充颜色
$obj->getActiveSheet()->getStyle('A1')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$obj->getActiveSheet()->getStyle('A1')->getFill()->getStartColor()->setARGB('FF808080');