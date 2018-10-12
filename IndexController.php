<?php

namespace App\Http\Controllers\Admin;
use Illuminate\Http\Request;
use Excel;
use PHPExcel;
use App\Http\Controllers\Controller;
use App\Helper\DateHelper;
include_once dirname(dirname(dirname(dirname(__DIR__)))).'/vendor/phpoffice/phpexcel/Classes/PHPExcel/IOFactory.php';
class IndexController extends Controller
{
    //
    public function index($id){
        echo storage_path('template');
        exit();
        return View('admin.index.index',['id' => $id]);
    }



    public function excel1(Request $request){
        Excel::load(config('dict.attendance_template_path'),function($reader){
            $reader = $reader->getSheet(0); //获取第一个个工作表
            $result = $reader->toArray(); // 获取所有的数据
            //获取当前的数组了  然后进行数据拼接  最后生成新的xls
            unset($result[0]);
            $attendance_info = config('test');
            $attendance_info = array_merge($result,$attendance_info);
            $new_fileName ='考勤情况'; //这里暂且不做编码转化
            $sheet_name = '考勤详情';  //这里暂且不做编码转化
            Excel::create($new_fileName,function($excel) use ($attendance_info,$sheet_name){
                $excel->sheet($sheet_name,function($sheet) use ($attendance_info){
                    $sheet->fromArray($attendance_info); //根据$attendance_info生成工作表
                    $sheet->setFontSize(30);  //设置字体大小
                });
            })
            ->export('xls');  //直接导出数据到浏览器 store是存储文件到服务器 保存到/storage/exports/
        });
    }


    public function excel(Request $request){
        $excel = new PHPExcel();
        $topNumber = 4; //表头部分
        $xlsName = '月份考勤情况'; //现在不需要进行编码转化
        $cellKey = [
            'A','B','C','D','E','F','G','H','I','J','K','L','M',
            'N','O','P','Q','R','S','T','U','V','W','X','Y','Z',
            'AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK',
            'AL','AM','AN','AO'
        ];

        //处理表头标题  这里是合并行的单元格  
        $excel->getActiveSheet()->mergeCells('A1:'.$cellKey[count($cellKey)-1].'1'); //合并单元格
        $excel->setActiveSheetIndex(0)->setCellValue('A1','员工考勤记录表'); //设置第一个单元格的内容
        $excel->getActiveSheet()->getRowDimension('1')->setRowHeight(30); //设置特定行的高度的为30
        $excel->getActiveSheet()->getStyle('A1')->getFont()->setSize(18);//设置第一个单元格(也就是第一行标题)字体为18
        $excel->getActiveSheet()->getStyle('A1')->getFont()->setBold(true); //设置一个单元格文字加粗
        $excel->getActiveSheet()->getStyle('A1')->getFont()->getColor()->setARGB(\PHPExcel_Style_Color::COLOR_BLACK);// 设置文字颜色
        $excel->getActiveSheet()->getStyle('A1')->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);//设置文字竖直居中
        $excel->getActiveSheet()->getStyle('A1')->getAlignment()->setVertical(\PHPExcel_Style_Alignment::VERTICAL_CENTER);//文字垂直居中


        $excel->getDefaultStyle()->getFont()->setSize(13);   // 设置默认字号
        $excel->getDefaultStyle()->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);//设置默认文字竖直居中
        $excel->getDefaultStyle()->getAlignment()->setVertical(\PHPExcel_Style_Alignment::VERTICAL_CENTER);//设置默认文字垂直居中

        //设置边框
        $styleThinBlackBorderOutline = array(
            'borders' => array(
                'allborders' => array( //设置全部边框
                    'style' => \PHPExcel_Style_Border::BORDER_THIN //粗的是thick
                ),

            ),
        );
        $excel->getActiveSheet()->getStyle('A1:AO3')->applyFromArray($styleThinBlackBorderOutline);


        $excel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(5);//所有单元格(行)默认高度为20
        $excel->getActiveSheet()->getDefaultColumnDimension()->setWidth(4); //设置单元格的默认的宽度为20



        $excel->getActiveSheet()->getRowDimension('3')->setRowHeight(30); //设置特定行的高度的为30
        $excel->getActiveSheet()->getColumnDimension('B')->setWidth(10);//设置特定列的宽度40

        //合并列单元格  这里需要注意
        $mergeColumn = $excel->getActiveSheet()->mergeCellsByColumnAndRow(0,2,0,3);
        $mergeColumn->setCellValueByColumnAndRow(0,2,'序号');
        $mergeColumn->getStyleByColumnAndRow(0,2,0,3)->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);//设置文字竖直居中
        $mergeColumn->getStyleByColumnAndRow(0,2,0,3)->getAlignment()->setVertical(\PHPExcel_Style_Alignment::VERTICAL_CENTER);//文字垂直居中

        //添加日期
        $excel->setActiveSheetIndex(0)->setCellValue('B2','日期');
        //添加具体日期
        for($i=1;$i<=31;$i++){
            $skip = 1;
            $total = $i+$skip;
            //添加日期
            $excel->setActiveSheetIndex(0)->setCellValue($cellKey[$total].'2',$i);
        }

        //添加星期
        $weeks = DateHelper::generateWeekArray(9);
        $total_num = count($weeks);
        for($i=1;$i<=$total_num;$i++){
            $skip = 1;
            $total = $i+$skip;
            $excel->setActiveSheetIndex(0)->setCellValue($cellKey[$total].'3',$weeks[$i-1]);
        }

        //合并单元格 填充考勤情况
        $excel->getActiveSheet()->mergeCells('AH2:AO2');
        $excel->setActiveSheetIndex(0)->setCellValue('AH2','月统计考勤情况'); //合并单元格


        //设置斜线
        $excel->getActiveSheet()->getStyle('B3')->getBorders()->setDiagonalDirection(\PHPExcel_Style_Borders::DIAGONAL_DOWN );
        $excel->getActiveSheet()->getStyle('B3')->getBorders()->getDiagonal()-> setBorderStyle(\PHPExcel_Style_Border::BORDER_THIN);


        $excel->setActiveSheetIndex(0)->setCellValue('B3','姓名       星期');
        $excel->getActiveSheet()->getStyle('B3')->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_LEFT);


        //添加考勤标注
        $excel->setActiveSheetIndex(0)->setCellValue('AH3','应出勤');
        $AH3 = $excel->getActiveSheet()->getStyle('AH3');
        $AH3->getAlignment()->setWrapText(true); //设置文字在一个单元格换行
        $AH3->getFont()->setSize(10);

        $excel->setActiveSheetIndex(0)->setCellValue('AI3','实际出勤');
        $AI3 = $excel->getActiveSheet()->getStyle('AI3');
        $AI3->getAlignment()->setWrapText(true); //设置文字在一个单元格换行
        $AI3->getFont()->setSize(10);

        $excel->setActiveSheetIndex(0)->setCellValue('AJ3','调休');
        $AJ3 =  $excel->getActiveSheet()->getStyle('AJ3');
        $AJ3->getAlignment()->setWrapText(true); //设置文字在一个单元格换行
        $AJ3->getFont()->setSize(10);

        $excel->setActiveSheetIndex(0)->setCellValue('AK3','事假');
        $AK3 = $excel->getActiveSheet()->getStyle('AK3');
        $AK3->getAlignment()->setWrapText(true); //设置文字在一个单元格换行
        $AK3->getFont()->setSize(10);

        $excel->setActiveSheetIndex(0)->setCellValue('AL3','加班');
        $AL3 = $excel->getActiveSheet()->getStyle('AL3');
        $AL3->getAlignment()->setWrapText(true); //设置文字在一个单元格换行
        $AL3->getFont()->setSize(10);

        $excel->setActiveSheetIndex(0)->setCellValue('AM3','出差');
        $AM3 = $excel->getActiveSheet()->getStyle('AM3');
        $AM3->getAlignment()->setWrapText(true); //设置文字在一个单元格换行
        $AM3->getFont()->setSize(10);

        $excel->setActiveSheetIndex(0)->setCellValue('AN3','迟到');
        $AN3 = $excel->getActiveSheet()->getStyle('AN3');
        $AN3->getAlignment()->setWrapText(true); //设置文字在一个单元格换行
        $AN3->getFont()->setSize(10);

        $excel->setActiveSheetIndex(0)->setCellValue('AO3','其他');
        $AO3 = $excel->getActiveSheet()->getStyle('AO3');
        $AO3->getAlignment()->setWrapText(true); //设置文字在一个单元格换行
        $AO3->getFont()->setSize(10);


         //导入考勤数据
         //  todo

        header('pragma:public');
        header('Content-type:application/vnd.ms-excel;charset=utf-8;name="'.$xlsName.'.xls"');
        header("Content-Disposition:attachment;filename=$xlsName.xls");//attachment新窗口打印inline本窗口打印
        $xlsWriter = \PHPExcel_IOFactory::createWriter($excel,'Excel2007');
        $xlsWriter->save('php://output');
    }

}
