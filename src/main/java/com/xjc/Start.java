package com.xjc;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;

import java.io.*;
import java.util.Scanner;

/* *
  新建并获取工作薄：Workbook workbook = Workbook.getWorkbook(inputStream);
  读取工作表：workbook.getSheet(int index);//index从0开始，0对应Sheet1
  获取单元格：sheet.getCell(int columnIndex, int rowIndex);
  读取单元格内容：cell.getContents();
* */
public class Start {
    public static void readExcel(File file, int rowCount, int sheetIndex) throws Exception {
        InputStream inputStream = new FileInputStream(file.getAbsoluteFile());
        Workbook workbook = Workbook.getWorkbook(inputStream);
        Sheet sheet = workbook.getSheet(sheetIndex);
        int rows = sheet.getRows();//如果空单元格也加了边线，那么也算进行列
        int columns = sheet.getColumns();
        String strStart = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?>\n" +
                "<Questions version=\"2.0\">\n" +
                "    <question type=\"1\" duration=\"60\">\n" +
                "        <title>\n" +
                "            <![CDATA[";
        String strEnd = "]]>\n" +
                "</item>\n" +
                "        <image />\n" +
                "    </question>\n" +
                "</Questions>\n";
        int lie3 = 2;
        int lie7 = 6;
        int lie8, lie9, lie10, lie11;
        lie8 = 7;
        lie9 = 8;
        lie10 = 9;
        lie11 = 10;
        String allText = "";
        String daAnA = "0";
        String daAnB = "0";
        String daAnC = "0";
        String daAnD = "0";
        /*创建一个文件夹，用于存放最终生成的文件*/
        File folder = new File("D:\\题目导出文件夹");
        /*i的最大值就是表中的行数，因为第22行代码解释的原因，所以这里需要手动输入*/
        for (int i = 1; i < rowCount; i++) {
            Cell cell = sheet.getCell(lie3, i);//获取第三列第二行的单元格
            String tiMu = cell.getContents();//获取单元格中的内容
            cell = sheet.getCell(lie7, i);
            String daAn = cell.getContents();//这里取到的是abcd字符串，还要判断
            /*daAn字符串中包含的答案用1表示，没有用0表示，1代表要勾选，0代表不勾选*/
            if (daAn.contains("A")) {
                daAnA = "1";
            } else {
                daAnA = "0";
            }
            if (daAn.contains("B")) {
                daAnB = "1";
            } else {
                daAnB = "0";
            }
            if (daAn.contains("C")) {
                daAnC = "1";
            } else {
                daAnC = "0";
            }
            if (daAn.contains("D")) {
                daAnD = "1";
            } else {
                daAnD = "0";
            }
            cell = sheet.getCell(lie8, i);
            String XA = cell.getContents();//选项A
            cell = sheet.getCell(lie9, i);
            String XB = cell.getContents();
            cell = sheet.getCell(lie10, i);
            String XC = cell.getContents();
            cell = sheet.getCell(lie11, i);
            String XD = cell.getContents();

            /*把所有字符串拼接起来*/
            allText = strStart + tiMu + "]]>\n" +
                    "</title>\n" +
                    "        <item checked=\"" + daAnA + "\">\n" +
                    "            <![CDATA[" + XA + "]]>\n" +
                    "</item>\n" +
                    "        <item checked=\"" + daAnB + "\">\n" +
                    "            <![CDATA[" + XB + "]]>\n" +
                    "</item>\n" +
                    "        <item checked=\"" + daAnC + "\">\n" +
                    "            <![CDATA[" + XC + "]]>\n" +
                    "</item>\n" +
                    "        <item checked=\"" + daAnD + "\">\n" +
                    "            <![CDATA[" + XD + strEnd;

            /*检查有没有这个文件夹，没有创建一个*/
            if (!folder.exists()) {
                folder.mkdir();
            }
            /*得到Excel文件的名字*/
            String fileName = file.getName();

            /*创建一个.survey后缀的文件，这是极域要求的格式*/
            BufferedWriter bw = new BufferedWriter(new FileWriter(folder + "\\" + fileName + "_" + i + ".survey"));
            /*把最终的文本写入创建好的文件*/
            bw.write(allText);
            /*关闭资源*/
            bw.close();
            /*置空文本，为下一次写入做准备*/
            allText = "";
        }
        System.out.println("文件存放到了：" + folder + "[如果你先修改存放位置，请修改41行代码中的路径]");
    }
    /*
        第三列第二行开始：题目
        第七列第二行开始：答案
        第八，九，十，十一列第二行开始：选项
    */

    public static void main(String[] args) {
        Scanner input = new Scanner(System.in);
        System.out.println("请输入要读取的Excel文件路径：");
        String path = input.nextLine();
        System.out.println("请输入要读取的Excel表格有多少行：");
        int rowCount = input.nextInt();
        System.out.println("请输入要读取的Excel表格在Excel工作蒲的第几个工作表(注意下标从0开始）：");
        int sheetIndex = input.nextInt();
        File file = new File(path);
        try {
            readExcel(file, rowCount, sheetIndex);//开始
            System.out.println("======读取完毕======");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
