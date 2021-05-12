package com.ruoyi.bigdataAnalysis.until;

import com.ruoyi.bigdataAnalysis.domain.*;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.*;

public class ExcelUtil {


    /**
     * 读取excel数据
     * 文件路径
     *
     * @param file
     */
    public List<ImportDo> readExcel(File file,String filename) {

        Workbook wb = null;
        List<ImportDo> importDos = new ArrayList<>();
        ImportDo result = new ImportDo();

        try {
            wb = WorkbookFactory.create(file);
            for (int a =0; a<wb.getNumberOfSheets();a++){
                if (wb.getSheetAt(a).getPhysicalNumberOfRows()<=1){
                    continue;
                }
                result = readExcel(wb.getSheetAt(a));
                if (result!=null){
                    result.setOldtitlename(filename);
                    importDos.add(result);
                }
            }

        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return importDos;
    }



    /**
     * MultipartFile 转 File
     *
     * @param file
     * @throws Exception
     */
    public static File multipartFileToFile(MultipartFile file) throws Exception {
        // 获取文件名
        String fileName = file.getOriginalFilename();
        // 获取文件后缀
        String prefix=fileName.substring(fileName.lastIndexOf("."));
        File toFile  = File.createTempFile(UUID.randomUUID().toString(), prefix);
        // MultipartFile to File
        file.transferTo(toFile);
        return toFile;
    }





    /**
     * 读取excel文件
     *
     * @param sheet
     */
    private ImportDo readExcel( Sheet sheet) {

        Row row = null;
        ArrayList<Map<String, String>> result = new ArrayList<Map<String, String>>();
        CellDo cellHead = new CellDo();
        //一、获取名称。
        String titlename = null;
        String unit = null;
        row = sheet.getRow(0);
        if (row==null){
            return null;
        }
//        表头
        RowHead rowHead = new RowHead();
        CellRangeAddress range = isMergedRegion(sheet, 0, 0);
        //处理名称
        String name = sheet.getRow(0).getCell(0).getStringCellValue();
        String truename = "";
        String trueyearname = "";
        //处理标题中的 续 年
        if (name.contains("续")) {
            String str = name.substring(name.indexOf("续"), name.indexOf(" "));
            truename = name.replace(str, "");
            //标题中 存在续 存在年
            if (truename.contains("年")) {
                String stra = truename.substring(truename.indexOf("年")-4, truename.indexOf("年") + 1);
                trueyearname = stra.replace(" ", "");
                truename = truename.replace(trueyearname, "");
            } else {
                System.out.println("标题中只存在续，不存在年");
            }
        } else {
            //标题中存在年
            if (name.contains("年")) {
                String str = name.substring(name.indexOf("年")-4, name.indexOf("年") + 1);
                trueyearname = str.replace(" ", "");
                truename = name.replace(trueyearname, "");
            } else {
                truename = name;
            }
        }
        //最后处理空格
        titlename = truename.replace(" ", "");
        if (range!=null){
            rowHead.setCellsum(range.getLastColumn() + 1);
        }else {
            rowHead.setCellsum(row.getPhysicalNumberOfCells());
        }

        //二、单位
        row = sheet.getRow(1);
        String Tunit = "";
        for (Cell c : row) {
            if (c.getRichStringCellValue().toString() != "") {
                Tunit = c.getRichStringCellValue().getString().replace(" ", "");
                break;
            }
        }
        if (Tunit.contains("单位：")){
            unit = Tunit.replace("单位：","");
        }else if (Tunit.contains("单位:")){
            unit = Tunit.replace("单位:","");
        }else {
            unit = "";
        }
//        System.out.println(titlename + unit);
//          三、名称 操作表头
        row = sheet.getRow(2);
        Cell c = row.getCell(0);
        CellRangeAddress range1 = isMergedRegion(sheet, 2, c.getColumnIndex());

        cellHead.setCellname(getMergedRegionValue(sheet, row.getRowNum(), c.getColumnIndex()));
        if (range1 != null) {
//                是合并列
            cellHead.setMerged(true);
            cellHead.setLastColumn(range1.getLastColumn());
            cellHead.setLastRow(range1.getLastRow());
        } else {
            cellHead.setMerged(false);
        }
//            第三行
//            处理表头
        rowHead.setCellDo1(getcelldos(sheet.getRow(2)));
        if (cellHead.isMerged()) {
            if (cellHead.getLastRow() >= 3) {
                rowHead.setCellDo2(getcelldos(sheet.getRow(3)));
            }
            if (cellHead.getLastRow() >= 4) {
                rowHead.setCellDo3(getcelldos(sheet.getRow(4)));
            }

        }
        List<Target> targets = new ArrayList<>();
        for (int i = 0; i < rowHead.getCellsum(); i++) {
            Target target = new Target();
            if (i == 0) {
                target.setIstarget(false);
                targets.add(target);
                continue;
            }

            if (rowHead.getCellDo3() != null) {
                String name1 = ((rowHead.getCellDo1().get(i).getCellname()).replace(" ", "")).replaceAll("\r|\n","");
                String name2 = ((rowHead.getCellDo2().get(i).getCellname()).replace(" ", "")).replaceAll("\r|\n","");
                String name3 = ((rowHead.getCellDo3().get(i).getCellname()).replace(" ", "")).replaceAll("\r|\n","");
//                    判断是否为年
                if (name3.contains("年")) {
                    target.setIstarget(true);
                   //是取指标
                    if (name1.equals(name2)){
                        target.setName(name2);
                    } else {
                        //不等相加取指标
                        target.setName(name1+name2);
                    }
                    //取年
                    target.setYaername(name3);
                } else {
                    //最后一行不存在年 标题中存在年
                    if (trueyearname != "") {
                        target.setIstarget(true);
                        //是取指标
                        if (name1.equals(name2)){
                            if (name2.equals(name3)) {
                                target.setName(name1);
                            }else{
                                //不等相加取指标
                                target.setName(name1+name3);
                            }
                        } else {
                            if (name2.equals(name3)) {
                                //不等相加取指标
                                target.setName(name1 + name2);
                            }else{
                                //不等相加取指标
                                target.setName(name1+name2+name3);
                            }
                        }
                        //取年
                        target.setYaername(trueyearname);
                    }
                }
            }
            else if (rowHead.getCellDo2() != null) {
                String name1 = ((rowHead.getCellDo1().get(i).getCellname()).replace(" ", "")).replaceAll("\r|\n","");
                String name2 = ((rowHead.getCellDo2().get(i).getCellname()).replace(" ", "")).replaceAll("\r|\n","");
                //最后一行不存在年 标题中存在年
                    if (trueyearname != "") {
                        target.setIstarget(true);
                        //取指标
                        //是取指标
                        if (name2.equals(name1)){
                            target.setName(name2);
                        } else {
                            //不等相加取指标
                            target.setName(name1+name2);
                        }
                        //取年
                        target.setYaername(trueyearname);
                    } else {
                        //                    判断是否为年
                        if (name2.contains("年")) {
                            target.setIstarget(true);
                            //取指标
                            target.setName(name1);
                            //取年
                            target.setYaername(name2);
                        }
                    }
            }
            else {
                //最后一行不存在年 标题中存在年
                if (trueyearname != "") {
                    target.setIstarget(true);
                    String name1 = ((rowHead.getCellDo1().get(i).getCellname()).replace(" ", "")).replaceAll("\r|\n","");
                    target.setName(name1);
                    //取年
                    target.setYaername(trueyearname);
                }
            }

            targets.add(target);
        }
        //target   纬度
//        System.out.println(targets);
//        处理数据
        int start =3;
        if (range1!=null){
            start =   range1.getLastRow();
        }

        List<RegionDo> regionDos = new ArrayList<>();
        for (int i = start; i < sheet.getLastRowNum()+1; i++) {
            RegionDo regionDo = new RegionDo();
            row = sheet.getRow(i);

            if (row==null) {
                continue;
            }
            List<String> datas = new ArrayList<>();
            for (Cell cello : row) {
                String returnStr = getCellValue(cello);//获取值
                if ("".equals(returnStr) && cello.getColumnIndex() == 0) {
                    break;
                }
                if (cello.getColumnIndex() == 0) {
                    String regname = returnStr.replace(" ", "");
                    String str = "";
                    String trueregname = "";
                    //除去空格  转换括号
                    if (regname.contains("（")){
                        str = regname.replace("（", "(");
                    }else {
                        str = regname;
                    }

                    if (str.contains("）")){
                        trueregname = str.replace("）", ")");
                    }else {
                        trueregname = str;
                    }

                    regionDo.setRegname(trueregname);
                } else {
                    datas.add(returnStr);
                }
            }
            regionDo.setValues(datas);
            regionDos.add(regionDo);
        }
        ImportDo importDo = new ImportDo();
        importDo.setTitlename(titlename);
        importDo.setUnit(unit);
        importDo.setTargets(targets);
        importDo.setRegionDos(regionDos);
        if (trueyearname!=""){
            importDo.setTrueyearname(trueyearname);
        }else {
            importDo.setTrueyearname(targets.get(targets.size()-1).getYaername());
        }
        return importDo;
    }



    /**
     * 获取 行数据
     *
     * @param row
     * @return
     */
    private List<CellDo> getcelldos(Row row) {
        List<CellDo> cellDos1 = new ArrayList<>();
        for (Cell cell : row) {
            CellDo cellone = new CellDo();
            CellRangeAddress rangeone = isMergedRegion(row.getSheet(), row.getRowNum(), cell.getColumnIndex());
            cellone.setCellname(getMergedRegionValue(row.getSheet(), row.getRowNum(), cell.getColumnIndex()));
            if (rangeone != null) {
//                是合并列
                cellone.setMerged(true);
                cellone.setLastColumn(rangeone.getLastColumn());
                cellone.setLastRow(rangeone.getLastRow());
            } else {
                cellone.setMerged(false);
            }
            cellDos1.add(cellone);

        }
        return cellDos1;
    }

    /**
     * 获取合并单元格的值
     *
     * @param sheet
     * @param row
     * @param column
     * @return
     */
    public String getMergedRegionValue(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();
            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    Row fRow = sheet.getRow(firstRow);
                    Cell fCell = fRow.getCell(firstColumn);
                    return getCellValue(fCell);
                }
            }
        }
        return getCellValue(sheet.getRow(row).getCell(column));
    }

    /**
     * 判断合并了行
     *
     * @param sheet
     * @param row
     * @param column
     * @return
     */
    private boolean isMergedRow(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (row == firstRow && row == lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * 判断指定的单元格是否是合并单元格
     *
     * @param sheet
     * @param row    行下标
     * @param column 列下标
     * @return
     */
    private CellRangeAddress isMergedRegion(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        CellRangeAddress rangeone = null;
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (row == firstRow) {
                if (column == firstColumn) {
                    rangeone = range;
                } else {

                }

            }
        }
        return rangeone;
    }
    /**
     * 获取单元格的值
     *
     * @param cell
     * @return
     */
    public String getCellValue(Cell cell) {

        if (cell == null) return "";

        if (cell.getCellType() == Cell.CELL_TYPE_STRING) {

            return cell.getStringCellValue();

        } else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {

            return String.valueOf(cell.getBooleanCellValue());

        } else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {

            FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateFormulaCell(cell);
            CellValue cellValue = evaluator.evaluate(cell);
            Double celldata = cellValue.getNumberValue();

            return celldata.toString();

        } else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {

            return String.valueOf(cell.getNumericCellValue());

        }
        return "";
    }

    /**
     * 从excel读取内容
     */
    public static void readContent(String fileName) {
        boolean isE2007 = false;    //判断是否是excel2007格式
        if (fileName.endsWith("xlsx"))
            isE2007 = true;
        try {
            InputStream input = new FileInputStream(fileName);  //建立输入流
            Workbook wb = null;
            //根据文件格式(2003或者2007)来初始化
            if (isE2007)
                wb = new XSSFWorkbook(input);
            else
                wb = new HSSFWorkbook(input);
            Sheet sheet = wb.getSheetAt(0);     //获得第一个表单
            Iterator<Row> rows = sheet.rowIterator(); //获得第一个表单的迭代器
            while (rows.hasNext()) {
                Row row = rows.next();  //获得行数据
//                System.out.println("Row #" + row.getRowNum());  //获得行号从0开始
                Iterator<Cell> cells = row.cellIterator();    //获得第一行的迭代器
                while (cells.hasNext()) {
                    Cell cell = cells.next();
//                    System.out.println("Cell #" + cell.getColumnIndex());
                    switch (cell.getCellType()) {   //根据cell中的类型来输出数据
                        case HSSFCell.CELL_TYPE_NUMERIC:
                            System.out.println(cell.getNumericCellValue());
                            break;
                        case HSSFCell.CELL_TYPE_STRING:
                            System.out.println(cell.getStringCellValue());
                            break;
                        case HSSFCell.CELL_TYPE_BOOLEAN:
                            System.out.println(cell.getBooleanCellValue());
                            break;
                        case HSSFCell.CELL_TYPE_FORMULA:
                            System.out.println(cell.getCellFormula());
                            break;
                        default:
                            System.out.println("unsuported sell type=======" + cell.getCellType());
                            break;
                    }
                }
            }
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }



}

