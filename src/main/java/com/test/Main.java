package com.test;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.*;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Main {
    static String FILE_ERROR = "TestFold/error.txt";


    public static void main(String[] args) throws IOException {
        System.out.println("Program start!");
        readXLSX();
    }

 /*   private static String getCellName(Cell cell) {
        return CellReference.convertNumToColString(cell.getColumnIndex()) + (cell.getRowIndex() + 1);
    }*/

    private static void readXLSX() throws IOException {
        String path = "TestFold/data.xlsx";
        FileInputStream excelFile = null;
        try {
            excelFile = new FileInputStream(new File(path));
        } catch (FileNotFoundException e) {
            addErrorInFile("Файл " + path + " не найден");
            e.printStackTrace();
        }
        XSSFWorkbook workbook = new XSSFWorkbook(excelFile);
        XSSFSheet datatypeSheet = workbook.getSheetAt(0);
        Iterator iterator = datatypeSheet.iterator();

        while (iterator.hasNext()) {
            Row currentRow = (Row) iterator.next();
            Iterator cellIterator = currentRow.iterator();
            List<String> row = new ArrayList<>();
            //номер текущей строки
            System.out.println(currentRow.getRowNum());
            while (cellIterator.hasNext()) {
                Cell currentCell = (Cell) cellIterator.next();
                if (currentCell.getCellType() == CellType.STRING) {
                    row.add(currentCell.getStringCellValue());
                    System.out.print(currentCell.getStringCellValue() + "--");
                } else if (currentCell.getCellType() == CellType.NUMERIC) {
                    row.add(currentCell.getStringCellValue());
                    System.out.print(currentCell.getNumericCellValue() + "--");
                }
                if (!cellIterator.hasNext() && currentRow.getRowNum() != 0) {
                    System.out.println();
                    System.out.println("массив:");
                    System.out.println("тукущая строка " + currentRow.getRowNum());
                    System.out.println("Всего строк " + currentRow.getSheet().getLastRowNum());
                    writeInDoc(row);
                }
            }
            System.out.println();
        }
    }

    private static void addErrorInFile(String s) {
        try {
            FileWriter writer = new FileWriter(FILE_ERROR, true);
            BufferedWriter bufferWriter = new BufferedWriter(writer);
            bufferWriter.write(s + "\n");
            bufferWriter.close();
        }
        catch (IOException e) {
            System.out.println(e);
        }
    }

    private static void writeInDoc(List<String> row) throws IOException {
       /* for (String f : row) {
            System.out.println(f);
        }*/
//        for (int i = 1; i < row.size(); i++) {
//            Template template = new Template();
//            template.setImgCount(row.get(i));
//        }
        String in = "TestFold/template.docx";
        String out = "TestFold/Служ_задание_Акт_" + row.get(1) + ".docx";
        InputStream is = new FileInputStream(in);
        OutputStream os = new FileOutputStream(out);
        byte[] buffer = new byte[1024];
        int length;
        while ((length = is.read(buffer)) > 0) {
            os.write(buffer, 0, length);
        }
        is.close();
        os.close();
        try {
            XWPFDocument doc = new XWPFDocument(OPCPackage.open(out));
            for (XWPFParagraph p : doc.getParagraphs()) {
                List<XWPFRun> runs = p.getRuns();
                if (runs != null) {
                    for (XWPFRun r : runs) {
//                        System.out.println(runs);
                        replaceString(row, r);
                        //создание таблицы с изображениями:
                        String text = r.getText(0);
                        if (text != null && text.contains("&table")) {
                            text = text.replace("&table", "");
                            r.setText(text, 0);
//                            System.out.println("TABLE");
                            String img = "TestFold/images/1.jpg";
                            InputStream pic = new FileInputStream(img);
                            BufferedImage bi = ImageIO.read(new File(img));
                            Pattern pattern = Pattern.compile("\\d+");
                            Matcher matcher = pattern.matcher(row.get(1));
                            int imgNum = 1;
                            if (matcher.find()) {
                                String value = matcher.group();
                                imgNum = Integer.parseInt(value);
                                //кол-во картинок
                                System.out.println("кол-во картинок: " + imgNum);
                            }
                            int width = 100;
                            int height = bi.getHeight() / (bi.getWidth() / 100);
                            System.out.println(width + " : " + height);
                            r.addBreak();
                            try {
                                r.addPicture(pic, XWPFDocument.PICTURE_TYPE_JPEG, img, Units.toEMU(width), Units.toEMU(height));
                            } catch (Exception e) {

                            }
                            XWPFTable tableX = doc.createTable();
//                            int rowNum = imgNum / 6;
//                            if (imgNum % 5 != 0) {
//                                rowNum++;
//                            }
                            for (int i = 0; i < 3; i++) {
                                XWPFTableRow rowX = tableX.getRow(i);
                                for (int j = 0; j < 5; j++) {
                                    XWPFTableCell cellX = rowX.getCell(j);
                                    try {
                                        cellX.setText(i + "-" + j);
                                        r.addPicture(pic, XWPFDocument.PICTURE_TYPE_JPEG, img, Units.toEMU(width), Units.toEMU(height));
                                    } catch (Exception e) {
                                    }
                                    if (i == 0 && j != 4)
                                        rowX.createCell();
                                }
                                if (i != 2)
                                    tableX.createRow(); //создание строки в таблице.
                            }
                        }
                    }
                }
            }

            for (XWPFTable tbl : doc.getTables()) {
                for (XWPFTableRow xwpfTableRow : tbl.getRows()) {
                    for (XWPFTableCell cell : xwpfTableRow.getTableCells()) {
                        for (XWPFParagraph p : cell.getParagraphs()) {
                            for (XWPFRun r : p.getRuns()) {
                                replaceString(row, r);
                            }
                        }
                    }
                }
            }
            String out2 = "TestFold/Служ_задание" + new Date() + ".docx";
            FileOutputStream out3 = new FileOutputStream(out2);
            File file = new File(out2);
            doc.write(out3);
            doc.close();
            out3.close();
            if (file.delete()) {
                System.out.println("файл удален");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void replaceString(List<String> row, XWPFRun r) throws IOException {
        String text = r.getText(0);
        if (text != null) {
            if (text.contains("&imgCount")) {
                text = text.replace("&imgCount", row.get(1));
                r.setText(text, 0);
            }
            if (text.contains("&itemGroup")) {
                text = text.replace("&itemGroup", row.get(2));
                r.setText(text, 0);
            }
            if (text.contains("&itemName")) {
                text = text.replace("&itemName", row.get(3));
                r.setText(text, 0);
            }
            if (text.contains("&color")) {
                text = text.replace("&color", row.get(4));
                r.setText(text, 0);
            }
            if (text.contains("&interiorColor")) {
                text = text.replace("&interiorColor", row.get(5));
                r.setText(text, 0);
            }
            if (text.contains("&startDate")) {
                text = text.replace("&startDate", row.get(6));
                r.setText(text, 0);
            }
            if (text.contains("&endDate")) {
                text = text.replace("&endDate", row.get(7));
                r.setText(text, 0);
            }
            if (text.contains("&size")) {
                text = text.replace("&size", row.get(8));
                r.setText(text, 0);
            }
            if (text.contains("&imgs")) {
       /* XWPFTable tableX = doc.createTable();//Создаю таблицу. Сразу создаётся XWPFTableRow и
        //XWPFTableCell.
        for (int i = 0; i < 8; i++) {
            XWPFTableRow rowX = tableX.getRow(i); //Не создаю, а получаю уже существующую строку
            for (int j = 0; j < 8; j++) {
                XWPFTableCell cellX = rowX.getCell(j); //не создаю, а получаю уже имеющуюся строку
                cellX.setText(i + "-" + j); //вставляю текст
                if (i == 0 && j != 7)  //Создаю ячейки только в первой строке, и не создаю лишнюю ячейку в
                    //конце строки.
                    rowX.createCell();//создание ячейки в строке
            }
            if (i != 7) //не создаю лишнюю строку в конце таблицы
                tableX.createRow(); //создание строки в таблице.
        }*/
                /*text = text.replace("&imgs", "");
                r.setText(text, 0);
                String img = "TestFold/images/1.jpg";
                InputStream pic = new FileInputStream(img);
                BufferedImage bi = ImageIO.read(new File(img));

                Pattern pattern = Pattern.compile("\\d+");
                Matcher matcher = pattern.matcher(row.get(1));
                if (matcher.find()) {
                    String value = matcher.group();
                    int result = Integer.parseInt(value);
                    System.out.println(result);
                }
//              int width   = bi.getWidth()/ Integer.parseInt(row.get(1).split("\\D+")[1]);
//              int height   = bi.getHeight()/Integer.parseInt(row.get(1).split("\\D+")[1]);
                int width = 100;
                int height = bi.getHeight() / (bi.getWidth() / 100);
                System.out.println(width + " : " + height);
//                    byte [] picbytes = IOUtils.toByteArray(pic);
                r.addBreak();
                try {

                    r.addPicture(pic, XWPFDocument.PICTURE_TYPE_JPEG, img, Units.toEMU(width), Units.toEMU(height));
                } catch (Exception e) {

                }*/
//                    text = text.replace("&size", row.get(8));
//                    r.setText(text, 0);
            }
        }
    }

    private static void readDoc() {
        try {
            XWPFDocument docx = new XWPFDocument(new FileInputStream("/home/asya/Загрузки/Projects2020/ТЗ/Служ_задание_Акт_template.docx"));
            XWPFWordExtractor we = new XWPFWordExtractor(docx);
            System.out.println(we.getText());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}