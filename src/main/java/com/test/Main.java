package com.test;

import org.apache.poi.openxml4j.opc.OPCPackage;
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
    static String FILE_ERROR = "ChekFold/error.txt";
    
    public static void main(String[] args) throws IOException {
        System.out.println("Program start!");
        readXLSX();

    }

    private static String getCellName(Cell cell) {
        return CellReference.convertNumToColString(cell.getColumnIndex()) + (cell.getRowIndex() + 1);
    }

    private static void readXLSX() throws IOException {
        String path = "ChekFold/data.xlsx";
        FileInputStream excelFile = null;
        try {
            excelFile = new FileInputStream(new File(path));
        } catch (FileNotFoundException e) {
            addErrorInFile("- Файл " + path + " не найден");
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
                System.out.println( "first: " + getCellName(currentCell));
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
                    System.out.println("текущая строка " + currentRow.getRowNum());
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
        } catch (IOException e) {
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
        String in = "ChekFold/template.docx";
        String out = "ChekFold/Служ_задание_Акт_" + row.get(1) + ".docx";
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
                        replaceString(doc, row, r);
                    }
                }
            }

            for (XWPFTable tbl : doc.getTables()) {
                for (XWPFTableRow xwpfTableRow : tbl.getRows()) {
                    for (XWPFTableCell cell : xwpfTableRow.getTableCells()) {
                        for (XWPFParagraph p : cell.getParagraphs()) {
                            for (XWPFRun r : p.getRuns()) {
                                replaceString(doc, row, r);
                            }
                        }
                    }
                }
            }
            String out2 = "ChekFold/Служ_задание" + new Date() + ".docx";
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

    private static void replaceString(XWPFDocument doc, List<String> row, XWPFRun r) throws IOException {
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
            //создание таблицы с изображениями:
            if (text.contains("&table")) {
                text = text.replace("&table", "");
                r.setText(text, 0);
                Pattern pattern = Pattern.compile("\\d+");
                Matcher matcher = pattern.matcher(row.get(1));
                int imgNum = 0;
                if (matcher.find()) {
                    String value = matcher.group();
                    imgNum = Integer.parseInt(value);
                    //кол-во картинок
                    System.out.println("кол-во картинок: " + imgNum);
                }
                String[] getImg = row.get(10).split(",");
                imgNum = getImg.length;
                System.out.println("число картинок: " + imgNum);
//                String img = "ChekFold/images/1.jpg";
//                String img = "ChekFold/images/" + getImg[0];
                String img = "";
//                InputStream pic = new FileInputStream(img);
//                BufferedImage bi = ImageIO.read(new File(img));
                int width = 100;
               /* int height = bi.getHeight() / (bi.getWidth() / 100);
                System.out.println(width + " : " + height);*/
//                r.addBreak();
//                XWPFTable tableX = doc.createTable();
                int rowNum = imgNum / 5;
                if (imgNum % 5 != 0) {
                    rowNum++;
                }
                rowNum++;
                int k = 0;
                for (int i = 0; i < rowNum; i++) {
//                    XWPFTableRow rowX = tableX.getRow(i);
                    for (int j = 0; j < 4; j++) {
                        if (k <= imgNum) {
                        System.out.println(k);
//                        XWPFTableCell cellX = rowX.getCell(j);
                        try {
                            img = "ChekFold/images/" + getImg[k];
                            InputStream pic = new FileInputStream(img);
                            BufferedImage bi = ImageIO.read(new File(img));
                            int height = bi.getHeight() / (bi.getWidth() / 100);
                            System.out.println(width + " : " + height);
//                            cellX.setText(i + "-" + j);
                            r.addPicture(pic, XWPFDocument.PICTURE_TYPE_JPEG, img, Units.toEMU(width), Units.toEMU(height));
                        } catch (Exception e) {
                        }
                        if (i == 0 && j != 3)
//                            rowX.createCell();
                        k++;
                        }
                    }
                    if (i != rowNum-1){
//                        tableX.createRow(); //создание строки в таблице.
                    }
                }
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