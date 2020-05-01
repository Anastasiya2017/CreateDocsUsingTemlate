package com.test;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.*;

public class Main {
    static String FILE_ERROR = "CheckFold/error.txt";
    static Map<Integer, String> listNameHeadsInMap = new HashMap<>();

    public static void main(String[] args) throws IOException {
        System.out.println("Program start!");
        addErrorInFile(new Date() + "\n", false);
//        FileWriter fileEr = new FileWriter(FILE_ERROR, true);
//        fileEr.write(new Date() + "\n");
//        fileEr.close();
        readXLSX();
    }

    private static void readXLSX() throws IOException {
        String path = "CheckFold/data.xlsx";
        FileInputStream excelFile = null;
        try {
            excelFile = new FileInputStream(new File(path));
        } catch (FileNotFoundException e) {
            addErrorInFile("- Файл " + path + " не найден", true);
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
                if (!cellIterator.hasNext() && currentRow.getRowNum() == 0) {
                    checkNameHeaderTableXLSX(row);
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

    private static void checkNameHeaderTableXLSX(List<String> nameHeadTable) {
        String[] listNameHeads = {"TaskNum", "ImgCount", "ItemGroup", "ItemName",
                "Color", "InteriorColor", "StartDate", "EndDate",
                "Size", "MainImage", "ResultImage"};
        listNameHeadsInMap = new HashMap<>();
        for (int i = 0; i < listNameHeads.length; i++) {
            listNameHeadsInMap.put(i, listNameHeads[i]);

        }
        System.out.println(listNameHeads.length + " = " + nameHeadTable.size());
        if (listNameHeads.length != nameHeadTable.size()) {
            addErrorInFile("количество столбцов в файле .xlsx не равно " + listNameHeads.length, true);
            System.exit(0);
        }
        for (int i = 0; i < listNameHeads.length; i++) {
            if (!listNameHeads[i].equals(nameHeadTable.get(i))) {
                addErrorInFile("название столбцов в файле .xlsx не соответсвует требованиям.\n " +
                        "\t Названия и порядок столбцов должны быть следующими: {\"TaskNum\", \"ImgCount\"," +
                        "\"ItemGroup\", \"ItemName\",\n\t " +
                        "\"Color\", \"InteriorColor\", \"StartDate\", \"EndDate\",\n\t " +
                        "\"Size\", \"MainImage\", \"ResultImage\"}", true);
                System.exit(0);
            }
        }
    }

    private static void addErrorInFile(String s, boolean status) {
        try {
            FileWriter writer = new FileWriter(FILE_ERROR, status);
            BufferedWriter bufferWriter = new BufferedWriter(writer);
            bufferWriter.write("- " + s + "\n");
            bufferWriter.close();
        } catch (IOException e) {
            System.out.println(e);
        }
    }

    private static void writeInDoc(List<String> row) throws IOException {
        String in = "CheckFold/template.docx";
        String out = "CheckFold/Служ_задание_Акт_" + row.get(1) + ".docx";
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
                        System.out.println(runs);
                        replaceString(row, r);
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
            String out2 = "CheckFold/Служ_задание" + new Date() + ".docx";
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

    private static void replaceString(List<String> rowXLSX, XWPFRun paragrafDOCS) {
        checkEmptyAndAddInDocs("&ImgCount", rowXLSX, paragrafDOCS);
        checkEmptyAndAddInDocs("&ItemGroup", rowXLSX, paragrafDOCS);
        checkEmptyAndAddInDocs("&ItemName", rowXLSX, paragrafDOCS);
        checkEmptyAndAddInDocs("&Color", rowXLSX, paragrafDOCS);
        checkEmptyAndAddInDocs("&InteriorColor", rowXLSX, paragrafDOCS);
        checkEmptyAndAddInDocs("&StartDate", rowXLSX, paragrafDOCS);
        checkEmptyAndAddInDocs("&EndDate", rowXLSX, paragrafDOCS);
        checkEmptyAndAddInDocs("&Size", rowXLSX, paragrafDOCS);
        //создание таблицы с изображениями:
        String text = paragrafDOCS.getText(0);
        if (text != null) {
            if (text.contains("&Table")) {
                text = text.replace("&Table", "");
                paragrafDOCS.setText(text, 0);
                String[] getImg = rowXLSX.get(10).split(",");
                int imgNum = getImg.length;
                System.out.println("число картинок: " + imgNum);
                String img = "";
                int width = 100;
                int rowNum = imgNum / 5;
                if (imgNum % 5 != 0) {
                    rowNum++;
                }
                rowNum++;
                int k = 0;
                for (int i = 0; i < rowNum; i++) {
                    for (int j = 0; j < 4; j++) {
                        if (k <= imgNum) {
                            System.out.println(k);
                            try {
                                img = "CheckFold/images/" + getImg[k];
                                InputStream pic = new FileInputStream(img);
                                BufferedImage bi = ImageIO.read(new File(img));
                                int height = bi.getHeight() / (bi.getWidth() / 100);
                                System.out.println(width + " : " + height);
                                paragrafDOCS.addPicture(pic, XWPFDocument.PICTURE_TYPE_JPEG, img, Units.toEMU(width), Units.toEMU(height));
                            } catch (Exception e) {
                            }
                            k++;
                        }
                    }
                }
            }
        }
    }

    private static void checkEmptyAndAddInDocs(String variable, List<String> rowXLSX, XWPFRun paragrafDOCS) {
        String text = paragrafDOCS.getText(0);
        int key = 0;
        if (text != null) {
            if (text.contains(variable)) {
                System.out.println(variable.substring(1));
                System.out.println(listNameHeadsInMap.entrySet().toString());
                Set<Map.Entry<Integer, String>> entrySet = listNameHeadsInMap.entrySet();
                for (Map.Entry<Integer, String> pair : entrySet) {
                    System.out.println("ddd " + pair.getValue() + " : " + pair.getKey());
                    if (variable.substring(1).equals(pair.getValue())) {
                        key = pair.getKey();// нашли наше значение и возвращаем  ключ
                        System.out.println("key: " + key);
                        break;
                    }
                }
                System.out.println(text);
                text = text.replace(variable, rowXLSX.get(key));
                System.out.println(text);
//                System.out.println("rowXLSX.get(key) " + variable + "  " + rowXLSX.get(key));
                paragrafDOCS.setText(text, 0);
            }
        }
    }
}