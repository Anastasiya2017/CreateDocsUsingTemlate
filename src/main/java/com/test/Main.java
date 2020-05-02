package com.test;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
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
    static Set<String> setNameHeads;

    public static void main(String[] args) throws IOException {
        System.out.println("Program start!");
        addErrorInFile(new Date() + "\n", false);
        readXLSX();
    }

    private static void readXLSX() throws IOException {
        String path = "CheckFold/data.xlsx";
        FileInputStream excelFile = null;
        try {
            excelFile = new FileInputStream(new File(path));
        } catch (FileNotFoundException e) {
            addErrorInFile("File " + path + " not found", true);
            e.printStackTrace();
            System.exit(0);
        }
        XSSFWorkbook workbook = new XSSFWorkbook(excelFile);
        XSSFSheet datatypeSheet = workbook.getSheetAt(0);
        Iterator iterator = datatypeSheet.iterator();

        while (iterator.hasNext()) {
            Row currentRow = (Row) iterator.next();
            Iterator cellIterator = currentRow.iterator();
            List<String> row = new ArrayList<>();
            //номер текущей строки
            while (cellIterator.hasNext()) {
                Cell currentCell = (Cell) cellIterator.next();
                if (currentCell.getCellType() == CellType.STRING) {
                    row.add(currentCell.getStringCellValue());
                } else if (currentCell.getCellType() == CellType.NUMERIC) {
                    row.add(currentCell.getStringCellValue());
                }
                if (!cellIterator.hasNext() && currentRow.getRowNum() == 0) {
                    checkNameHeaderTableXLSX(row);
                }
                if (!cellIterator.hasNext() && currentRow.getRowNum() != 0) {
                    writeInDoc(row);
                }
            }
        }
    }

    private static void checkNameHeaderTableXLSX(List<String> nameHeadTable) {
        String[] listNameHeads = {"TaskNum", "ImgCount", "ItemGroup", "ItemName",
                "Color", "InteriorColor", "StartDate", "EndDate",
                "Size", "MainImage", "ResultImage"};
        setNameHeads = new HashSet<>(Arrays.asList(listNameHeads));
        listNameHeadsInMap = new HashMap<>();
        for (int i = 0; i < listNameHeads.length; i++) {
            listNameHeadsInMap.put(i, listNameHeads[i]);

        }
        if (listNameHeads.length != nameHeadTable.size()) {
            addErrorInFile("the number of columns in the .xlsx file is not equal " + listNameHeads.length, true);
            System.exit(0);
        }
        for (int i = 0; i < listNameHeads.length; i++) {
            if (!listNameHeads[i].equals(nameHeadTable.get(i))) {
                addErrorInFile("the name of the columns in the .xlsx file does not meet the requirements.\n " +
                        "\t The names and order of the columns should be as follows: {\"TaskNum\", \"ImgCount\"," +
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
            writer.write("- " + s + "\n");
            writer.close();
        } catch (IOException e) {
            System.out.println(e);
        }
    }

    private static void writeInDoc(List<String> row) {
        new File("CheckFold/result").mkdirs();
        String in = "CheckFold/template.docx";
        String out = "CheckFold/result/Служ_задание_Акт_" + row.get(0).replace('/', '_') + ".docx";
        try {
            InputStream is = new FileInputStream(in);
            OutputStream os = new FileOutputStream(out);
            byte[] buffer = new byte[1024];
            int length;
            while ((length = is.read(buffer)) > 0) {
                os.write(buffer, 0, length);
            }
            is.close();
            os.close();
        } catch (Exception e) {
            addErrorInFile("file not found: " + in, true);
            e.printStackTrace();
            System.exit(0);
        }
        try {
            System.out.print("-- ");
            XWPFDocument doc = new XWPFDocument(OPCPackage.open(out));
            for (XWPFParagraph p : doc.getParagraphs()) {
                List<XWPFRun> runs = p.getRuns();
                if (runs != null) {
                    for (XWPFRun r : runs) {
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
            String out2 = "CheckFold/result/" + new Date().getTime() + ".docx";
            FileOutputStream out3 = new FileOutputStream(out2);
            File file = new File(out2);
            doc.write(out3);
            doc.close();
            out3.close();
            file.delete();
        } catch (Exception e) {
            e.printStackTrace();
        }
        if (!setNameHeads.isEmpty()) {
            addErrorInFile(("not replace: " + setNameHeads.toString()), true);
            System.out.println("не были заменены: " + setNameHeads.toString());
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
        if (text != null && text.contains("&MainImage")) {
            setNameHeads.remove("MainImage");
            String img = "CheckFold/images/" + rowXLSX.get(9);
            try {
                InputStream pic = new FileInputStream(img);
                BufferedImage bi = ImageIO.read(new File(img));
                int width = 300;
                double height = (double) bi.getHeight() / ((double) bi.getWidth() / width);
                text = text.replace("&MainImage", "");
                paragrafDOCS.setText(text, 0);
                paragrafDOCS.addPicture(pic, XWPFDocument.PICTURE_TYPE_JPEG, img, Units.toEMU(width), Units.toEMU(height));
            } catch (Exception e) {
                System.out.println("файл " + img + " не найден ");
                addErrorInFile("file " + img + " not found ", true);
            }
        }

        if (text != null && text.contains("&Table")) {
            setNameHeads.remove("ResultImage");
            text = text.replace("&Table", "");
            paragrafDOCS.setText(text, 0);
            String imgs = rowXLSX.get(10).replace(" ", "");
            String[] getImg = imgs.split(",");
            int imgNum = getImg.length;
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
                    if (k < imgNum) {
                        try {
                            img = "CheckFold/images/" + getImg[k];
                            InputStream pic = new FileInputStream(img);
                            BufferedImage bi = ImageIO.read(new File(img));
                            double height = bi.getHeight() / ((double) bi.getWidth() / width);
                            paragrafDOCS.addPicture(pic, XWPFDocument.PICTURE_TYPE_JPEG, img, Units.toEMU(width), Units.toEMU(height));
                        } catch (Exception e) {
                            System.out.println("Не найдены файл " + img);
                            addErrorInFile("file " + img + " not found ", true);
                        }
                        k++;
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
                setNameHeads.remove(variable.substring(1));
                Set<Map.Entry<Integer, String>> entrySet = listNameHeadsInMap.entrySet();
                for (Map.Entry<Integer, String> pair : entrySet) {
                    if (variable.substring(1).equals(pair.getValue())) {
                        key = pair.getKey();
                        break;
                    }
                }
                text = text.replace(variable, rowXLSX.get(key));
                paragrafDOCS.setText(text, 0);
            }
        }
    }
}