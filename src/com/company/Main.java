package com.company;

import org.apache.poi.ooxml.POIXMLException;
import org.apache.poi.openxml4j.opc.*;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.formula.functions.Log;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.net.URL;

import static org.apache.poi.ooxml.POIXMLDocument.DOCUMENT_CREATOR;

public class Main {


    public static void main(String[] args) throws IOException {
        ZipSecureFile.setMinInflateRatio(0);

        String[] url = new String[4];
        String[] waytosavesource = new String[4];
        String[] waytosave = new String[4];
        Integer items = 4;
        waytosave[0] = "C:/Users/lul/excelfilesguu/bak1.xlsx";
        waytosave[1] = "C:/Users/lul/excelfilesguu/bak2.xlsx";
        waytosave[2] = "C:/Users/lul/excelfilesguu/bak3.xlsx";
        waytosave[3] = "C:/Users/lul/excelfilesguu/bak4.xlsx";
        waytosavesource[0] = "C:/excelfiles/bak1.xlsx";
        waytosavesource[1] = "C:/excelfiles/bak2.xlsx";
        waytosavesource[2] = "C:/excelfiles/bak3.xlsx";
        waytosavesource[3] = "C:/excelfiles/bak4.xlsx";
        url[0] = "https://guu.ru/wp-content/uploads/1-%D0%BA%D1%83%D1%80%D1%81-%D0%B1%D0%B0%D0%BA%D0%B0%D0%BB%D0%B0%D0%B2%D1%80%D0%B8%D0%B0%D1%82-%D0%9E%D0%A4%D0%9E-36.xlsx";
        url[1] = "https://my.guu.ru/student/messages/get-file?file=b94bacc0de5b0a160eba2d3fc4c106931b1e7247&msg=hJfnCZxiPxnxQAwoor1bk8FUFjDBRBn8";
        url[2] = "https://guu.ru/wp-content/uploads/3-%D0%BA%D1%83%D1%80%D1%81-%D0%B1%D0%B0%D0%BA%D0%B0%D0%BB%D0%B0%D0%B2%D1%80%D0%B8%D0%B0%D1%82-%D0%9E%D0%A4%D0%9E-39.xlsx";
        url[3] = "https://guu.ru/wp-content/uploads/4-%D0%BA%D1%83%D1%80%D1%81-%D0%B1%D0%B0%D0%BA%D0%B0%D0%BB%D0%B0%D0%B2%D1%80%D0%B8%D0%B0%D1%82-%D0%9E%D0%A4%D0%9E-31.xlsx";
        for (int i = 0; i < items; i++) {
            try {
                downloadUsingStream(url[i], waytosavesource[i]);

            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        for (int i = 0; i < items; i++) {
            convertexcel(waytosavesource[i], waytosave[i]);
        }
        searchbook(waytosave);
    }


    private static void searchbook(String[] books) throws IOException {

        String string = "";

        for (int b = 0; b < books.length; b++) {
            try {


                FileInputStream fis = new FileInputStream(books[b]);
                XSSFWorkbook wbsource = new XSSFWorkbook(fis);
                for (int s = 0; s < wbsource.getNumberOfSheets(); s++) {

                    Sheet sheet = wbsource.getSheetAt(s);
                    for (int c = 4; c < 60; c++) {
                    for (int r = 8; r < 60; r++) {

                            ;

                            try {
                                if (sheet.getRow(r).getCell(c).toString() != null) {
                                    if(sheet.getRow(r).getCell(c).toString().length() > 2 && !sheet.getRow(r).getCell(c).toString().contains("Директор") && !string.contains(sheet.getRow(r).getCell(c).toString()
                                            + "Y" + sheet.getRow(r).getCell(3).toString()
                                            + "Y" + sheet.getRow(r).getCell(2).toString()
                                            + "Y" + sheet.getRow(r).getCell(1).toString()
                                            + "Y" + "Курс-" + (b+1)))
                                    string = string + "\nX\n" + sheet.getRow(r).getCell(c).toString()
                                            + "Y" + sheet.getRow(r).getCell(3).toString()
                                            + "Y" + sheet.getRow(r).getCell(2).toString()
                                            + "Y" + sheet.getRow(r).getCell(1).toString()
                                            + "Y" + "Курс-" + (b+1) + "Y";
                                }

                            } catch (Exception e) {

                            }


                        }
                    }

                }

            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }


            FileOutputStream fos = new FileOutputStream("C:/Users/lul/excelfilesguu/search.txt");
            fos.write(string.getBytes());
        }




    }

    private static void convertexcel(String source, String endway) throws IOException {

        InputStream fs = new FileInputStream(source);
        XSSFWorkbook wb = new XSSFWorkbook(fs);
        Sheet[] sheets = new Sheet[wb.getNumberOfSheets()];
        for (int s = 0; s < wb.getNumberOfSheets(); s++) {
            Sheet sheet = wb.getSheetAt(s);
            try {

                int lenmergedregion = sheet.getMergedRegions().size();
                int[] CellStart = new int[lenmergedregion];
                int[] CellEnd = new int[lenmergedregion];
                int[] RowStart = new int[lenmergedregion];
                int[] RowEnd = new int[lenmergedregion];
                for (int i = 0; i < lenmergedregion; i++) {
                    CellStart[i] = sheet.getMergedRegions().get(i).getFirstColumn();
                    RowStart[i] = sheet.getMergedRegions().get(i).getFirstRow();
                    CellEnd[i] = sheet.getMergedRegions().get(i).getLastColumn();
                    RowEnd[i] = sheet.getMergedRegions().get(i).getLastRow();
                }

                for (int i = 0; i < lenmergedregion; i++) {
                    String mergedstring = "";
                    for (int r = RowStart[i]; r <= RowEnd[i]; r++) {

                        for (int c = CellStart[i]; c <= CellEnd[i]; c++) {
                            if (sheet.getRow(r).getCell(c).toString().length() > 1) {
                                mergedstring = sheet.getRow(r).getCell(c).toString();
                            }

                        }

                    }
                    for (int r = RowStart[i]; r <= RowEnd[i]; r++) {

                        for (int c = CellStart[i]; c <= CellEnd[i]; c++) {
                            sheet.getRow(r).getCell(c).setCellValue(mergedstring);
                        }
                    }
                }


                for (int i = 0; i < sheet.getNumMergedRegions(); ++i) {

                    sheet.removeMergedRegion(i);
                }


            } catch (Exception e) {
                e.printStackTrace();
            }


        }

        try (OutputStream fileOut = new FileOutputStream(endway)) {
            wb.write(fileOut);
        }
        wb.close();

    }

    private static void downloadUsingStream(String urlStr, String file) throws IOException {
        URL url = new URL(urlStr);
        BufferedInputStream bis = new BufferedInputStream(url.openStream());
        FileOutputStream fis = new FileOutputStream(file);
        byte[] buffer = new byte[1024];
        int count = 0;
        while ((count = bis.read(buffer, 0, 1024)) != -1) {
            fis.write(buffer, 0, count);
        }
        fis.close();
        bis.close();
    }

}
