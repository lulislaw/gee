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
import java.util.Arrays;
import java.util.Locale;

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
        url[0] = "https://guu.ru/wp-content/uploads/1-%D0%BA%D1%83%D1%80%D1%81-%D0%B1%D0%B0%D0%BA%D0%B0%D0%BB%D0%B0%D0%B2%D1%80%D0%B8%D0%B0%D1%82-%D0%9E%D0%A4%D0%9E-37.xlsx";
        url[1] = "https://guu.ru/wp-content/uploads/2-%D0%BA%D1%83%D1%80%D1%81-%D0%B1%D0%B0%D0%BA%D0%B0%D0%BB%D0%B0%D0%B2%D1%80%D0%B8%D0%B0%D1%82-%D0%9E%D0%A4%D0%9E-37.xlsx";
        url[2] = "https://guu.ru/wp-content/uploads/3-%D0%BA%D1%83%D1%80%D1%81-%D0%B1%D0%B0%D0%BA%D0%B0%D0%BB%D0%B0%D0%B2%D1%80%D0%B8%D0%B0%D1%82-%D0%9E%D0%A4%D0%9E-40.xlsx";
        url[3] = "https://guu.ru/wp-content/uploads/4-%D0%BA%D1%83%D1%80%D1%81-%D0%B1%D0%B0%D0%BA%D0%B0%D0%BB%D0%B0%D0%B2%D1%80%D0%B8%D0%B0%D1%82-%D0%9E%D0%A4%D0%9E-32.xlsx";
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


                            try {
                                if (sheet.getRow(r).getCell(c).toString() != null) {
                                    if (sheet.getRow(r).getCell(c).toString().length() > 2 && !sheet.getRow(r).getCell(c).toString().contains("Директор") && !string.contains(sheet.getRow(r).getCell(c).toString()
                                            + "Y" + sheet.getRow(r).getCell(3).toString()
                                            + "Y" + sheet.getRow(r).getCell(2).toString()
                                            + "Y" + sheet.getRow(r).getCell(1).toString()
                                            + "Y" + "Курс-" + (b + 1))) {
                                        string = string + "\nX\n" + sheet.getRow(r).getCell(c).toString()
                                                + "Y" + sheet.getRow(r).getCell(3).toString()
                                                + "Y" + sheet.getRow(r).getCell(2).toString()
                                                + "Y" + sheet.getRow(r).getCell(1).toString()
                                                + "Y" + "Курс-" + (b + 1) + "Y";

                                        if (sheet.getRow(r).getCell(c).toString().contains("(Л")) {
                                            string = string + "Лекция" + "Y";
                                        } else {
                                            string = string + "ЛЗ/ПЗ" + "Y";
                                        }

                                        if (sheet.getRow(r).getCell(3).toString().contains("НЕЧЁТ")) {
                                            string = string + "Нечётная" + "Y";
                                        } else {
                                            string = string + "Четная" + "Y";
                                        }
                                    }
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

                for (int c = 4; c < 60; c++) {
                    for (int r = 8; r < 60; r++) {
                        try {
                            if (sheet.getRow(r).getCell(c) != null) {

                                sheet.getRow(r).getCell(c).setCellValue(decomposition(sheet.getRow(r).getCell(c).toString()));

                            }

                        } catch (Exception e) {
                            e.printStackTrace();
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

    private static String decomposition(String sourceString) {
        sourceString = sourceString.replace('\n', ' ');
        String[] tempstring = new String[4];
        /*
         *  tempstring[0] - Название
         *  tempstring[1] - Тип
         *  tempstring[2] - Препод
         *  tempstring[3] - Аудитория
         */
        tempstring[0] = sourceString.split("\\(")[0];
        if (sourceString.contains("(ЛЗ"))
            tempstring[1] = "Лабораторное занятие";
        else if (sourceString.contains("(ПЗ"))
            tempstring[1] = "Практическое занятие";
        else if (sourceString.contains("(Л"))
            tempstring[1] = "Лекция";
        else
            tempstring[1] = "null";
        int temp_index = 0;
        for (int i = 1; i < sourceString.length() - 1; i++) {
            if (sourceString.charAt(i - 1) == '.' && Character.isUpperCase(sourceString.charAt(i)) && sourceString.charAt(i + 1) == '.') {
                temp_index = i;
                break;
            }
        }
        if(temp_index != 0) {
            for (int i = temp_index - 3; i > 0; i--) {
                if (Character.isUpperCase(sourceString.charAt(i))) {
                    tempstring[2] = sourceString.substring(i, temp_index + 2);

                    break;
                }
            }
        }
        else
        {
            tempstring[2] = "null";
        }
        for (int i = 10; i < 900; i++) {
            if (sourceString.toLowerCase().contains("этаж")) {
                tempstring[3] = "6 этаж";
                break;
            }
            if (sourceString.toLowerCase().contains("спортивный")) {
                tempstring[3] = "Спортивный комплекс";
                break;
            }
            //А У ЛК ПА
            if (i < 100) {
                if (sourceString.toLowerCase().contains("па-" + i)) {
                    tempstring[3] = "ПА-" + i;
                    break;
                }
            } else {
                if (sourceString.toLowerCase().contains("лк-" + i)) {
                    tempstring[3] = "ЛК-" + i;
                    break;
                } else if (sourceString.toLowerCase().contains("у-" + i)) {
                    tempstring[3] = "У-" + i;
                    break;
                } else if (sourceString.toLowerCase().contains("а-" + i)) {
                    tempstring[3] = "А-" + i;
                    break;
                }
             else if (sourceString.toLowerCase().contains("цувп-" + i)) {
                tempstring[3] = "цувп-" + i;
                break;
            }

            }

        }
        String alphabet = "abcdefghijklmnopqrstuvwxyzабвгдежзийклмнопрстуфхцчшщъыьэюя";

        for(int i = 0; i < alphabet.length() ; i++)
        {
            if(sourceString.toLowerCase().contains(tempstring[3].toLowerCase()+alphabet.charAt(i)))
            {
                tempstring[3] = tempstring[3] + alphabet.charAt(i);
            }
        }




        tempstring[0] = tempstring[0].replaceAll(tempstring[2], "");
        return tempstring[0] + "\n" + tempstring[1] + "\n" + tempstring[2] + "\n" + tempstring[3];
    }

}
