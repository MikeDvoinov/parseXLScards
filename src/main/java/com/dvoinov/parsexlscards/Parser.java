package com.dvoinov.parsexlscards;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * Created by dvoinov on 31.08.2016.
 */
public class Parser {

    public static List<Card> cardList;

    private static int cellName=3, cellAcc=4, cellCard=2, cellPanel=1, cellDate=0;

    public static void main(String[] args) throws IOException {
        cardList=new ArrayList<Card>();
        readFromExcel("c:/temp/111/events2.xlsx");
        readFromExcelSys("c:/temp/111/sysevent2.xlsx");
        writeIntoExcel("c:/temp/111/svod2.xlsx",cardList);
    }

    public static void readFromExcel(String file) throws IOException {
        XSSFWorkbook myExcelBook = new XSSFWorkbook(new FileInputStream(file));
        XSSFSheet myExcelSheet = myExcelBook.getSheet("Лист1");

        int rowNumber=0;

        XSSFRow row = myExcelSheet.getRow(rowNumber);

        try {
            while (true) {

                row = myExcelSheet.getRow(rowNumber);
                if(row.getCell(cellAcc).getCellType() == XSSFCell.CELL_TYPE_STRING){
                    String access = row.getCell(cellAcc).getStringCellValue();
                    if (access.equals("Доступ по карте")) {

                        if(row.getCell(cellCard).getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
                            int id = (int) row.getCell(cellCard).getNumericCellValue();
                            Card card = isCardAlreadyAddedToList(id);
                            if(row.getCell(cellName).getCellType() == XSSFCell.CELL_TYPE_STRING) {
                                String name = row.getCell(cellName).getStringCellValue();
                                card.setCardHolder(name);
                            }
                            if(row.getCell(cellPanel).getCellType() == XSSFCell.CELL_TYPE_STRING) {
                                String place = row.getCell(cellPanel).getStringCellValue();
                                if (place.equals("Серверная")) card.setAccServ(true);
                            }
                            if(row.getCell(cellPanel).getCellType() == XSSFCell.CELL_TYPE_STRING) {
                                String place = row.getCell(cellPanel).getStringCellValue();
                                if (place.equals("Редакторы")) card.setAccBuhg(true);
                            }
                            if(row.getCell(cellPanel).getCellType() == XSSFCell.CELL_TYPE_STRING) {
                                String place = row.getCell(cellPanel).getStringCellValue();
                                if (place.equals("Директор")) card.setAccDir(true);
                            }
                            if(row.getCell(cellPanel).getCellType() == XSSFCell.CELL_TYPE_STRING) {
                                String place = row.getCell(cellPanel).getStringCellValue();
                                if (place.equals("Кладовка")) card.setAccKlad(true);
                            }
                            if(row.getCell(cellPanel).getCellType() == XSSFCell.CELL_TYPE_STRING) {
                                String place = row.getCell(cellPanel).getStringCellValue();
                                if (place.equals("Вход1")) card.setAccEnter1(true);
                            }
                            if(row.getCell(cellPanel).getCellType() == XSSFCell.CELL_TYPE_STRING) {
                                String place = row.getCell(cellPanel).getStringCellValue();
                                if (place.equals("Выход1")) card.setAccExit1(true);
                            }
                            if(row.getCell(cellPanel).getCellType() == XSSFCell.CELL_TYPE_STRING) {
                                String place = row.getCell(cellPanel).getStringCellValue();
                                if (place.equals("Вход2")) card.setAccEnter2(true);
                            }
                            if(row.getCell(cellPanel).getCellType() == XSSFCell.CELL_TYPE_STRING) {
                                String place = row.getCell(cellPanel).getStringCellValue();
                                if (place.equals("Выход2")) card.setAccExit2(true);
                            }

                            if(row.getCell(cellDate).getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
                                Date lastAccessDate = row.getCell(cellDate).getDateCellValue();
                                card.setLastAccess(lastAccessDate);
                            }
                        }
                    }
                }
                rowNumber++;
            }
        }
        catch (NullPointerException e){
            System.out.println("Пустая ячейка. Парсинг завершен!");
        }
        myExcelBook.close();

    }

    public static Card isCardAlreadyAddedToList(int id){
        for (Card card : cardList) {
            if (card.getCardId()==id) return card;
        }
        Card card = new Card();
        card.setCardId(id);
        cardList.add(card);
        return card;
    }

    public static void writeIntoExcel(String file, List<Card> list) throws FileNotFoundException, IOException{
        Workbook book = new XSSFWorkbook();
        Sheet sheet = book.createSheet("Лист2");
        int rowCounter=0;
        boolean acc=false;

        DataFormat format = book.createDataFormat();
        CellStyle dateStyle = book.createCellStyle();
        dateStyle.setDataFormat(format.getFormat("dd.mm.yyyy"));


        Row row = sheet.createRow(rowCounter);
        Cell cell = row.createCell(0);
        cell.setCellValue("ФИО");
        cell=row.createCell(1);
        cell.setCellValue("Карта");
        cell=row.createCell(2);
        cell.setCellValue("Заблокирована");
        cell=row.createCell(3);
        cell.setCellValue("Дата последнего использования");


        cell=row.createCell(4);
        cell.setCellValue("Серверная");
        cell=row.createCell(5);
        cell.setCellValue("Бухгалтерия");
        cell=row.createCell(6);
        cell.setCellValue("Директор");
        cell=row.createCell(7);
        cell.setCellValue("Кладовка");

        cell=row.createCell(8);
        cell.setCellValue("Вход1");
        cell=row.createCell(9);
        cell.setCellValue("Выход1");
        cell=row.createCell(10);
        cell.setCellValue("Вход2");
        cell=row.createCell(11);
        cell.setCellValue("Выход2");
        rowCounter++;

        for (Card card: list ) {
            row = sheet.createRow(rowCounter);
            cell = row.createCell(0);
            cell.setCellValue(card.getCardHolder());

            cell=row.createCell(1);
            cell.setCellValue(card.getCardId());
            cell=row.createCell(2);
            boolean block = card.isBlocked();
            if (block==true) cell.setCellValue("БЛОК");

            cell=row.createCell(3);
            cell.setCellStyle(dateStyle);
            cell.setCellValue(card.getLastAccess());
            cell=row.createCell(4);
            acc=card.isAccServ();
            if (acc==true) cell.setCellValue("Да");
            cell=row.createCell(5);
            acc=card.isAccBuhg();
            if (acc==true) cell.setCellValue("Да");
            cell=row.createCell(6);
            acc=card.isAccDir();
            if (acc==true) cell.setCellValue("Да");
            cell=row.createCell(7);
            acc=card.isAccKlad();
            if (acc==true) cell.setCellValue("Да");

            cell=row.createCell(8);
            acc=card.isAccEnter1();
            if (acc==true) cell.setCellValue("Да");
            cell=row.createCell(9);
            acc=card.isAccExit1();
            if (acc==true) cell.setCellValue("Да");
            cell=row.createCell(10);
            acc=card.isAccEnter2();
            if (acc==true) cell.setCellValue("Да");
            cell=row.createCell(11);
            acc=card.isAccExit2();
            if (acc==true) cell.setCellValue("Да");




            rowCounter++;
        }

        // Меняем размер столбца
        for (int i = 0; i < 12; i++) {
            sheet.autoSizeColumn(i);
        }


        // Записываем всё в файл
        book.write(new FileOutputStream(file));
        book.close();
    }

    public static void readFromExcelSys(String file) throws IOException {
        XSSFWorkbook myExcelBook = new XSSFWorkbook(new FileInputStream(file));
        XSSFSheet myExcelSheet = myExcelBook.getSheet("Лист1");

        int rowNumber=0;

        XSSFRow row = myExcelSheet.getRow(rowNumber);

        try {
            while (true) {

                row = myExcelSheet.getRow(rowNumber);

                if(row.getCell(1).getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
                    int id = (int) row.getCell(1).getNumericCellValue();
                    Card card = isCardAlreadyAddedToList(id);

                    if(row.getCell(0).getCellType() == XSSFCell.CELL_TYPE_STRING) {
                        String blocked = row.getCell(0).getStringCellValue();
                        if (blocked.equals("Блокировка")) card.setBlocked(true);
                        else if (blocked.equals("Разблокировка")) card.setBlocked(false);
                    }
                }

                rowNumber++;
            }
        }
        catch (NullPointerException e){
            System.out.println("Пустая ячейка. Парсинг завершен!");
        }
        myExcelBook.close();
    }


}

