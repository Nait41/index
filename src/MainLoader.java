import data.InfoList;
import fileView.XLXSOpen;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.ArrayList;

public class MainLoader extends JFrame {
    Workbook workbook;
    public MainLoader(File file) throws IOException, InvalidFormatException {
        String filePath = file.getPath();
        workbook = new XSSFWorkbook(new FileInputStream(filePath));
    }

    public void getClose() throws IOException {
    }

    public void setAllInfo(InfoList infoList){
        ArrayList<String> newBacteriasList = new ArrayList<>();
        ArrayList<String> removingBacterias = new ArrayList<>();
        ArrayList<ArrayList<String>> tempList = new ArrayList<>();

        for (int i =0; i < infoList.bacterInfo.size();i++){
            newBacteriasList.add(infoList.bacterInfo.get(i).get(0));
        }

        for (int j = 0; j < workbook.getSheetAt(0).getRow(0).getPhysicalNumberOfCells(); j++) {
            if (workbook.getSheetAt(0).getRow(0).getCell(j).getStringCellValue().length() > 0){
                if(workbook.getSheetAt(0).getRow(0).getCell(j).getStringCellValue().charAt(0) >= 'A'
                        && workbook.getSheetAt(0).getRow(0).getCell(j).getStringCellValue().charAt(0) <= 'Z'
                        && !workbook.getSheetAt(0).getRow(0).getCell(j).getStringCellValue().contains("MGG")
                        && !newBacteriasList.contains(workbook.getSheetAt(0).getRow(0).getCell(j).getStringCellValue())
                ){
                    removingBacterias.add(workbook.getSheetAt(0).getRow(0).getCell(j).getStringCellValue());
                }
            }
            for (int k = 0; k < infoList.bacterInfo.size(); k++) {
                if (workbook.getSheetAt(0).getRow(0).getCell(j).getStringCellValue().equals(infoList.bacterInfo.get(k).get(0))) {
                    newBacteriasList.remove(infoList.bacterInfo.get(k).get(0));
                }
            }
        }

        if(removingBacterias.size() != 0){
            for(int l = 0; l < removingBacterias.size(); l++){
                boolean checkTemp = false;
                int startPosition = 0;

                for (int j = 1; j < workbook.getSheetAt(0).getRow(0).getPhysicalNumberOfCells(); j++) {
                    if(workbook.getSheetAt(0).getRow(0).getCell(j).getStringCellValue().equals(removingBacterias.get(l)) && !checkTemp)
                    {
                        for(int p = 0; p < workbook.getSheetAt(0).getPhysicalNumberOfRows(); p++){
                            workbook.getSheetAt(0).getRow(p).createCell(j);
                        }
                        checkTemp = true;
                        startPosition = j;
                        ++j;
                    }
                    if(checkTemp)
                    {
                        tempList.add(new ArrayList<>());
                        for (int i = 0; i < workbook.getSheetAt(0).getPhysicalNumberOfRows();i++) {
                            if(workbook.getSheetAt(0).getRow(i).getCell(j) == null){
                                tempList.get(tempList.size()-1).add("");
                            } else {
                                if(workbook.getSheetAt(0).getRow(i).getCell(j).getCellType().equals(CellType.NUMERIC))
                                {
                                    Double temp = workbook.getSheetAt(0).getRow(i).getCell(j).getNumericCellValue();
                                    tempList.get(tempList.size()-1).add(Double.toString(temp));
                                } else if(workbook.getSheetAt(0).getRow(i).getCell(j).getCellType().equals(CellType.FORMULA))
                                {
                                    workbook.getSheetAt(0).getRow(i).getCell(j).removeFormula();
                                    tempList.get(tempList.size()-1).add("");
                                }
                                else {
                                    tempList.get(tempList.size()-1).add(workbook.getSheetAt(0).getRow(i).getCell(j).getStringCellValue());
                                }
                            }
                            workbook.getSheetAt(0).getRow(i).createCell(j);
                        }
                    }
                }

                DecimalFormat df = new DecimalFormat("###.#");

                for (int i = 0; i < tempList.size(); i++) {
                    for (int j = 0; j < tempList.get(i).size(); j++){
                        if (tempList.get(i).get(0).equals("ИМТ") && j>0 && !tempList.get(i-1).get(j).equals("")){
                            String temp = Double.toString(Double.parseDouble(tempList.get(i-1).get(j))/Math.pow(Double.parseDouble(tempList.get(i-2).get(j)), 2));
                            workbook.getSheetAt(0).getRow(j).createCell(i+startPosition)
                                    .setCellValue(Double.parseDouble(
                                            df.format(Double.parseDouble(temp)).replace(",", ".")));
                            workbook.getSheetAt(0).getRow(j).getCell(i+startPosition).setCellStyle(workbook.getSheetAt(0).getRow(1).getCell(8).getCellStyle());
                        } else {
                            workbook.getSheetAt(0).getRow(j).createCell(i+startPosition).setCellValue(tempList.get(i).get(j));
                            workbook.getSheetAt(0).getRow(j).getCell(i+startPosition).setCellStyle(workbook.getSheetAt(0).getRow(1).getCell(8).getCellStyle());
                        }
                        if (j==0)
                        {
                            workbook.getSheetAt(0).getRow(0).getCell(i+startPosition).setCellStyle(workbook.getSheetAt(0).getRow(0).getCell(8).getCellStyle());
                            workbook.getSheetAt(0).setColumnWidth(i+startPosition, 6000);
                        }
                    }
                }

                int countRightColumn = 0;
                for (int i = 0; i < workbook.getSheetAt(0).getRow(0).getPhysicalNumberOfCells();i++)
                {
                    if(!workbook.getSheetAt(0).getRow(0).getCell(i).getStringCellValue().equals("")){
                        countRightColumn++;
                    }
                }

                workbook.getSheetAt(0).setAutoFilter(new CellRangeAddress(0, workbook.getSheetAt(0).getPhysicalNumberOfRows(),
                        0, --countRightColumn));

                tempList = new ArrayList<>();
            }
        }
        if(newBacteriasList.size() != 0) {
            boolean checkTemp = false;
            int startPosition = 0;

            for (int j = 1; j < workbook.getSheetAt(0).getRow(0).getPhysicalNumberOfCells(); j++) {
                if(workbook.getSheetAt(0).getRow(0).getCell(j).getStringCellValue().equals("Место проживания") && !checkTemp
                ) {
                    checkTemp = true;
                    startPosition = j;
                }
                if(checkTemp)
                {
                    tempList.add(new ArrayList<>());
                    for (int i = 0; i < workbook.getSheetAt(0).getPhysicalNumberOfRows();i++) {
                        if(workbook.getSheetAt(0).getRow(i).getCell(j) == null){
                            tempList.get(tempList.size()-1).add("");
                        } else {
                            if(workbook.getSheetAt(0).getRow(i).getCell(j).getCellType().equals(CellType.NUMERIC))
                            {
                                Double temp = workbook.getSheetAt(0).getRow(i).getCell(j).getNumericCellValue();
                                tempList.get(tempList.size()-1).add(Double.toString(temp));
                            } else if(workbook.getSheetAt(0).getRow(i).getCell(j).getCellType().equals(CellType.FORMULA))
                            {
                                workbook.getSheetAt(0).getRow(i).getCell(j).removeFormula();
                                tempList.get(tempList.size()-1).add("");
                            }
                            else {
                                tempList.get(tempList.size()-1).add(workbook.getSheetAt(0).getRow(i).getCell(j).getStringCellValue());
                            }
                        }
                        workbook.getSheetAt(0).getRow(i).createCell(j);
                    }
                }
            }


            for (int i = 0; i < newBacteriasList.size();i++)
            {
                workbook.getSheetAt(0).getRow(0).createCell(startPosition+i).setCellValue(newBacteriasList.get(i));
                workbook.getSheetAt(0).setColumnWidth(startPosition+i, 6000);
                workbook.getSheetAt(0).getRow(0).getCell(startPosition+i).setCellStyle(workbook.getSheetAt(0).getRow(0).getCell(8).getCellStyle());
            }

            DecimalFormat df = new DecimalFormat("###.#");

            for (int i = 0; i < tempList.size(); i++) {
                for (int j = 0; j < tempList.get(i).size(); j++){
                    if (tempList.get(i).get(0).equals("ИМТ") && j>0 && !tempList.get(i-1).get(j).equals("")){
                        String temp = Double.toString(Double.parseDouble(tempList.get(i-1).get(j))/Math.pow(Double.parseDouble(tempList.get(i-2).get(j)), 2));
                        workbook.getSheetAt(0).getRow(j).createCell(i+startPosition+newBacteriasList.size())
                                .setCellValue(Double.parseDouble(
                                        df.format(Double.parseDouble(temp)).replace(",", ".")));
                        workbook.getSheetAt(0).getRow(j).getCell(i+startPosition+newBacteriasList.size()).setCellStyle(workbook.getSheetAt(0).getRow(1).getCell(8).getCellStyle());
                    } else {
                        workbook.getSheetAt(0).getRow(j).createCell(i+startPosition+newBacteriasList.size()).setCellValue(tempList.get(i).get(j));
                        workbook.getSheetAt(0).getRow(j).getCell(i+startPosition+newBacteriasList.size()).setCellStyle(workbook.getSheetAt(0).getRow(1).getCell(8).getCellStyle());
                    }
                    if (j==0)
                    {
                        workbook.getSheetAt(0).getRow(0).getCell(i+startPosition+newBacteriasList.size()).setCellStyle(workbook.getSheetAt(0).getRow(0).getCell(8).getCellStyle());
                        workbook.getSheetAt(0).setColumnWidth(i+startPosition+newBacteriasList.size(), 6000);
                    }
                }
            }

            int countRightColumn = 0;
            for (int i = 0; i < workbook.getSheetAt(0).getRow(0).getPhysicalNumberOfCells();i++)
            {
                if(!workbook.getSheetAt(0).getRow(0).getCell(i).getStringCellValue().equals("")){
                    countRightColumn++;
                }
            }

            workbook.getSheetAt(0).setAutoFilter(new CellRangeAddress(0, workbook.getSheetAt(0).getPhysicalNumberOfRows(),
                    0, --countRightColumn));
        }

        for (int i = 0; i < workbook.getSheetAt(0).getPhysicalNumberOfRows();i++)
        {
            if(workbook.getSheetAt(0).getRow(i).getCell(2).getStringCellValue().equals("")
                    || workbook.getSheetAt(0).getRow(i).getCell(2).getStringCellValue().contains(infoList.fileName.replace(".xlsx", ""))){
                workbook.getSheetAt(0).getRow(i).createCell(2).setCellValue(infoList.fileName.replace(".xlsx", ""));
                workbook.getSheetAt(0).getRow(i).getCell(2).setCellStyle(workbook.getSheetAt(0).getRow(i-1).getCell(2).getCellStyle());
                workbook.getSheetAt(0).getRow(i).createCell(8).setCellValue(infoList.mainInfo.get(1));
                if(infoList.mainInfo.get(2).contains("Низкий")){
                    workbook.getSheetAt(0).getRow(i).createCell(9).setCellValue("Низкое значение");
                } else if(infoList.mainInfo.get(2).contains("Средний")){
                    workbook.getSheetAt(0).getRow(i).createCell(9).setCellValue("Среднее значение");
                } else if(infoList.mainInfo.get(2).contains("Высокий")){
                    workbook.getSheetAt(0).getRow(i).createCell(9).setCellValue("Высокое значение");
                }
                workbook.getSheetAt(0).getRow(i).getCell(8).setCellStyle(workbook.getSheetAt(0).getRow(1).getCell(8).getCellStyle());
                workbook.getSheetAt(0).getRow(i).getCell(9).setCellStyle(workbook.getSheetAt(0).getRow(1).getCell(8).getCellStyle());

                for(int j = 0; j < workbook.getSheetAt(0).getRow(0).getPhysicalNumberOfCells();j++)
                {
                    for(int k = 0; k < infoList.bacterInfo.size(); k++){
                        if(workbook.getSheetAt(0).getRow(0).getCell(j).getStringCellValue().equals(infoList.bacterInfo.get(k).get(0))){
                            if (infoList.bacterInfo.get(k).size() == 1){
                                workbook.getSheetAt(0).getRow(i).createCell(j).setCellValue("Отсутствует");
                                workbook.getSheetAt(0).getRow(i).getCell(j).setCellStyle(workbook.getSheetAt(0).getRow(1).getCell(8).getCellStyle());
                            } else {
                                if(infoList.bacterInfo.get(k).get(1).equals("0.0")){
                                    workbook.getSheetAt(0).getRow(i).createCell(j).setCellValue("Отсутствует");
                                    workbook.getSheetAt(0).getRow(i).getCell(j).setCellStyle(workbook.getSheetAt(0).getRow(1).getCell(8).getCellStyle());
                                } else {
                                    for(int t = 0; t < infoList.algs.size(); t++){
                                        if (workbook.getSheetAt(0).getRow(0).getCell(j).getStringCellValue().equals(infoList.algs.get(t).get(0))
                                                && !infoList.algs.get(t).get(1).equals("0.0")){
                                            if(checkValueRange(infoList.algs.get(t).get(1), infoList.bacterInfo.get(k).get(1))){
                                                if(infoList.algs.get(t).get(2).contains("высокое") || infoList.algs.get(t).get(2).contains("Высокое")){
                                                    workbook.getSheetAt(0).getRow(i).getCell(j).setCellValue(infoList.algs.get(t).get(2).replace("высокое", "Высокое"));
                                                } else if(infoList.algs.get(t).get(2).contains("среднее") || infoList.algs.get(t).get(2).contains("Среднее")){
                                                    workbook.getSheetAt(0).getRow(i).getCell(j).setCellValue(infoList.algs.get(t).get(2).replace("среднее", "Среднее"));
                                                } else if(infoList.algs.get(t).get(2).contains("низкое") || infoList.algs.get(t).get(2).contains("Низкое")){
                                                    workbook.getSheetAt(0).getRow(i).getCell(j).setCellValue(infoList.algs.get(t).get(2).replace("низкое", "Низкое"));
                                                }
                                                workbook.getSheetAt(0).getRow(i).getCell(j).setCellStyle(workbook.getSheetAt(0).getRow(1).getCell(8).getCellStyle());
                                            }
                                        }
                                    }
                                }
                                if(workbook.getSheetAt(0).getRow(i).getCell(j).getStringCellValue().equals("")){
                                    workbook.getSheetAt(0).getRow(i).getCell(j).setCellValue(infoList.bacterInfo.get(k).get(1));
                                    workbook.getSheetAt(0).getRow(i).getCell(j).setCellStyle(workbook.getSheetAt(0).getRow(1).getCell(8).getCellStyle());
                                }
                            }
                        }
                    }
                }
                break;
            }
        }
    }

    boolean checkValueRange(String range, String checkNumber){
        String firstNumber = "", secondNumber = "";
        boolean checkChoice = true;
        for(int i = 0;i<range.length();i++){
            if(checkChoice){
                if(range.charAt(i) == '-' || range.charAt(i) == '–')
                {
                    checkChoice = false;
                }
                else{
                    firstNumber += range.charAt(i);
                }
            }else
            {
                secondNumber += range.charAt(i);
            }
        }
        if(Double.parseDouble(checkNumber.replace(",",".")) > Double.parseDouble(firstNumber.replace(",","."))
                && Double.parseDouble(checkNumber.replace(",",".")) < Double.parseDouble(secondNumber.replace(",","."))){
            return true;
        }
        else{
            return false;
        }
    }

    public void saveFile(File saveSample) throws IOException {
        workbook.write(new FileOutputStream(saveSample));
    }
}

