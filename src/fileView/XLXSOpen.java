package fileView;

import data.InfoList;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

public class XLXSOpen {
    String fileName;
    Workbook workbook;
    public XLXSOpen(File file) throws IOException, InvalidFormatException {
        String filePath = file.getPath();
        fileName = file.getName();
        workbook = new XSSFWorkbook(new FileInputStream(filePath));
    }

    public void getClose() throws IOException {
        workbook.close();
    }

    public void getBacterInfo(InfoList infoList){
        for(int i = 0; i < workbook.getSheetAt(2).getPhysicalNumberOfRows(); i++){
            for (int j = 0; j < infoList.bacterInfo.size();j++){
                if(workbook.getSheetAt(2).getRow(i).getCell(0).getStringCellValue().equals(infoList.bacterInfo.get(j).get(0))){
                    Double temp = workbook.getSheetAt(2).getRow(i).getCell(1).getNumericCellValue();
                    infoList.bacterInfo.get(j).add(Double.toString(temp));
                }
            }
        }
    }

    public void getMainInfo(InfoList infoList) throws IOException {
        infoList.mainInfo.add(workbook.getSheetAt(6).getRow(0).getCell(0).getStringCellValue());
        Double temp = workbook.getSheetAt(6).getRow(0).getCell(1).getNumericCellValue();
        infoList.mainInfo.add(Double.toString(temp));
        infoList.mainInfo.add(workbook.getSheetAt(6).getRow(0).getCell(2).getStringCellValue());
    }

    public void getFileName(InfoList infoList){
        infoList.fileName = fileName;
    }
}