import data.InfoList;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

public class AlgOpen {
    public AlgOpen(InfoList infoList) throws IOException, InvalidFormatException {
        File file = new File(Application.rootDirPath + "\\algs.xlsx");
        String filePath = file.getPath();
        Workbook workbook = new XSSFWorkbook(new FileInputStream(filePath));
        for(int i = 0; i < workbook.getSheetAt(0).getPhysicalNumberOfRows();i++)
        {
            infoList.algs.add(new ArrayList<>());
            infoList.algs.get(i).add(workbook.getSheetAt(0).getRow(i).getCell(0).getStringCellValue());
            if(workbook.getSheetAt(0).getRow(i).getCell(1).getCellType() == CellType.NUMERIC){
                Double num = workbook.getSheetAt(0).getRow(i).getCell(1).getNumericCellValue();
                infoList.algs.get(i).add(num.toString());
            } else {
                infoList.algs.get(i).add(workbook.getSheetAt(0).getRow(i).getCell(1).getStringCellValue().replace(",", "."));
            }
            infoList.algs.get(i).add(workbook.getSheetAt(0).getRow(i).getCell(2).getStringCellValue());
        }
        workbook.close();
    }
}
