import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;



public class ExcelReader {

    String path;

    public List<List<String>> readExcel(String path)
    {
        this.path = path;
        List<List<String>> records = new ArrayList<>();
        try
        {
            FileInputStream file = new FileInputStream(new File(path));
            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            //Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);
            //Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();
            //to evaluate cell type
            FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
            for(Row row : sheet)
            {
                if(row.getRowNum()>4) {
                    List<String> columns = new ArrayList<>();
                    for(int j=0; j<85; j++){
                       Cell cell = row.getCell(j);
                        if(cell !=null) {
                            cell.setCellType(Cell.CELL_TYPE_STRING);
                            switch (cell.getCellTypeEnum()) {
                                case STRING:
                                     columns.add(cell.getStringCellValue());
                                    break;
                                default:
                                    break;

                            }
                        }else{
                            columns.add("");
                        }

                    }
                    records.add(columns);

                }


            }

            file.close();
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
        return records;
    }
}
