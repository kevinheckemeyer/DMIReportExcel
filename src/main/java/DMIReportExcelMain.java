import java.util.List;

public class DMIReportExcelMain {

    public static void main(String args[]){
        try {
            ExcelReader excelReader = new ExcelReader();
            DMIReportWriter dmiReportWriter = new DMIReportWriter();
            String basePath = "/Users/georgepeter/Downloads/DMI Reports";
            String reportFilePath = basePath + "/report_test.xlsx";
            List<List<String>> records = excelReader.readExcel(reportFilePath);
            dmiReportWriter.writeToExcel(records, basePath);
            System.out.println("success");
        }catch(Exception ex){
            System.out.println(ex);
        }

    }
}
