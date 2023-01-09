import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;

public class AppInit {

    private static FileInputStream existingFile = null;
    private static Workbook workbook = null;

    public static void main(String[] args) throws IOException, InvalidFormatException {
        System.out.println("App Init!!!!");
        AppInit obj = new AppInit();
        existingFile = obj.checkExistingFile();
        workbook =  new XSSFWorkbook(existingFile);
        obj.setCellValue("A1", 13);
        obj.setCellValue("A2", 14);

        System.out.println(obj.getCellValue("A1"));

        obj.setCellValue("A3", "=A1+A2");
        obj.setCellValue("A4", "=A1+A2+A3");

        System.out.println(obj.getCellValue("A4"));

        obj.setCellValue("A5", "Rajan");

        // Closing the InputFileStream before moving to save the changes.
        existingFile.close();

        //New FileOutputStream with same name and changes.
        FileOutputStream outputStream = new FileOutputStream("src/main/resources/students.xlsx");
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
    }

    private FileInputStream checkExistingFile() throws FileNotFoundException {
        // Creating file object of existing excel file
        File xlsxFile = new File("src/main/resources/students.xlsx");

        //Creating workbook from input stream
        //Creating input stream
        FileInputStream inputStream = new FileInputStream(xlsxFile);

        return  inputStream;

    }

    private void setCellValue(String cellId, Object value) throws IOException, InvalidFormatException {

        Sheet sheet = workbook.getSheetAt(0);
        CellAddress cellAddress = new CellAddress(cellId);
        Row row = sheet.createRow(cellAddress.getRow());
        Cell cell = row.createCell(cellAddress.getColumn());

        if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
        }
        else if (value instanceof String ) {
            if (((String) value).startsWith("=")) {
                value = ((String) value).replaceAll("=", "");
                FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                cell.setCellFormula((String) value);
                evaluator.evaluateFormulaCell(cell);
            }
            else {
                cell.setCellValue((String) value);
            }

        }
    }

    private int getCellValue(String cellId) throws IOException, RuntimeException {

        if ( cellId == null)
            throw new RuntimeException("Invalid cellID!");

        Sheet sheet = workbook.getSheetAt(0);
        CellAddress cellAddress = new CellAddress(cellId);
        Row row = sheet.getRow(cellAddress.getRow());
        Cell cell = row.getCell(cellAddress.getColumn());
        return (int) cell.getNumericCellValue();
    }

}




