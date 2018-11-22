package threepm;


import java.io.File;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import de.siegmar.fastcsv.writer.CsvWriter;

/**
 *
 * @author the_fegati
 */
public class ExcelReader {

    private static File file;
    DataFormatter formatter = new DataFormatter();
    private String reportingPeriod;
        
    public String getReportingPeriod() {
        return reportingPeriod;
    }
    
    public void setReportingPeriod(String reportingPeriod) {
        this.reportingPeriod = reportingPeriod;
    }

    public ExcelReader(File file) {
        ExcelReader.file = file;
    }

    public String getRowAsListFromExcel() {
        List<String[]> csvList = new ArrayList<String[]>();
        Workbook workbook;
        try {
            
            if (file == null) {
                return "Select file to extract";
            }
            
            if (reportingPeriod.equals("") || !reportingPeriod.matches("^[0-9]{6}$")) {
                return "Missing or invalid reporting period";
            }
            
            String fileExtension = file.toString().substring(file.toString().indexOf("."));
            System.err.println("The file extension is " + fileExtension);
            //use xssf for xlsx format else hssf for xls format
            if (fileExtension.toLowerCase().equals(".xlsx")) {
                    OPCPackage oPCPackage = OPCPackage.open(file);
                    workbook = new XSSFWorkbook(oPCPackage);
            }
            else if (fileExtension.equals(".xls")) {
//                    workbook = new HSSFWorkbook(new POIFSFileSystem(fis));
                    System.err.println("Wrong file type selected! Should be .xlsx");
                    return "Wrong file type selected!";
            } else {
                    System.err.println("Wrong file type selected!");
                    return "Wrong file type selected!";
            }
            
            

//          get number of worksheets in the workbook
            int numberOfSheets = 9;
            String currentPeriod = reportingPeriod;
            String[] dataRows = new String[8];
            dataRows[0] = "dataelementUID";
            dataRows[1] = "period";
            dataRows[2] = "orgUnitUID";
            dataRows[3] = "categoryOptionComboUID";
            dataRows[4] = "ImplementingMechanismUID";
            dataRows[5] = "dataValue";

            csvList.add(dataRows);

//          iterating over each workbook sheet
            for (int i = 1; i <= numberOfSheets; i++) {
                Sheet sheet = workbook.getSheetAt(i);

                String fileNameSuffix = "";

//              int period = (int) sheet.getRow(1).getCell(3, Row.CREATE_NULL_AS_BLANK).getNumericCellValue();
                int numberOfRows = sheet.getLastRowNum();
                int startRow = 9;
                int lastRow = 9;
                
//              Compute the first row containing data
                for (int row = 0; row < numberOfRows; row++) {
                    Row currentRow = sheet.getRow(row);
                    String cellValue = "";
                    for (int cell = 0; cell < 7; cell++) {
                        cellValue = currentRow.getCell(cell, Row.CREATE_NULL_AS_BLANK).getStringCellValue();
                        if (cellValue.toUpperCase().equals("ICAP")) {
                            startRow = row;
                            break;
                        }
                    }
                }
                
                for (int row = startRow; row <= lastRow; row++) {
                    Row currentRow = sheet.getRow(row);
                    int numberOfCells = currentRow.getLastCellNum();
                    DataFormatter dataFormatter = new DataFormatter();
                    String formattedPeriod = dataFormatter.formatCellValue(currentRow.getCell(5));

                    int period = (int) currentRow.getCell(5, Row.CREATE_NULL_AS_BLANK).getNumericCellValue();
                    String facility = currentRow.getCell(0, Row.CREATE_NULL_AS_BLANK).getStringCellValue();
                    String implementingMechanism = sheet.getRow(2).getCell(1, Row.CREATE_NULL_AS_BLANK).getStringCellValue();
                    String orgUnitUUID = sheet.getRow(1).getCell(1, Row.CREATE_NULL_AS_BLANK).getStringCellValue();
                    
                    
                    System.out.println("Number of Rows is " + numberOfRows);
                    System.out.println("Period is " + " " + formattedPeriod + " " + period);
                    System.out.println("Facility is " + facility);
//                    System.exit(0);

//                    String implementingMechanism = sheet.getRow(2).getCell(1, Row.CREATE_NULL_AS_BLANK).getStringCellValue();

                    for (int cell = 6; cell < numberOfCells; cell++) {
                        String dataelementUID = sheet.getRow(2).getCell(cell, Row.CREATE_NULL_AS_BLANK).getStringCellValue();
                        String categoryOptionCombo = sheet.getRow(3).getCell(cell, Row.CREATE_NULL_AS_BLANK).getStringCellValue();
                        int dataValue = 0;
                        try {
                            dataValue = (int) currentRow.getCell(cell, Row.CREATE_NULL_AS_BLANK).getNumericCellValue();
                        } catch (Exception e) {
                        }
//                        if (dataValue == 0) {
//                            continue;
//                        }
//                        System.out.println("Data element " + dataelementUID.equals("") + ". Period " + formattedPeriod.equals("") + ". Data Value " + String.valueOf(dataValue).equals(""));
                        if (!dataelementUID.equals("") && !String.valueOf(dataValue).equals("") && formattedPeriod.equals(currentPeriod)) {
                            dataRows = new String[8];
                            dataRows[0] = String.valueOf(dataelementUID);
                            dataRows[1] = String.valueOf(formattedPeriod);
                            dataRows[2] = orgUnitUUID;
                            dataRows[3] = categoryOptionCombo;
                            dataRows[4] = implementingMechanism;
                            dataRows[5] = String.valueOf(dataValue);
                            csvList.add(dataRows);
                        
                            System.out.printf("%4d%4d%16s%8s%17s%17s%17s%10d\n", row, cell, dataelementUID, period, facility, categoryOptionCombo, implementingMechanism, dataValue);
                            fileNameSuffix = sheet.getSheetName() + "_" + formattedPeriod;
                        }
                    }
                }
                System.out.println(fileNameSuffix);
                writeRowToCSVFile( fileNameSuffix, csvList);
            }

//            System.out.println("");
            workbook.close();
            File dir = new File("./Output");
            return "CSVs successfully extracted to <br/> '" + dir.getAbsolutePath().replace(".\\", "") + "'";
        } catch (IOException e) {
            e.printStackTrace();
            return "CSVs extraction failed";
        } catch (InvalidFormatException e) {
            return "CSVs extraction failed";            
        }
    }

    /*
     * Write the rows into the CSV file
     */
    private static void writeRowToCSVFile(String filenameSuffix, List<String[]> cleanRows)
            throws IOException {
        
//        String fileName = "C:\\Users\\HISProgrammer\\Downloads\\3PM Import Files Assistance\\output\\" + file.getName().substring(0, file.getName().indexOf(".")) + "_" + filenameSuffix + ".csv";
        File dir = new File("./Output");
        dir.mkdirs();
        String fileName = "./Output/" + file.getName().substring(0, file.getName().indexOf(".")) + "_" + filenameSuffix + ".csv";
        File newFile = new File(fileName); 
        CsvWriter csvWriter = new CsvWriter(); 
        csvWriter.write(newFile, StandardCharsets.UTF_8, cleanRows);
        System.out.println(newFile.getAbsolutePath());
            
    }
    
}
