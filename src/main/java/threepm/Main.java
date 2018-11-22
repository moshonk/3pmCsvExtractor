package threepm;

import java.io.File;

public class Main {
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        /*JFileChooser fileChooser = new JFileChooser();
        int returnVal = fileChooser.showOpenDialog(null);
        if (returnVal == JFileChooser.APPROVE_OPTION) {
            new ExcelReader(
                    fileChooser.getSelectedFile().getAbsoluteFile()).getRowAsListFromExcel();
        }*/
        //new ExcelReader(new File("C:\\Users\\HISProgrammer\\Downloads\\3PM Import Files Assistance\\input_excel_file\\HTS_Optimization_for_3PM.xlsx")).getRowAsListFromExcel();
        
        ExcelReader xx = new ExcelReader(new File("C:\\Users\\HISProgrammer\\Downloads\\3PM Import Files Assistance\\input_excel_file\\HTS_Optimization_for_3PM.xlsx"));
        xx.setReportingPeriod("201809");
        xx.getRowAsListFromExcel();

    }

}
