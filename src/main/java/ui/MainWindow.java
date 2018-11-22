package ui;

import java.awt.Color;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.File;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JTextField;
import javax.swing.filechooser.FileFilter;

import com.jgoodies.forms.layout.ColumnSpec;
import com.jgoodies.forms.layout.FormLayout;
import com.jgoodies.forms.layout.FormSpecs;
import com.jgoodies.forms.layout.RowSpec;

import threepm.ExcelReader;


public class MainWindow {
    
    private JFrame parentFrame;
    private JLabel lblInformation = new JLabel("");
    private File inputFile;
    private JTextField txtReportingPeriod;
    
    /**
     * Launch the application.
     */
    public static void main(String[] args) {
        EventQueue.invokeLater(new Runnable() {
            
            public void run() {
                try {
                    MainWindow window = new MainWindow();
                    window.parentFrame.setVisible(true);
                }
                catch (Exception e) {
                    e.printStackTrace();
                }
            }
        });
    }
    
    /**
     * Create the application.
     */
    public MainWindow() {
        initialize();
    }
    
    /**
     * Initialize the contents of the frame.
     */
    private void initialize() {
        parentFrame = new JFrame();
        parentFrame.setResizable(false);
        parentFrame.setTitle("CSV Data generator for 3PM");
        parentFrame.setBounds(100, 100, 643, 300);
        parentFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        parentFrame.getContentPane().setLayout(new FormLayout(new ColumnSpec[] {
                FormSpecs.RELATED_GAP_COLSPEC,
                FormSpecs.DEFAULT_COLSPEC,
                FormSpecs.RELATED_GAP_COLSPEC,
                FormSpecs.DEFAULT_COLSPEC,
                FormSpecs.RELATED_GAP_COLSPEC,
                ColumnSpec.decode("max(49dlu;default)"),
                FormSpecs.RELATED_GAP_COLSPEC,
                ColumnSpec.decode("35dlu:grow"),},
            new RowSpec[] {
                FormSpecs.RELATED_GAP_ROWSPEC,
                FormSpecs.DEFAULT_ROWSPEC,
                FormSpecs.RELATED_GAP_ROWSPEC,
                FormSpecs.DEFAULT_ROWSPEC,
                FormSpecs.RELATED_GAP_ROWSPEC,
                FormSpecs.DEFAULT_ROWSPEC,
                FormSpecs.RELATED_GAP_ROWSPEC,
                FormSpecs.DEFAULT_ROWSPEC,
                FormSpecs.RELATED_GAP_ROWSPEC,
                FormSpecs.DEFAULT_ROWSPEC,}));
        
        JLabel lblChooseFile = new JLabel("Choose file");
        parentFrame.getContentPane().add(lblChooseFile, "4, 2");
        
        JButton btnSelectFile = new JButton("...");
        btnSelectFile.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
                JFileChooser fileChooser = new JFileChooser();
                fileChooser.setDialogTitle("Specify a file to save");
                fileChooser.setFileFilter(new FileFilter() {

                    public String getDescription() {
                        return "XLSX (*.xlsx)";
                    }

                    public boolean accept(File f) {
                        if (f.isDirectory()) {
                            return true;
                        } else {
                            String filename = f.getName().toLowerCase();
                            return filename.endsWith(".xlsx");
                        }
                    }
                 });                 
                int userSelection = fileChooser.showSaveDialog(parentFrame);
                 
                if (userSelection == JFileChooser.APPROVE_OPTION) {
                    inputFile = fileChooser.getSelectedFile();
                    System.out.println("Save as file: " + inputFile.getAbsolutePath());
                    lblInformation.setText(inputFile.getName());
                    
                }

            }
        });
        btnSelectFile.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
            }
        });
        parentFrame.getContentPane().add(btnSelectFile, "6, 2");
        
        JLabel lblCompatbleExcelTemplate = new JLabel("(Compatible excel template)");
        lblCompatbleExcelTemplate.setFont(new Font("Tahoma", Font.ITALIC, 10));
        parentFrame.getContentPane().add(lblCompatbleExcelTemplate, "4, 4");
        
        JLabel lblEnterPeriodyyyymm = new JLabel("Enter reporting period(YYYYMM)");
        parentFrame.getContentPane().add(lblEnterPeriodyyyymm, "4, 6, right, default");
        
        txtReportingPeriod = new JTextField();
        txtReportingPeriod.setText("201809");
        parentFrame.getContentPane().add(txtReportingPeriod, "6, 6");
        txtReportingPeriod.setColumns(10);
        lblInformation.setForeground(Color.RED);
        lblInformation.setFont(new Font("Tahoma", Font.BOLD | Font.ITALIC, 12));
        
        parentFrame.getContentPane().add(lblInformation, "2, 8, 7, 1, fill, fill");
        
        JButton btnGenerateCsv = new JButton("Generate CSV(s)");
        btnGenerateCsv.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
                lblInformation.setText("Processing....");  

                String reportingPeriod =  txtReportingPeriod.getText();
                ExcelReader fileReader = new ExcelReader(inputFile);
                fileReader.setReportingPeriod(reportingPeriod);
                String retMessage  = fileReader.getRowAsListFromExcel();                    

                lblInformation.setText("<html>"+ retMessage + "</html>");  
            }
        });
        btnGenerateCsv.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
            }
        });
        parentFrame.getContentPane().add(btnGenerateCsv, "6, 10");
        

    }
    
}
