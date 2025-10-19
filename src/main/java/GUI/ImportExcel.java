package GUI;

import DataAccess.ExcelRowDAO;
import Model.ExcelRow;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.FlowLayout;
import java.io.File;
import java.util.*;
import java.util.prefs.Preferences;

public class ImportExcel extends JFrame {

    private JButton selectFileButton;

    public ImportExcel() {
        setTitle("Excel Importer");
        setSize(400, 150);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLocationRelativeTo(null);
        setLayout(new FlowLayout());

        selectFileButton = new JButton("Select Excel File");

        selectFileButton.addActionListener(e -> openFileChooser());

        add(selectFileButton);
    }

    private void openFileChooser() {
        Preferences prefs = Preferences.userNodeForPackage(getClass());
        String lastDir = prefs.get("lastDir", null);

        JFileChooser fileChooser = (lastDir != null)
                ? new JFileChooser(new File(lastDir))
                : new JFileChooser();


        // filter for .xls and .xlsx files
        FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel Files", "xls", "xlsx");
        fileChooser.setFileFilter(filter);

        int result = fileChooser.showOpenDialog(this);

        if (result == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile();
            prefs.put("lastDir", selectedFile.getParent()); // save, so it opens here next time

            String filePath = selectedFile.getAbsolutePath();
            System.out.println("Selected file: " + filePath);

            // load the Excel data -> delegates readFromExcel method to read and parse the headers
            List<ExcelRow> rows = new ExcelRowDAO<>(ExcelRow.class).readFromExcel(filePath);
            // print fields and values
//            for (ExcelRow row : rows) {
//                System.out.println("Row:");
//                for (String fieldName : row.getFieldNames()) {
//                    System.out.println("  " + fieldName + ": " + row.get(fieldName));
//                }
//                System.out.println();
//            }

            List<String> fieldNames = rows.getFirst().getFieldNames(); // assuming headers exist

            new RuleManagerWindow(rows, new ArrayList<>(fieldNames)).setVisible(true);

        }
    }

    public static void main (String[]args){
        SwingUtilities.invokeLater(() -> {
            ImportExcel gui = new ImportExcel();
            gui.setVisible(true);
        });
    }

}


