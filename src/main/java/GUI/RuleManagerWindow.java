package GUI;

import Utils.ExcelSorter;
import BLL.Rule;
import DataAccess.ExcelRowDAO;
import Model.ExcelRow;

import javax.swing.*;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.prefs.Preferences;

public class RuleManagerWindow extends JFrame {

    private final List<ExcelRow> data;
    private final List<Rule> rules = new ArrayList<>();

    private final DefaultTableModel ruleTableModel; //the data!!!
    private final JTable ruleTable; //the table itself!!

    private final JComboBox<String> fieldCombo;
    private final JComboBox<String> orderCombo;

    private List<String> availableFields;

    public RuleManagerWindow(List<ExcelRow> data, List<String> fieldNames) {
        this.data = data;
        this.availableFields = new ArrayList<>(fieldNames);
        this.availableFields.sort(String::compareToIgnoreCase);


        setTitle("Rule Manager");
        setSize(700, 400);
        setLocationRelativeTo(null);
        setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
        setLayout(new BorderLayout());

        // left -> added rules table
        String[] columns = {"#", "Field", "Order"};
        ruleTableModel = new DefaultTableModel(columns, 0) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return false; // make table cells not editable
            }
        };

        ruleTable = new JTable(ruleTableModel);
        ruleTable.setRowHeight(30);
        ruleTable.setFillsViewportHeight(true);
        ruleTable.getColumnModel().getColumn(0).setPreferredWidth(30);
        ruleTable.getColumnModel().getColumn(1).setPreferredWidth(200);
        ruleTable.getColumnModel().getColumn(2).setPreferredWidth(70);

        DefaultTableCellRenderer centerRenderer = new DefaultTableCellRenderer();
        centerRenderer.setHorizontalAlignment(SwingConstants.CENTER);
        ruleTable.getColumnModel().getColumn(0).setCellRenderer(centerRenderer);
        ruleTable.getColumnModel().getColumn(1).setCellRenderer(centerRenderer);
        ruleTable.getColumnModel().getColumn(2).setCellRenderer(centerRenderer);

        JScrollPane scrollPane = new JScrollPane(ruleTable);
        scrollPane.setPreferredSize(new Dimension(350, 400));

        // right -> add rules
        JPanel rulePanel = new JPanel();
        rulePanel.setLayout(new BoxLayout(rulePanel, BoxLayout.Y_AXIS));

        fieldCombo = new JComboBox<>(availableFields.toArray(new String[0]));
        orderCombo = new JComboBox<>(new String[]{"Ascending", "Descending"});
        fieldCombo.setMaximumSize(new Dimension(Integer.MAX_VALUE, fieldCombo.getPreferredSize().height));
        orderCombo.setMaximumSize(new Dimension(Integer.MAX_VALUE, orderCombo.getPreferredSize().height));

        JButton addRuleButton = new JButton("Add Rule");
        addRuleButton.addActionListener(e -> addRule());

        JButton editRuleButton = new JButton("Edit Selected Rule");
        editRuleButton.addActionListener(e -> editSelectedRule());

        JButton applyButton = new JButton("Apply Rules");
        applyButton.addActionListener(e -> applyRules());

        rulePanel.add(new JLabel("Field:"));
        rulePanel.add(fieldCombo);
        rulePanel.add(new JLabel("Order:"));
        rulePanel.add(orderCombo);
        rulePanel.add(Box.createVerticalStrut(10));
        rulePanel.add(addRuleButton);
        rulePanel.add(Box.createVerticalStrut(20));
        rulePanel.add(editRuleButton);
        rulePanel.add(Box.createVerticalStrut(20));
        rulePanel.add(applyButton);

        add(scrollPane, BorderLayout.WEST);
        add(rulePanel, BorderLayout.CENTER);
    }

    private void addRule() {
        String selectedField = (String) fieldCombo.getSelectedItem();
        if (selectedField == null) {
            JOptionPane.showMessageDialog(this, "No field selected.");
            return;
        }

        boolean ascending = "Ascending".equals(orderCombo.getSelectedItem());
        Rule newRule = new Rule(selectedField, ascending);

        rules.add(newRule);

        ruleTableModel.addRow(new Object[]{rules.size(), selectedField, ascending ? "Ascending" : "Descending"});

        availableFields.remove(selectedField);
        refreshFieldCombo();

    }


    private void editSelectedRule() {
        int selectedRow = ruleTable.getSelectedRow();
        if (selectedRow < 0) {
            JOptionPane.showMessageDialog(this, "Please select a rule to edit.", "No Rule Selected", JOptionPane.WARNING_MESSAGE);
            return;
        }

        Rule selectedRule = rules.get(selectedRow);
        String oldField = selectedRule.fieldName();

        // dropdown: all available fields plus the current field
        List<String> comboFields = new ArrayList<>(availableFields);
        if (!comboFields.contains(oldField)) comboFields.add(oldField);

        JComboBox<String> tempFieldCombo = new JComboBox<>(comboFields.toArray(new String[0]));
        tempFieldCombo.setSelectedItem(oldField);

        JComboBox<String> tempOrderCombo = new JComboBox<>(new String[]{"Ascending", "Descending"});
        tempOrderCombo.setSelectedItem(selectedRule.ascending() ? "Ascending" : "Descending");

        int option = JOptionPane.showConfirmDialog(
                this,
                new Object[]{"Select Field:", tempFieldCombo, "Order:", tempOrderCombo},
                "Edit Rule",
                JOptionPane.OK_CANCEL_OPTION
        );

        if (option == JOptionPane.OK_OPTION) {
            String newField = (String) tempFieldCombo.getSelectedItem();
            boolean ascending = "Ascending".equals(tempOrderCombo.getSelectedItem());

            if (!oldField.equals(newField)) {
                availableFields.add(oldField);
                availableFields.remove(newField);
                refreshFieldCombo();
            }

            // update the rule in the list
            rules.set(selectedRow, new Rule(newField, ascending));

            refreshTable();
        }
    }


    private void applyRules() {
        List<ExcelRow> processed = ExcelSorter.applyRules(data, rules);
        JOptionPane.showMessageDialog(this, "Applied rules. Rows: " + processed.size());

        // use Preferences to remember last save directory
        Preferences prefs = Preferences.userNodeForPackage(getClass());
        String lastDir = prefs.get("lastSaveDir", null);

        JFileChooser saveChooser = (lastDir != null)
                ? new JFileChooser(new File(lastDir))
                : new JFileChooser();

        saveChooser.setDialogTitle("Save Sorted Excel Output");
        saveChooser.setSelectedFile(new File("sorted_output.xlsx"));

        int saveResult = saveChooser.showSaveDialog(this);
        if (saveResult == JFileChooser.APPROVE_OPTION) {
            File saveFile = saveChooser.getSelectedFile();
            prefs.put("lastSaveDir", saveFile.getParent()); // remember for next time

            String savePath = saveFile.getAbsolutePath();
            if (!savePath.toLowerCase().endsWith(".xlsx")) {
                savePath += ".xlsx";
            }

            // write Excel
            new ExcelRowDAO<>(ExcelRow.class).writeToExcel(processed, savePath);
            JOptionPane.showMessageDialog(this, "File saved to: " + savePath);
        }
    }


    private void refreshTable() {
        ruleTableModel.setRowCount(0); // clear existing rows

        for (int i = 0; i < rules.size(); i++) {
            Rule r = rules.get(i);
            ruleTableModel.addRow(new Object[]{
                    i + 1, // keep numbering consistent
                    r.fieldName(),
                    r.ascending() ? "Ascending" : "Descending"
            });
        }
    }


    private void refreshFieldCombo() {
        fieldCombo.removeAllItems();
        availableFields.sort(String::compareToIgnoreCase);
        for (String field : availableFields) {
            fieldCombo.addItem(field);
        }
        fieldCombo.setEnabled(!availableFields.isEmpty());
    }


}
