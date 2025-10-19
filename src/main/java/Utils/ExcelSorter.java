package Utils;

import Model.ExcelRow;
import BLL.Rule;

import java.util.*;

public class ExcelSorter {

    public static List<ExcelRow> applyRules(List<ExcelRow> rows, List<Rule> rules) {
        if (rows == null || rows.isEmpty()) return Collections.emptyList();
        if (rules == null || rules.isEmpty()) return new ArrayList<>(rows);

        List<ExcelRow> result = new ArrayList<>(rows);

        // build a combined comparator safely
        Comparator<ExcelRow> combinedComparator = null;

        for (Rule rule : rules) {
            Comparator<ExcelRow> comp = (a, b) -> {
                Object valA = a.get(rule.fieldName());
                Object valB = b.get(rule.fieldName());

                // handle nulls first
                if (valA == null && valB == null) return 0;
                if (valA == null) return rule.ascending() ? -1 : 1;
                if (valB == null) return rule.ascending() ? 1 : -1;

                // try to compare Comparable types
                if (valA instanceof Comparable && valB instanceof Comparable) {
                    try {
                        int cmp = ((Comparable) valA).compareTo(valB);
                        return rule.ascending() ? cmp : -cmp;
                    } catch (ClassCastException e) {
                        // fallback if incompatible types
                        return valA.toString().compareToIgnoreCase(valB.toString()) * (rule.ascending() ? 1 : -1);
                    }
                }

                // fallback to string comparison
                return valA.toString().compareToIgnoreCase(valB.toString()) * (rule.ascending() ? 1 : -1);
            };

            combinedComparator = (combinedComparator == null)
                    ? comp
                    : combinedComparator.thenComparing(comp);
        }

        if (combinedComparator != null) {
            result.sort(combinedComparator);
        }

        return result;
    }

}
