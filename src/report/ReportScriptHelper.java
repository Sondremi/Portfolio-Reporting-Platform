package report;

import java.io.FileWriter;
import java.io.IOException;

final class ReportScriptHelper {

    private ReportScriptHelper() {}

    static void writeDetailsToggleScript(FileWriter writer) throws IOException {
        writer.write("function toggleOverviewDetails(rowId, button) {\n");
        writer.write("  var row = document.getElementById(rowId);\n");
        writer.write("  if (!row) return;\n");
        writer.write("  var isOpen = row.style.display === 'table-row';\n");
        writer.write("  row.style.display = isOpen ? 'none' : 'table-row';\n");
        writer.write("  if (button) button.textContent = isOpen ? 'Show' : 'Hide';\n");
        writer.write("}\n");
        writer.write("function toggleDetailGroup(groupName, button) {\n");
        writer.write("  var rows = document.querySelectorAll('tr.details-row[data-group=\\\"' + groupName + '\\\"]');\n");
        writer.write("  if (!rows.length) return;\n");
        writer.write("  window.__detailGroupNextAction = window.__detailGroupNextAction || {};\n");
        writer.write("  var action = window.__detailGroupNextAction[groupName] || 'open';\n");
        writer.write("  var open = action === 'open';\n");
        writer.write("  rows.forEach(function(row) {\n");
        writer.write("    row.style.display = open ? 'table-row' : 'none';\n");
        writer.write("    var rowId = row.id;\n");
        writer.write("    if (!rowId) return;\n");
        writer.write("    var rowButton = document.querySelector('button.expand-btn[data-target=\\\"' + rowId + '\\\"]');\n");
        writer.write("    if (rowButton) rowButton.textContent = open ? 'Hide' : 'Show';\n");
        writer.write("  });\n");
        writer.write("  window.__detailGroupNextAction[groupName] = open ? 'close' : 'open';\n");
        writer.write("  if (button) button.textContent = open ? '▾' : '▸';\n");
        writer.write("}\n");
    }
}
