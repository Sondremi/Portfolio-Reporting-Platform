package report;

import java.io.FileWriter;
import java.io.IOException;

final class ReportScriptHelper {

    private ReportScriptHelper() {}

    private static final String DETAILS_TOGGLE_SCRIPT_1 =
            "function toggleOverviewDetails(rowId, button) {\n"
            + "  var row = document.getElementById(rowId);\n"
            + "  if (!row) return;\n"
            + "  var isOpen = row.style.display === 'table-row';\n"
            + "  row.style.display = isOpen ? 'none' : 'table-row';\n"
            + "  if (button) button.textContent = isOpen ? 'Show' : 'Hide';\n"
            + "}\n"
            + "function toggleDetailGroup(groupName, button) {\n"
            + "  var rows = document.querySelectorAll('tr.details-row[data-group=\\\"' + groupName + '\\\"]');\n"
            + "  if (!rows.length) return;\n"
            + "  window.__detailGroupNextAction = window.__detailGroupNextAction || {};\n"
            + "  var action = window.__detailGroupNextAction[groupName] || 'open';\n"
            + "  var open = action === 'open';\n"
            + "  rows.forEach(function(row) {\n"
            + "    row.style.display = open ? 'table-row' : 'none';\n"
            + "    var rowId = row.id;\n"
            + "    if (!rowId) return;\n"
            + "    var rowButton = document.querySelector('button.expand-btn[data-target=\\\"' + rowId + '\\\"]');\n"
            + "    if (rowButton) rowButton.textContent = open ? 'Hide' : 'Show';\n"
            + "  });\n"
            + "  window.__detailGroupNextAction[groupName] = open ? 'close' : 'open';\n"
            + "  if (button) {\n"
            + "    var label = button.getAttribute('data-detail-label');\n"
            + "    if (label) button.textContent = label + ' ' + (open ? '▾' : '▸');\n"
            + "    else button.textContent = open ? '▾' : '▸';\n"
            + "  }\n"
            + "}\n";

    static void writeDetailsToggleScript(FileWriter writer) throws IOException {
        writer.write(DETAILS_TOGGLE_SCRIPT_1);
    }
}
