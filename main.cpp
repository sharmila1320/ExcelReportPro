#include "xlsxwriter.h"
#include <iostream>

int main() {
    lxw_workbook  *workbook  = workbook_new("Employee_Report.xlsx");
    lxw_worksheet *worksheet = workbook_add_worksheet(workbook, NULL);

    // Write headers
    worksheet_write_string(worksheet, 0, 0, "Name", NULL);
    worksheet_write_string(worksheet, 0, 1, "Task", NULL);
    worksheet_write_string(worksheet, 0, 2, "Date", NULL);
    worksheet_write_string(worksheet, 0, 3, "Score", NULL);

    // Dummy data entries
    worksheet_write_string(worksheet, 1, 0, "Alice", NULL);
    worksheet_write_string(worksheet, 1, 1, "Bug Fix", NULL);
    worksheet_write_string(worksheet, 1, 2, "2025-06-01", NULL);
    worksheet_write_number(worksheet, 1, 3, 92.5, NULL);

    worksheet_write_string(worksheet, 2, 0, "Bob", NULL);
    worksheet_write_string(worksheet, 2, 1, "New Feature", NULL);
    worksheet_write_string(worksheet, 2, 2, "2025-06-01", NULL);
    worksheet_write_number(worksheet, 2, 3, 88.0, NULL);

    worksheet_write_string(worksheet, 3, 0, "Clara", NULL);
    worksheet_write_string(worksheet, 3, 1, "Testing", NULL);
    worksheet_write_string(worksheet, 3, 2, "2025-06-01", NULL);
    worksheet_write_number(worksheet, 3, 3, 95.0, NULL);

    workbook_close(workbook);

    std::cout << "✅ Excel report 'Employee_Report.xlsx' generated successfully!\n";

    return 0;
}
