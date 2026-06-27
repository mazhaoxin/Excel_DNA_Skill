using System;
using System.Drawing;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

public class YearCalendarRibbon : ExcelRibbon
{
    private string selectedYear;
    private int selectedWeekStartIndex; // 0=周一, 1=周六, 2=周日

    public void RibbonLoad(IRibbonUI ribbon)
    {
        selectedYear = DateTime.Now.Year.ToString();
        selectedWeekStartIndex = 0;
    }

    // ===== ComboBox: Year Selector =====
    public int GetYearItemCount(IRibbonControl control)
    {
        return 21;
    }

    public string GetYearItemLabel(IRibbonControl control, int index)
    {
        int currentYear = DateTime.Now.Year;
        return (currentYear - 10 + index).ToString();
    }

    public string GetYearText(IRibbonControl control)
    {
        if (string.IsNullOrEmpty(selectedYear))
        {
            selectedYear = DateTime.Now.Year.ToString();
        }
        return selectedYear;
    }

    public void OnYearChanged(IRibbonControl control, string text)
    {
        selectedYear = text;
    }

    // ===== ComboBox: Week Start Selector =====
    public int GetWeekStartItemCount(IRibbonControl control)
    {
        return 3;
    }

    public string GetWeekStartItemLabel(IRibbonControl control, int index)
    {
        switch (index)
        {
            case 0: return "周一";
            case 1: return "周六";
            case 2: return "周日";
            default: return "";
        }
    }

    public string GetWeekStartText(IRibbonControl control)
    {
        switch (selectedWeekStartIndex)
        {
            case 0: return "周一";
            case 1: return "周六";
            case 2: return "周日";
            default: return "周一";
        }
    }

    public void OnWeekStartChanged(IRibbonControl control, string text)
    {
        if (text == "周一")
            selectedWeekStartIndex = 0;
        else if (text == "周六")
            selectedWeekStartIndex = 1;
        else if (text == "周日")
            selectedWeekStartIndex = 2;
    }

    // ===== Button: Generate Calendar =====
    public void GenerateCalendarClick(IRibbonControl control)
    {
        ExcelApp app = (ExcelApp)ExcelDnaUtil.Application;

        bool oldScreenUpdating = app.ScreenUpdating;
        app.ScreenUpdating = false;

        try
        {
            int year;
            if (!int.TryParse(selectedYear, out year))
            {
                MessageBox.Show("请先选择一个有效的年份。");
                return;
            }

            int weekStart;
            switch (selectedWeekStartIndex)
            {
                case 0: weekStart = 1; break;  // Monday
                case 1: weekStart = 6; break;  // Saturday
                case 2: weekStart = 0; break;  // Sunday
                default: weekStart = 1; break;
            }

            Workbook wb = (Workbook)app.ActiveWorkbook;
            if (wb == null)
            {
                MessageBox.Show("请先打开或创建一个Excel工作簿。");
                return;
            }

            string sheetName = year.ToString();

            // Delete existing sheet with same name (suppress alerts)
            bool oldDisplayAlerts = app.DisplayAlerts;
            app.DisplayAlerts = false;

            try
            {
                object nameObj = sheetName;
                Worksheet existingSheet = (Worksheet)wb.Sheets.get_Item(nameObj);
                existingSheet.Delete();
            }
            catch { }

            app.DisplayAlerts = oldDisplayAlerts;

            // Add new sheet at the end
            int sheetCount = wb.Sheets.Count;
            object afterObj = wb.Sheets[sheetCount];
            Worksheet ws = (Worksheet)wb.Sheets.Add(System.Reflection.Missing.Value, afterObj, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
            ws.Name = sheetName;

            GenerateCalendar(ws, year, weekStart);

            // MessageBox.Show(string.Format("已成功生成 {0} 年日历！", year));
        }
        catch (Exception ex)
        {
            MessageBox.Show("错误: " + ex.Message);
        }
        finally
        {
            app.ScreenUpdating = oldScreenUpdating;
        }
    }

    private void GenerateCalendar(Worksheet ws, int year, int weekStart)
    {
        string[] monthNames = { "一月", "二月", "三月", "四月", "五月", "六月",
                                "七月", "八月", "九月", "十月", "十一月", "十二月" };

        string[] weekDayNames;
        switch (weekStart)
        {
            case 1:
                weekDayNames = new string[] { "一", "二", "三", "四", "五", "六", "日" };
                break;
            case 6:
                weekDayNames = new string[] { "六", "日", "一", "二", "三", "四", "五" };
                break;
            case 0:
            default:
                weekDayNames = new string[] { "日", "一", "二", "三", "四", "五", "六" };
                break;
        }

        // Colors (OLE format: R + G*256 + B*65536)
        int titleBgColor = ColorToOle(Color.FromArgb(68, 114, 196));
        int titleFontColor = ColorToOle(Color.White);
        int headerBgColor = ColorToOle(Color.FromArgb(180, 200, 230));
        int headerFontColor = ColorToOle(Color.Black);
        int saturdayColor = ColorToOle(Color.Blue);
        int sundayColor = ColorToOle(Color.Red);
        int normalColor = ColorToOle(Color.Black);
        int whiteBg = ColorToOle(Color.White);

        // Set column widths
        for (int col = 1; col <= 23; col++)
        {
            ((Range)ws.Columns[col]).ColumnWidth = 5.5;
        }
        ((Range)ws.Columns[8]).ColumnWidth = 1.5;
        ((Range)ws.Columns[16]).ColumnWidth = 1.5;

        // Set default row height
        ws.Rows.RowHeight = 18;

        // Separator rows
        for (int sr = 9; sr <= 30; sr += 9)
        {
            ((Range)ws.Rows[sr]).RowHeight = 6;
        }

        // Generate each month
        for (int month = 0; month < 12; month++)
        {
            int gridRow = month / 3;
            int gridCol = month % 3;

            int baseRow = 1 + gridRow * 9;
            int baseCol = 1 + gridCol * 8;

            // --- Month Title ---
            Range titleRange = (Range)ws.Range[ws.Cells[baseRow, baseCol], ws.Cells[baseRow, baseCol + 6]];
            titleRange.MergeCells = true;
            titleRange.Value = monthNames[month];
            titleRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            titleRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            titleRange.Font.Bold = true;
            titleRange.Font.Size = 14;
            titleRange.Interior.Color = titleBgColor;
            titleRange.Font.Color = titleFontColor;
            titleRange.RowHeight = 26;
            titleRange.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            titleRange.Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThin;
            titleRange.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            titleRange.Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThin;
            titleRange.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            titleRange.Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;
            titleRange.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            titleRange.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;

            // --- Weekday Headers ---
            for (int w = 0; w < 7; w++)
            {
                Range headerCell = (Range)ws.Cells[baseRow + 1, baseCol + w];
                headerCell.Value = weekDayNames[w];
                headerCell.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                headerCell.VerticalAlignment = XlVAlign.xlVAlignCenter;
                headerCell.Interior.Color = headerBgColor;
                headerCell.Font.Color = headerFontColor;
                headerCell.Font.Bold = true;
                headerCell.Font.Size = 10;
                headerCell.Borders.LineStyle = XlLineStyle.xlContinuous;
                headerCell.Borders.Weight = XlBorderWeight.xlThin;
            }

            // --- Date Cells ---
            DateTime firstDay = new DateTime(year, month + 1, 1);
            int daysInMonth = DateTime.DaysInMonth(year, month + 1);

            int firstDayOffset;
            if (weekStart == 1)  // Monday start
            {
                firstDayOffset = ((int)firstDay.DayOfWeek + 6) % 7;
            }
            else if (weekStart == 6)  // Saturday start
            {
                firstDayOffset = ((int)firstDay.DayOfWeek + 1) % 7;
            }
            else  // Sunday start
            {
                firstDayOffset = (int)firstDay.DayOfWeek;
            }

            // Initialize all 42 date cells (6 rows x 7 cols)
            for (int r = 0; r < 6; r++)
            {
                for (int c = 0; c < 7; c++)
                {
                    Range cell = (Range)ws.Cells[baseRow + 2 + r, baseCol + c];
                    cell.Value = "";
                    cell.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    cell.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    cell.Font.Size = 10;
                    cell.Font.Color = normalColor;
                    cell.Interior.Color = whiteBg;
                    cell.Borders.LineStyle = XlLineStyle.xlContinuous;
                    cell.Borders.Weight = XlBorderWeight.xlThin;
                }
            }

            // Fill in actual dates
            for (int d = 1; d <= daysInMonth; d++)
            {
                int offset = firstDayOffset + d - 1;
                int r = offset / 7;
                int c = offset % 7;

                Range cell = (Range)ws.Cells[baseRow + 2 + r, baseCol + c];
                cell.Value = d;

                DateTime date = new DateTime(year, month + 1, d);
                if (date.DayOfWeek == DayOfWeek.Saturday)
                {
                    cell.Font.Color = saturdayColor;
                }
                else if (date.DayOfWeek == DayOfWeek.Sunday)
                {
                    cell.Font.Color = sundayColor;
                }
            }
        }
    }

    private int ColorToOle(Color color)
    {
        return color.R | (color.G << 8) | (color.B << 16);
    }
}
