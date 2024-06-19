using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project_Demo_Excel
{
    public class Header
    {
        public void CreateHeader(ExcelWorksheet worksheet, int totalColumns)
        {
            int midColumn = totalColumns / 2;
            string midColumnName = GetExcelColumnName(midColumn);
            string endColumnName = GetExcelColumnName(totalColumns + 1);
            int startColumnPart2 = midColumn + 4;
            string startColumnPart2Name = GetExcelColumnName(startColumnPart2);
            
            for (int i = 1; i <= totalColumns; i++)
            {
                worksheet.Column(i).Width = 20;
            }
            // Set row heights
            worksheet.Row(2).Height = 20;
            worksheet.Row(3).Height = 20;
            worksheet.Row(4).Height = 20;
            // Header title 1
            worksheet.Cells[$"B2:{midColumnName}2"].Merge = true;
            worksheet.Cells["B2"].Value = "SỞ NÔNG NGHIỆP VÀ PTNT TỈNH LẠNG SƠN";
            worksheet.Cells[$"B3:{midColumnName}3"].Merge = true;
            worksheet.Cells["B3"].Value = "CHI CỤC THUỶ LỢI";
            worksheet.Cells["B3"].Style.Font.Bold = true;
            worksheet.Cells[$"B4:{midColumnName}4"].Merge = true;
            worksheet.Cells["B4"].Value = "------------------------------";
            // Header title 2
            worksheet.Cells[$"{startColumnPart2Name}2:{endColumnName}2"].Merge = true;
            worksheet.Cells[$"{startColumnPart2Name}2"].Value = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM";
            worksheet.Cells[$"{startColumnPart2Name}3:{endColumnName}3"].Merge = true;
            worksheet.Cells[$"{startColumnPart2Name}3"].Value = "Độc lập - Tự do - Hạnh phúc";
            worksheet.Cells[$"{startColumnPart2Name}2:{endColumnName}3"].Style.Font.Bold = true;
            worksheet.Cells[$"{startColumnPart2Name}2:{endColumnName}3"].Style.Font.Italic = true;
            worksheet.Cells[$"{startColumnPart2Name}4:{endColumnName}4"].Merge = true;
            worksheet.Cells[$"{startColumnPart2Name}4"].Value = "------------------------------";
            // Fomat
            worksheet.Cells[$"B2:{endColumnName}3"].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
            worksheet.Cells[$"B2:{endColumnName}4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[$"B4:{endColumnName}4"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            worksheet.Cells[$"{startColumnPart2Name}2:{endColumnName}3"].Style.Font.Bold = true;
            worksheet.Cells[$"B2:{endColumnName}3"].Style.Font.Size = 13;
            worksheet.Cells.Style.Font.Name = "\"Times New Roman\"";
            worksheet.View.ShowGridLines = false;

            ApplyBordersToAllCells(worksheet);
        }

            
        public static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }

        private void ApplyBordersToAllCells(ExcelWorksheet worksheet)
        {
            var lastRow = worksheet.Dimension.End.Row;
            var lastColumn = GetExcelColumnName(worksheet.Dimension.End.Column);
            var fullRange = $"B9:{lastColumn}{lastRow}";
            var cellRange = worksheet.Cells[fullRange];

            cellRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            cellRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            cellRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            cellRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
        }
    }
}
