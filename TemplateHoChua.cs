using OfficeOpenXml.Style;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

namespace Project_Demo_Excel
{
    internal class TemplateHoChua
    {
        public void CreateHeaderTable2(ExcelWorksheet worksheet)
        {
            worksheet.Cells["B6:S6"].Merge = true;
            worksheet.Cells["B6"].Value = "Số Lượng Hồ Chứa Hiện Có";
            worksheet.Cells["B6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["B6"].Style.Font.Bold = true;
            worksheet.Cells["B7:S7"].Merge = true;
            worksheet.Cells["B7"].Value = "Năm 2023 <Năm xây dựng>";
            worksheet.Cells["B7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Cells["B7"].Style.Font.Bold = true;
            // Merge cells for main headers
            worksheet.Cells["B9:B10"].Merge = true;
            worksheet.Cells["B9"].Value = "STT";
            worksheet.Cells["C9:C10"].Merge = true;
            worksheet.Cells["C9"].Value = "Huyện, Thành phố";
            worksheet.Cells["D9:D10"].Merge = true;
            worksheet.Cells["D9"].Value = "Tổng diện tích lưu vực (km2)";
            worksheet.Cells["E9:H9"].Merge = true;
            worksheet.Cells["E9"].Value = "Đơn vị quản lý";
            worksheet.Cells["I9:I10"].Merge = true;
            worksheet.Cells["I9"].Value = "Tổng dung tích toàn bộ (triệu/m3)";
            worksheet.Cells["J9:J10"].Merge = true;
            worksheet.Cells["J9"].Value = "Tổng dung tích hữu ích (triệu/m3)";
            worksheet.Cells["K9:K10"].Merge = true;
            worksheet.Cells["K9"].Value = "Tổng hồ chứa (cái)";
            worksheet.Cells["L9:O9"].Merge = true;
            worksheet.Cells["L9"].Value = "Chia ra:";
            worksheet.Cells["P9:Q9"].Merge = true;
            worksheet.Cells["Q9"].Value = "Năng lực tưới (ha)";
            worksheet.Cells["R9:R10"].Merge = true;
            worksheet.Cells["R9"].Value = "Ghi chú";

            // Sub-Headers
            worksheet.Cells["E10"].Value = "UBND huyện, thành phố";
            worksheet.Cells["F10"].Value = "UBND xã";
            worksheet.Cells["G10"].Value = "Công ty TNHH MTV Khai thác công trình thủy lợi";
            worksheet.Cells["H10"].Value = "Khác";
            worksheet.Cells["L10"].Value = "Hồ chứa quan trọng đặc biệt (cái)";
            worksheet.Cells["M10"].Value = "Hồ chứa lớn (cái)";
            worksheet.Cells["N10"].Value = "Hồ chứa vừa (cái)";
            worksheet.Cells["O10"].Value = "Hồ chứa nhỏ (cái)";
            worksheet.Cells["P10"].Value = "Kế hoạch";
            worksheet.Cells["Q10"].Value = "Nghiệm thu";

            worksheet.Cells["B9:R10"].Style.Font.Bold = true;
            worksheet.Cells["B9:R10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["B9:R10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["B9:R10"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells["B9:R10"].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            worksheet.Cells["B9:R10"].Style.WrapText = true;

            AdjustColumnWidths2(worksheet);

            CenterData(worksheet);
        }

        private void AdjustColumnWidths2(ExcelWorksheet worksheet)
        {
            int totalColumns = worksheet.Dimension.End.Column;

            worksheet.Column(1).Width = 2.4;
            worksheet.Column(2).Width = 8.2;
            worksheet.Column(3).Width = 20;
            for (int i = 4; i <= totalColumns; i++)
            {
                worksheet.Column(i).Width = 11.86;
            }
            worksheet.Row(9).Height = 30;
            worksheet.Row(10).Height = 90;
        }

        private void CenterData(ExcelWorksheet worksheet)
        {
            int lastRow = worksheet.Dimension.End.Row;
            string dataRange = $"E11:P{lastRow}";
            worksheet.Cells[dataRange].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[dataRange].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        }
    }
}

