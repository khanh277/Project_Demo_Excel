using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace Project_Demo_Excel
{
    internal class TemplateDapTran
    {
        public void CreateHeaderTable1(ExcelWorksheet worksheet)
        {
            worksheet.Cells["B6:S6"].Merge = true;
            worksheet.Cells["B6"].Value = "TỔNG CÔNG TRÌNH THỦY LỢI";
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
            worksheet.Cells["E9:E10"].Merge = true;
            worksheet.Cells["E9"].Value = "Tổng đập tràn (cái)";
            worksheet.Cells["F9:I9"].Merge = true;
            worksheet.Cells["F9"].Value = "Đơn vị quản lý";
            worksheet.Cells["J9:M9"].Merge = true;
            worksheet.Cells["J9"].Value = "Chia ra:";
            worksheet.Cells["N9:P9"].Merge = true;
            worksheet.Cells["N9"].Value = "Chiều dài đập";
            worksheet.Cells["Q9:R9"].Merge = true;
            worksheet.Cells["Q9"].Value = "Năng lực tưới (ha)";
            worksheet.Cells["S9:S10"].Merge = true;
            worksheet.Cells["S9"].Value = "Ghi chú";

            // Sub-Headers
            worksheet.Cells["F10"].Value = "UBND huyện, thành phố";
            worksheet.Cells["G10"].Value = "UBND xã";
            worksheet.Cells["H10"].Value = "Công ty TNHH MTV Khai thác công trình thủy lợi";
            worksheet.Cells["I10"].Value = "Khác";
            worksheet.Cells["J10"].Value = "Đập quan trọng đặc biệt (cái)";
            worksheet.Cells["K10"].Value = "Đập lớn (cái)";
            worksheet.Cells["L10"].Value = "Đập vừa (cái)";
            worksheet.Cells["M10"].Value = "Đập nhỏ (cái)";
            worksheet.Cells["N10"].Value = "Tổng số (km)";
            worksheet.Cells["O10"].Value = "Đã kiên cố (km)";
            worksheet.Cells["P10"].Value = "Mương đất (km)";
            worksheet.Cells["Q10"].Value = "Kế hoạch";
            worksheet.Cells["R10"].Value = "Nghiệm thu";

            worksheet.Cells["B9:S10"].Style.Font.Bold = true;
            worksheet.Cells["B9:S10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["B9:S10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["B9:S10"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells["B9:S10"].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            worksheet.Cells["B9:S10"].Style.WrapText = true;
            
            AdjustColumnWidths1(worksheet);

            CenterData(worksheet);
        }

        private void AdjustColumnWidths1(ExcelWorksheet worksheet)
        {
            int totalColumns = worksheet.Dimension.End.Column;

            worksheet.Column(1).Width = 2.4;
            worksheet.Column(2).Width = 8.2;
            worksheet.Column(3).Width = 20;
            for (int i = 4; i <= totalColumns; i++)
            {
                worksheet.Column(i).Width = 15;
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
