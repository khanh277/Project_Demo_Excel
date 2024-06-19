using Microsoft.EntityFrameworkCore;
using Microsoft.Win32;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Project_Demo_Excel.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using static Project_Demo_Excel.Models.DatabasePrjContext;

namespace Project_Demo_Excel
{
    public partial class MainWindow : Window
    {
        private readonly DatabasePrjContext context;

        public MainWindow()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var optionsBuilder = new DbContextOptionsBuilder<DatabasePrjContext>();
            context = new DatabasePrjContext(optionsBuilder.Options);
        }

        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            var data = context.GetDapTranSummary();
            if (data != null && data.Any())
            {
                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel files (*.xlsx)|*.xlsx",
                    DefaultExt = "xlsx",
                    FileName = "ExportedData.xlsx"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    try
                    {
                        using (var package = new ExcelPackage(new FileInfo(saveFileDialog.FileName)))
                        {
                            var worksheet1 = package.Workbook.Worksheets.Add("DapTranSummary");
                            ExportDataToExcel(data, worksheet1);
                            CreateHeaders(worksheet1, "DapTranSummary Headers", new TemplateDapTran());
                            CalculateAndInsertColumnTotals(worksheet1);

                            var worksheet2 = package.Workbook.Worksheets.Add("OtherSummary");
                            ExportDataToExcel(data, worksheet2);
                            CreateHeaders(worksheet2, "OtherSummary Headers", new TemplateDapTran());
                            CalculateAndInsertColumnTotals(worksheet2);

                            package.Save();
                        }

                        MessageBox.Show("Data exported successfully to " + saveFileDialog.FileName, "Export Successful", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Failed to export data: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
            else
            {
                MessageBox.Show("No data fetched from the database.", "No Data", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void ExportDataToExcel(IEnumerable<DapTranSummary> data, ExcelWorksheet worksheet)
        {
            worksheet.Cells["B11"].LoadFromCollection(data, PrintHeaders: false);
            FormatWorksheet(worksheet);
        }

        private void CreateHeaders(ExcelWorksheet worksheet, string headerText, TemplateDapTran template)
        {
            var header = new Header();
            int totalColumns = worksheet.Dimension.End.Column;

            header.CreateHeader(worksheet, totalColumns);
            template.CreateHeaderTable1(worksheet);
        }

        private void FormatWorksheet(ExcelWorksheet worksheet)
        {
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
            worksheet.Cells["B9:R10"].Style.Font.Bold = true;
        }

        private void CalculateAndInsertColumnTotals(ExcelWorksheet worksheet)
        {
            int startRow = 11;
            int endRow = worksheet.Dimension.End.Row;
            int startColumn = 2; 
            int endColumn = worksheet.Dimension.End.Column;
            int totalRow = endRow + 1;

            worksheet.Cells[totalRow, startColumn].Value = "Tổng";
            worksheet.Cells[totalRow, startColumn, totalRow, startColumn + 1].Merge = true;
            worksheet.Cells[totalRow, startColumn].Style.Font.Bold = true;
            worksheet.Cells[totalRow, startColumn].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[totalRow, startColumn].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            for (int col = startColumn + 2; col <= endColumn - 2; col++)
            {
                worksheet.Cells[totalRow, col].Formula = $"=SUM({worksheet.Cells[startRow, col].Address}:{worksheet.Cells[endRow, col].Address})";
                worksheet.Cells[totalRow, col].Style.Font.Bold = true;
                worksheet.Cells[totalRow, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                if (col == startColumn + 2)
                {
                    worksheet.Cells[totalRow, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                }
                else
                {
                    worksheet.Cells[totalRow, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }
            }
            worksheet.Cells[totalRow, endColumn - 2].Formula = $"=SUM({worksheet.Cells[startRow, endColumn - 2].Address}:{worksheet.Cells[endRow, endColumn - 2].Address})";
            worksheet.Cells[totalRow, endColumn - 2].Style.Font.Bold = true;
            worksheet.Cells[totalRow, endColumn - 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

            worksheet.Cells[totalRow, endColumn - 1].Formula = $"=SUM({worksheet.Cells[startRow, endColumn - 1].Address}:{worksheet.Cells[endRow, endColumn - 1].Address})";
            worksheet.Cells[totalRow, endColumn - 1].Style.Font.Bold = true;
            worksheet.Cells[totalRow, endColumn - 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

            worksheet.Cells[totalRow, endColumn].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[totalRow, endColumn].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[totalRow, endColumn].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[totalRow, endColumn].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            using (var range = worksheet.Cells[totalRow, startColumn, totalRow, endColumn])
            {
                range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
            worksheet.Calculate();
        }
    }
}
