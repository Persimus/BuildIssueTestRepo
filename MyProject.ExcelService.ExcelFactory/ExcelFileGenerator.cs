using System;
using System.Data;
using System.IO;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;

namespace MyProject.ExcelService.ExcelFactory
{
    public class ExcelFileGenerator : IExcelFileGenerator
    {
        private readonly ILogger<ExcelFileGenerator> _log;

        public ExcelFileGenerator(ILogger<ExcelFileGenerator> log)
        {
            _log = log;
        }

        public Stream GetExcelDocument(DataSet dataSet)
        {
            try
            {
                using (var doc = new ExcelPackage())
                {
                    foreach (DataTable dataTable in dataSet.Tables)
                    {
                        AddSheet(doc, dataTable);
                    }

                    var memoryStream = new MemoryStream();
                    doc.SaveAs(memoryStream);
                    memoryStream.Position = 0;
                    return memoryStream;
                }
            }
            catch (Exception exception)
            {
                _log.LogError(exception, "Error trying to generate excel file");
                throw;
            }
        }

        private void AddSheet(ExcelPackage document, DataTable dataTable)
        {
            document.Workbook.Worksheets.Add(dataTable.TableName);
            var excelWorksheet = document.Workbook.Worksheets[^1];

            for (var columnIndex = 0; columnIndex < dataTable.Columns.Count; columnIndex++)
            {
                var dataColumn = dataTable.Columns[columnIndex];

                excelWorksheet.SetValue(1, columnIndex + 1, dataColumn.ColumnName);
                excelWorksheet.Column(columnIndex + 1).AutoFit(); // AutoFit uses System.Drawing.Common at runtime

                var rowIndex = 1;

                foreach (DataRow row in dataTable.Rows)
                {
                    excelWorksheet.SetValue(rowIndex + 1, columnIndex + 1, row.ItemArray[columnIndex]);
                    rowIndex++;
                }
            }

            excelWorksheet.Cells[1, 1, 1, dataTable.Columns.Count].AutoFilter = true;
        }
    }
}
