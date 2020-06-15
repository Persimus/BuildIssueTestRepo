using System;
using System.Data;
using Microsoft.AspNetCore.Mvc;
using MyProject.ExcelService.ExcelFactory;

namespace MyProject.ExcelService.API.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ExcelController : ControllerBase
    {
        private readonly IExcelFileGenerator _excelFileGenerator;

        public ExcelController(IExcelFileGenerator excelFileGenerator)
        {
            _excelFileGenerator = excelFileGenerator;
        }

        [HttpGet]
        public IActionResult Get()
        {
            var fileStream = _excelFileGenerator.GetExcelDocument(CreateDataSet());

            var result = new FileStreamResult(fileStream, "application/xslx")
            {
                FileDownloadName = "Excel File.xlsx"
            };

            return result;
        }

        private DataSet CreateDataSet()
        {
            var dataSet = new DataSet();
            dataSet.Tables.Add(CreateDataTable("Table"));
            dataSet.Tables.Add(CreateDataTable("Another Table"));

            return dataSet;
        }


        private DataTable CreateDataTable(string name)
        {
            var dataTable = new DataTable(name);

            for (var i = 0; i < 10; i++)
            {
                var columnName = $"{name} {i}";
                dataTable.Columns.Add(columnName);
            }

            return dataTable;
        }
    }
}
