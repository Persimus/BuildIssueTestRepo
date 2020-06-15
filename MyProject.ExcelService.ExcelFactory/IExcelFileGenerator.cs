using System.Data;
using System.IO;

namespace MyProject.ExcelService.ExcelFactory
{
    public interface IExcelFileGenerator
    {
        Stream GetExcelDocument(DataSet dataSet);
    }
}