using Google.Apis.Sheets.v4.Data;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace BusinessLogic
{
    public interface ISheetsHandler
    {
        Task<bool> AppendAsync(ValueRange values, string range);
        Task<bool> CopyTemplateAsync(string sheetName);
        Task<bool> DeleteAsync(List<string> cellsToDelete);
        Task<bool> DeleteSheetAsync(string sheetName);
        Task<bool> NewSheetAsync(string sheetName);
        Task<IList<IList<object>>> ReadAsync(string readRange);
        Task<List<string>> SheetsList();
        Task<bool> WriteAsync(string rangeLow, string rangeHigh, IList<IList<object>> insertStrings);
    }
}