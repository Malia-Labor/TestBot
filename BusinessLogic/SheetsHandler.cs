using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Sheets.v4;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Linq;
using static Google.Apis.Sheets.v4.SpreadsheetsResource;
using Newtonsoft.Json;

namespace BusinessLogic
{
    public class SheetsHandler : ISheetsHandler
    {
        private static readonly string[] _scopes = { SheetsService.Scope.Spreadsheets };
        private const string _configFileName = "config.json";
        private string _spreadsheetId;
        private SpreadsheetsResource.ValuesResource _valuesResource;
        private SheetsService _sheetsService;

        public SheetsHandler()
        {
            using (var stream = new FileStream(_configFileName, FileMode.Open, FileAccess.Read))
            {
                _sheetsService = new SheetsService(new BaseClientService.Initializer
                {
                    HttpClientInitializer = GoogleCredential.FromStream(stream).CreateScoped(_scopes)
                });
                _valuesResource = _sheetsService.Spreadsheets.Values;
                _spreadsheetId = JsonConvert.DeserializeObject<dynamic>(File.ReadAllText(_configFileName)).SpreadsheetId;
            }
        }

        /// <summary>
        /// Writes to multiple cells in Google Sheets.
        /// </summary>
        /// <param name="rangeLow">Low end of range to write to</param>
        /// <param name="rangeHigh">High end of range to write to</param>
        /// <param name="insertStrings">List of data to enter.</param>
        /// <returns>Returns true if write was successful.</returns>
        public async Task<bool> WriteAsync(string rangeLow, string rangeHigh, IList<IList<object>> insertStrings)
        {
            try
            {
                // Add data to be entered into ValueRange
                ValueRange values = new ValueRange { Values = insertStrings };
                string writeRange = $"{rangeLow}:{rangeHigh}";
                // Create update request
                var update = _valuesResource.Update(values, _spreadsheetId, writeRange);
                update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                // Send request
                var response = await update.ExecuteAsync();
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }

        /// <summary>
        /// Appends data to Google Sheets
        /// </summary>
        /// <param name="values">Data to enter into sheets</param>
        /// <param name="range">Range to append into</param>
        /// <returns>Returns true if append was successful.</returns>
        public async Task<bool> AppendAsync(ValueRange values, string range)
        {
            try
            {
                // Create request to append
                var request = _valuesResource.Append(values, _spreadsheetId, range);
                // Set options for the request
                SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum valueInputOption = (SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum)1;
                SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum insertDataOption = (SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum)1;
                request.ValueInputOption = valueInputOption;
                request.InsertDataOption = insertDataOption;
                // Send request
                var response = await request.ExecuteAsync();
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }

        /// <summary>
        /// Reads Google Sheets and returns data from the sheet.
        /// </summary>
        /// <param name="readRange">Range of columns to read.</param>
        /// <returns>List of data from the sheets.</returns>
        public async Task<IList<IList<object>>> ReadAsync(string readRange)
        {
            try
            {
                // Send request to get data from the sheets
                var response = await _valuesResource.Get(_spreadsheetId, readRange).ExecuteAsync();
                IList<IList<object>> values = response.Values;
                return values;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }

        /// <summary>
        /// Deletes data in specific cells
        /// </summary>
        /// <param name="cellsToDelete">List of cells to delete data from.</param>
        /// <returns>Returns true if delete was successful.</returns>
        public async Task<bool> DeleteAsync(List<string> cellsToDelete)
        {
            try
            {
                // Create new request
                BatchClearValuesRequest requestBody = new BatchClearValuesRequest();
                requestBody.Ranges = cellsToDelete;
                SpreadsheetsResource.ValuesResource.BatchClearRequest request = _valuesResource.BatchClear(requestBody, _spreadsheetId);
                // Send request
                var response = await request.ExecuteAsync();
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }

        /// <summary>
        /// Creates a new sheet with a given name.
        /// </summary>
        /// <param name="sheetName">Name of new sheet.</param>
        /// <returns></returns>
        public async Task<bool> NewSheetAsync(string sheetName)
        {
            try
            {
                // List of requests to send with batch update
                List<Request> requests = new List<Request>();
                // Create request for a new sheet
                AddSheetRequest newSheet = new AddSheetRequest { Properties = new SheetProperties { Title = sheetName } };
                Request newRequest = new Request { AddSheet = newSheet };
                // Add to requests list
                requests.Add(newRequest);
                // Create batch update request
                BatchUpdateSpreadsheetRequest requestBody = new BatchUpdateSpreadsheetRequest();
                requestBody.Requests = requests;
                SpreadsheetsResource.BatchUpdateRequest request = _sheetsService.Spreadsheets.BatchUpdate(requestBody, _spreadsheetId);
                // Send request
                BatchUpdateSpreadsheetResponse response = await request.ExecuteAsync();
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }

        public async Task<bool> DeleteSheetAsync(string sheetName)
        {
            try
            {
                // Create request to get the available sheets
                List<Request> requests = new List<Request>();
                GetRequest getSheetsRequest = _sheetsService.Spreadsheets.Get(_spreadsheetId);
                // Send request
                var getResponse = await getSheetsRequest.ExecuteAsync();
                // Find matching sheet
                var sheet = getResponse.Sheets.Where(x => x.Properties.Title.ToLower() == sheetName.ToLower()).FirstOrDefault();
                // Return if sheet does not exist
                if (sheet == null)
                    return false;
                // set sheet id
                int? sheetId = sheet.Properties.SheetId;
                // Create request to delete a sheet and add to requests list
                DeleteSheetRequest deleteSheet = new DeleteSheetRequest { SheetId = sheetId };
                Request newRequest = new Request { DeleteSheet = deleteSheet };
                requests.Add(newRequest);
                // Create request for updating
                BatchUpdateSpreadsheetRequest requestBody = new BatchUpdateSpreadsheetRequest();
                requestBody.Requests = requests;
                // Send request
                SpreadsheetsResource.BatchUpdateRequest request = _sheetsService.Spreadsheets.BatchUpdate(requestBody, _spreadsheetId);
                BatchUpdateSpreadsheetResponse response = await request.ExecuteAsync();
                return true;

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }
        public async Task<bool> CopyTemplateAsync(string sheetName)
        {
            try
            {
                // Create request to get the available sheets
                List<Request> requests = new List<Request>();
                GetRequest getSheetsRequest = _sheetsService.Spreadsheets.Get(_spreadsheetId);
                // Send request
                var getResponse = await getSheetsRequest.ExecuteAsync();
                // Find template sheet
                var sheet = getResponse.Sheets.Where(x => x.Properties.Title == "template").FirstOrDefault();
                // No template found, return false
                if (sheet == null)
                    return false;
                // set template sheet id
                int? sheetId = sheet.Properties.SheetId;
                // Create request to copy template and add to requests list
                DuplicateSheetRequest copyRequest = new DuplicateSheetRequest { SourceSheetId = sheetId, NewSheetName = sheetName };
                Request newRequest = new Request { DuplicateSheet = copyRequest };
                requests.Add(newRequest);
                // Create request for batch update
                BatchUpdateSpreadsheetRequest requestBody = new BatchUpdateSpreadsheetRequest();
                requestBody.Requests = requests;
                SpreadsheetsResource.BatchUpdateRequest request = _sheetsService.Spreadsheets.BatchUpdate(requestBody, _spreadsheetId);
                // Send request
                BatchUpdateSpreadsheetResponse response = await request.ExecuteAsync();
                return true;

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }

        public async Task<List<string>> SheetsList()
        {
            try
            {
                List<string> sheets = new List<string>();
                // Create request to get the available sheets
                List<Request> requests = new List<Request>();
                GetRequest getSheetsRequest = _sheetsService.Spreadsheets.Get(_spreadsheetId);
                // Send request
                var getResponse = await getSheetsRequest.ExecuteAsync();
                // Add sheet names to list
                foreach (var sheet in getResponse.Sheets)
                {
                    if (sheet.Properties.Title.ToLower() != "template")
                        sheets.Add(sheet.Properties.Title);
                }
                return sheets;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }
    }
}
