using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessLogic
{
    public class SignupHandler : SheetsHandler, ISignupHandler
    {
        private string _appendRange = "A2:C2";
        private string _rangeLow = "A";
        private string _rangeHigh = "C";

        /// <summary>
        /// Adds a user to a specific Google Sheet. Overwrites existing row if user exists already.
        /// </summary>
        /// <param name="sheetName">Name of sheet to add to.</param>
        /// <param name="username">Name of the user using the command.</param>
        /// <param name="userID">Id of the user using the command</param>
        /// <param name="value">Data to enter with username and ID (used for extra data like roles)</param>
        /// <returns>True if successful. False if an error occurred in adding user.</returns>
        public async Task<bool> AddToSignup(string sheetName, string username, string userID, string value)
        {
            // Get data from sheet and check if user exists already
            var values = await ReadAsync($"{sheetName}!{_rangeLow}:{_rangeHigh}");
            if (values == null) // read is empty. Sheet does not exist
            {
                return false;
            }
            if (values.Any(x => x.Any(x => x.ToString() == userID)) == true) // user exists in sheet
            {
                // get location of user in signup
                int row = values.IndexOf(values.Where(x => x.Any(x => x.ToString() == userID)).FirstOrDefault()) + 1;
                // overwrite data
                return await WriteAsync($"{_rangeLow}{row}", $"{_rangeHigh}{row}", new List<IList<object>> { new List<object> { userID, username, value } });
            }
            else // user not in sheet
            {
                return await AppendAsync(new ValueRange { Values = new List<IList<object>> { new List<object> { userID, username, value } } }, _appendRange);
            }
        }

        /// <summary>
        /// Removes a user from a Google Sheet
        /// </summary>
        /// <param name="sheetName">Name of sheet to delete from</param>
        /// <param name="userID">ID of user to check for</param>
        /// <returns>False if an error occured in delete attempt.</returns>
        public async Task<bool> RemoveFromSignup(string sheetName, string userID)
        {
            // Get data from sheets and check if user exists
            var values = await ReadAsync($"{sheetName}!{_rangeLow}:{_rangeHigh}");
            if (values == null)
                return false;
            if (values.Any(x => x.Any(x => x.ToString() == userID)) == true) // user exists in sheet
            {
                // get location of user in signup
                int row = values.IndexOf(values.Where(x => x.Any(x => x.ToString() == userID)).FirstOrDefault()) + 1;
                // overwrite data
                return await DeleteAsync(new List<string> { $"A{row}", $"B{row}", $"C{row}" }); // cells columns are hardcoded, maybe a way around this?
            }
            else // user not in sheet
            {
                return true;
            }
        }
    }
}
