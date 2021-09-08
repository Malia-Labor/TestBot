using System.Threading.Tasks;

namespace BusinessLogic
{
    public interface ISignupHandler
    {
        Task<bool> AddToSignup(string sheetName, string username, string userID, string value);
        Task<bool> RemoveFromSignup(string sheetName, string userID);
    }
}