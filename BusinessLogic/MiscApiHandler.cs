using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace BusinessLogic
{
    public class MiscApiHandler
    {
        private string _catUri = "https://api.thecatapi.com/v1/images/search";
        public async Task<string> RandomCat()
        {
            string resultUrl = "";
            try
            {
                var client = new RestClient(_catUri);
                client.Timeout = -1;
                var request = new RestRequest(Method.GET);
                IRestResponse response = await client.ExecuteAsync(request);
                resultUrl = JsonConvert.DeserializeObject<dynamic>(response.Content.Replace('[', '\0').Replace(']', '\0')).url;
                return resultUrl;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return resultUrl;
            }
        }
    }
}
