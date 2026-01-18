using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace GoogleSheetsAutomation.Services
{
    public class GoogleSheetsAutomation
    {
        private readonly SheetsService _service;
        public GoogleSheetsAutomation()
        {
            var credetial = GoogleCredential
            .GetApplicationDefault()
            .CreateScoped(SheetsService.Scope.Spreadsheets);

            _service = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credetial,
                ApplicationName = "Google Sheets Automation C#"
            });
        }

        public IList<IList<object>> LerIntervalo(string spreadsheetId, string range)
        {
            var request = _service.Spreadsheets.Values.Get(spreadsheetId, range);
            ValueRange response = request.Execute();

            return response.Values ?? new List<IList<object>>();
        }
    }
}
