# Automatizando de planilhas do google planilhas com dotnet csharp

## Primeiro passo: Criando projeto

### Criando Projeto

```bash
dotnet new console -n GoogleSheetsAutomation
cd GoogleSheetsAutomation
```

### Instalando dependencias

```bash
dotnet add package Google.Apis.Sheets.v4
dotnet add package Google.Apis.Auth
```

### Estrutura recomendada

```md
GoogleSheetsAutomation/
 ├── Credentials/
 │    └── credentials.json   🔒 segredo
 ├── Services/
 │    └── GoogleSheetsService.cs
 ├── Models/
 ├── Program.cs
 ├── appsettings.json
 └── .gitignore
```

## Criando as credenciais

Será feito depois

## Criação da aplicação

### GoogleSheetsService.cs

#### V1

```cs
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace GoogleSheetsAutomation.Services
{
    public class GoogleSheetsService
    {
        private readonly SheetsService _service;

        public GoogleSheetsService()
        {
            GoogleCredential credential;

            using (var stream = new FileStream("Credentials/credentials.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(SheetsService.Scope.Spreadsheets);
            }

            _service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "Google Sheets Automation C#",
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
```

#### V2

```cs
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace GoogleSheetsAutomation.Services
{
    public class GoogleSheetsService
    {
        private readonly SheetsService _service;

        public GoogleSheetsService()
        {
            var json = File.ReadAllText("Credentials/credentials.json");

            var credential = GoogleCredential
                .FromJson(json)
                .CreateScoped(SheetsService.Scope.Spreadsheets);

            _service = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
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
```
