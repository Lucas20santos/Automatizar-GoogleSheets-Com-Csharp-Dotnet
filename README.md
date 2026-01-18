# Automatização de Planilhas do Google Sheets com C# (.NET)

Projeto demonstrando **integração profissional** entre **C# (.NET)** e a **API do Google Sheets** usando **Service Account**, com foco em automação, leitura de dados e boas práticas de segurança.

---

## 🚀 Funcionalidades

* Autenticação segura via **Service Account**
* Leitura de intervalos do Google Sheets
* Estrutura preparada para escrita e CRUD
* Credenciais protegidas com `.gitignore`

---

## 🔐 Configuração de Credenciais (Service Account)

> ⚠️ **Nunca versionar credenciais**

1. Crie uma **Service Account** no Google Cloud Console
2. Gere e baixe o arquivo JSON de credenciais
3. Salve em: `Credentials/credentials.json`
4. Compartilhe a planilha com o e-mail da Service Account (permissão **Editor**)

Exemplo de e-mail:

```md
sheets-automation@seu-projeto.iam.gserviceaccount.com
```

---

## 🛠️ Criação do Projeto

```bash
dotnet new console -n GoogleSheetsAutomation
cd GoogleSheetsAutomation
```

### 📦 Dependências

```bash
dotnet add package Google.Apis.Sheets.v4
dotnet add package Google.Apis.Auth
```

---

## 📁 Estrutura Recomendada

```text
GoogleSheetsAutomation/
 ├── Credentials/
 │    └── credentials.json        # 🔒 segredo (ignorado pelo git)
 ├── Services/
 │    └── GoogleSheetsService.cs  # camada de serviço
 ├── Program.cs
 ├── .gitignore
 └── README.md
```

No `.gitignore`:

```gitignore
Credentials/*.json
```

---

## 🧩 Implementação

### 📌 Service — `GoogleSheetsService.cs`

Classe responsável por encapsular o acesso à API do Google Sheets.

```csharp
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace GoogleSheetsAutomation.Services
{
    public class GoogleSheetsService
    {
        private readonly SheetsService _service;

        public GoogleSheetsService(string credentialPath)
        {
            var credential = CredentialFactory
                .FromFile<ServiceAccountCredential>(credentialPath)
                .ToGoogleCredential()
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

---

### ▶️ Program — `Program.cs`

```csharp
using System;
using GoogleSheetsAutomation.Services;

class Program
{
    static void Main()
    {
        string credentialPath = "Credentials/credentials.json";

        // 🔎 ID da planilha (NÃO versionar se for sensível)
        string spreadsheetId = "SEU_SPREADSHEET_ID";

        // ⚠️ Use o nome EXATO da aba
        string range = "Registro de Vendas!A1:C10";

        var sheetsService = new GoogleSheetsService(credentialPath);
        var values = sheetsService.LerIntervalo(spreadsheetId, range);

        if (values.Count == 0)
        {
            Console.WriteLine("⚠️ Nenhum dado encontrado.");
            return;
        }

        Console.WriteLine("📄 Dados da planilha:\n");
        foreach (var row in values)
        {
            Console.WriteLine(string.Join(" | ", row));
        }
    }
}
```

---

## 🧠 Boas Práticas Adotadas

* ❌ Credenciais fora do repositório
* ✅ Uso de Service Account
* ✅ Separação de responsabilidades (Service / Program)
* ✅ Código pronto para evolução (write, update, CRUD)

---

## 📌 Próximos Passos

* [ ] Inserir dados (AppendRow)
* [ ] Atualizar células
* [ ] Criar CRUD simples
* [ ] Transformar em API ASP.NET

---

## 👤 Autor

**Lucas de Souza Santos**
Engenharia de Controle e Automação | Backend C# (.NET)
