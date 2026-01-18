using System;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

class Program
{
    static void Main(string[] args)
    {
        // 1️⃣ Caminho do arquivo JSON da Service Account
        string credentialPath = "Credentials/credentials.json";

        // 2️⃣ Criação da credencial (FORMA ATUAL, NÃO OBSOLETA)
        GoogleCredential googleCredential =
            CredentialFactory
                .FromFile<ServiceAccountCredential>(credentialPath)
                .ToGoogleCredential()
                .CreateScoped(SheetsService.Scope.Spreadsheets);

        Console.WriteLine("✅ Autenticação realizada com sucesso!");

        // 3️⃣ Criação do serviço do Google Sheets
        var sheetsService = new SheetsService(new BaseClientService.Initializer
        {
            HttpClientInitializer = googleCredential,
            ApplicationName = "Minha Aplicacao Sheets"
        });

        Console.WriteLine("📡 Serviço do Google Sheets pronto para uso.");

        // 4️⃣ ID da planilha (da sua URL)
        string spreadsheetId = "10atKV1BR67qI5e-0QsnVP4L9Dirm15nHt6Kj3LUlPCM";

        // ⚠️ TROQUE pelo nome EXATO da aba
        string range = "Registro de Vendas!A1:C10";

        // 5️⃣ Cria a requisição de leitura
        var request = sheetsService.Spreadsheets.Values.Get(spreadsheetId, range);

        // 6️⃣ Executa a leitura
        ValueRange response = request.Execute();

        var values = response.Values;

        // 7️⃣ Exibe os dados
        if (values == null || values.Count == 0)
        {
            Console.WriteLine("⚠️ Nenhum dado encontrado.");
        }
        else
        {
            Console.WriteLine("\n📄 Dados da planilha:\n");

            foreach (var row in values)
            {
                Console.WriteLine(string.Join(" | ", row));
            }
        }
    }
}
