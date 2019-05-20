using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Xceed.Words.NET;
using System.Diagnostics;
using System.IO;
using System.Threading;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Errend.Core;
using Errend.Core.CTO;
using System.Globalization;
using System.Data.SqlServerCe;

namespace Errend
{
    class GoogleXLS
    {
        string[] Scopes = { SheetsService.Scope.Drive };
        string ApplicationName = "Google Sheets API .NET Quickstart";
        public ValueRange ReadDataFromGoogleXML(string firstPoint, string secondPont)
        {
            UserCredential credential;
            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            String spreadsheetId = Properties.Settings.Default["ContainerTable"].ToString();
            String range = "ПРИБЫТИЕ/ЭКСПОРТ!D" + firstPoint + ":" + "W" + secondPont;
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);
            ValueRange response = request.Execute();
            string[] contArr = new string[response.Values.Count];
            for (int i = 0; i < response.Values.Count; i++)
            {
                if (response.Values[i][2].ToString() == "")
                {
                    if (i - 1 >= 0)
                    {
                        response.Values[i][2] = response.Values[i - 1][2];
                    }
                }
            }
            for (int i = 0; i < response.Values.Count; i++)
            {
                contArr[i] = response.Values[i][0].ToString();
            }
            Array.Sort(contArr, StringComparer.InvariantCulture);
            ValueRange result = request.Execute();

            for (int i = 0; i < contArr.Length; i++)
            {
                for (int j = 0; j < contArr.Length; j++)
                {
                    if (response.Values[j][0].ToString() == contArr[i])
                    {
                        result.Values[i] = response.Values[j];
                    }
                }
            }
            return result;
        }

        public ValueRange ReadDataFromXMLErrend(string numberErrend)
        {
            UserCredential credential;
            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            String spreadsheetId = Properties.Settings.Default["ErrentTable"].ToString();
            String range = "ОМТП - эксп.!A" + numberErrend + ":" + "I" + numberErrend;
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);
            ValueRange response = request.Execute();
            return response;
        }
    }
}
