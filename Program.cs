using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace UploadInviteewithAnswersClient
{
    class Program
    {
        static string _token = "";
        static string TokenPath = "";
        static string clientid = "";
        static string clientsecret = "";
        static string uploadUrl = "";
        static async Task Main(string[] args)
        {
            Console.WriteLine("Please type your environment:       (staging / pilot / pilot2)");
            string env = Console.ReadLine().ToLower();
            switch (env)
            {
                case "staging":
                    Console.WriteLine("staging is selected");
                    TokenPath = "https://staging-api.alternacx.com/token";
                    clientid = "e7f8ae6d-92ac-4443-8f6d-1210b4dd21c3";
                    clientsecret = "ce4718e080d155ae7c3e7fbeb2793198";
                    uploadUrl = "https://staging-api.alternacx.com/api/upload/inviteesWithAnswers";
                    break;
                case "pilot":
                    Console.WriteLine("pilot is selected");
                    TokenPath = "https://pilot-api.alternacx.com/token";
                    clientid = "fa750f45-07d3-4549-b54f-45abda5491d6";
                    clientsecret = "60d620d3b6e1078a5ad00e914ac7c523";
                    uploadUrl = "https://pilot-api.alternacx.com/api/upload/inviteesWithAnswers";

                    break;
                case "pilot2":
                    Console.WriteLine("pilot2 is selected");
                    TokenPath = "https://pilot2-api.alternacx.com/token";
                    clientid = "121d6c4a-fced-481f-8a9d-164f6094c0bb";
                    clientsecret = "76238ea9bc4af9bd8db0640b5d0d5d58";
                    uploadUrl = "https://pilot2-api.alternacx.com/api/upload/inviteesWithAnswers";
                    break;
                default:
                    Console.WriteLine("Wrong environment, staging is selected as default environment");
                    env = "staging";
                    TokenPath = "https://staging-api.alternacx.com/token";
                    clientid = "e7f8ae6d-92ac-4443-8f6d-1210b4dd21c3";
                    clientsecret = "ce4718e080d155ae7c3e7fbeb2793198";
                    uploadUrl = "https://staging-api.alternacx.com/api/upload/inviteesWithAnswers";
                    break;
            }

            if (string.IsNullOrEmpty(_token))
            {
                GetToken();
            }
            Random num = new ();

            List<Invitee> invitees = new List<Invitee>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var file = new FileInfo(fileName: @"C:\Users\Mustafa\Documents\answer1.xlsx");
            using var package = new ExcelPackage(file);
            List<Invitee> peopleFromExcel = await LoadExcelFile(file);

            foreach (var p in peopleFromExcel)
            {
                int number = num.Next(1000, 100000);
                Invitee invitee = new ()
                {
                    InviteeId = env + "_poc_" + number,
                    InviteeMSISDN = "90570" + number,
                    InviteeFullName = env + "_poc_" + number,
                    InviteeEmail = env + "_poc_" + number + "@alternatest.com",
                    InviteeLanguage = p.InviteeLanguage,
                    InviteeLocation = p.InviteeLocation,
                    TransactionChannel = p.TransactionChannel,
                    TransactionType = p.TransactionType,
                    TransactionDate = p.TransactionDate,
                    InteractionChannel = p.InteractionChannel,
                    CustomData1 = p.CustomData1,
                    CustomData2 = p.CustomData2,
                    CustomData3 = p.CustomData3,
                    CustomData4 = p.CustomData4,
                    CustomData5 = p.CustomData5,
                    CustomData6 = p.CustomData6,
                    CustomData7 = p.CustomData7,
                    CustomData8 = p.CustomData8,
                    CustomData9 = p.CustomData9,
                    CustomData10 = p.CustomData10,
                    InviteeSegments = new List<InviteeSegment>(p.InviteeSegments),
                    InviteeAnswers = new List<InviteeAnswer>()
                    {
                        new InviteeAnswer
                        {
                          QuestionIntegrationCode  = p.InviteeAnswers.Select(i=>i.QuestionIntegrationCode).FirstOrDefault(),
                          Value = p.InviteeAnswers.Select(i=>i.Value).FirstOrDefault(),
                          AnsweredDate = p.InviteeAnswers.Select(i=>i.AnsweredDate).FirstOrDefault()
                        },
                        new InviteeAnswer
                        {
                          QuestionIntegrationCode  = p.InviteeAnswers.Select(i=>i.QuestionIntegrationCode).Last(),
                          Value = p.InviteeAnswers.Select(i=>i.Value).Last(),
                          AnsweredDate = p.InviteeAnswers.Select(i=>i.AnsweredDate).Last()
                        }
                    }
                };

                invitees.Add(invitee);
            }
            if (invitees.Count > 0)
            {
                double iterationCount = Math.Ceiling(invitees.Count / 200.0);
                for (int i = 0; i < (int)iterationCount; i++)
                {
                    try
                    {
                        var x = invitees.Skip(i * 200).Take(200).ToList();
                        using (var client = new HttpClient())
                        {
                            var json = JsonConvert.SerializeObject(invitees).ToString();
                            var data = JsonConvert.DeserializeObject(json).ToString();
                            var url = uploadUrl;
                            client.DefaultRequestHeaders.Add("Authorization", "Bearer " + _token);
                            var response = await client.PostAsync(url, new StringContent(data, Encoding.UTF8, "application/json"));
                            var result = response.Content.ReadAsStringAsync().Result;
                            Console.WriteLine(response);
                            Console.WriteLine(result);
                            Console.WriteLine(data);
                            Console.ReadLine();
                        };
                    }
                    catch (Exception)
                    {
                        throw;
                    }

                }
            }
            else { Console.WriteLine("kayıt yok"); }

        }
        private static async Task<List<Invitee>> LoadExcelFile(FileInfo file)
        {
            List<Invitee> output = new();
            using var package = new ExcelPackage(file);
            await package.LoadAsync(file);
            var ws = package.Workbook.Worksheets[PositionID: 0];
            int row = 2;
            int segInd, col, txc, txt, inl, txd, inc, sg1, sg2, qnp, vl1, qcm, vl2, asd, lct, cs1, cs2, cs3, cs4, cs5, cs6, cs7, cs8, cs9, cs10;
            col = 1; segInd = 0;
            txc = txt = inl = txd = inc = sg1 = sg2 = qnp = vl1 = qcm = vl2 = asd = lct = cs1 = cs2 = cs3 = cs4 = cs5 = cs6 = cs7 = cs8 = cs9 = cs10 = 0;
            List<object> AllSegments = new List<object>();

            List<InviteeSegment> tempSegments = new List<InviteeSegment>();

            List<SegmentColumn> AllOf = new List<SegmentColumn>();
            SegmentColumn segmentList = new SegmentColumn();
            List<InviteeSegment> newSegments = new List<InviteeSegment>();
            var aa = ws.Cells;

            try
            {
                int rowCount = aa.Worksheet.Columns.EndColumn;
                for (int col1 = 1; col1 <= rowCount; col1++)
                {
                    if (aa[1, col1].Value.ToString().Contains("TransactionChannel")) { txc = col1; }
                    if (aa[1, col1].Value.ToString().Contains("TransactionType")) { txt = col1; }
                    if (aa[1, col1].Value.ToString().Contains("InviteeLanguage")) { inl = col1; }
                    if (aa[1, col1].Value.ToString().Contains("TransactionDate")) { txd = col1; }
                    if (aa[1, col1].Value.ToString().Contains("InteractionChannel")) { inc = col1; }
                    if (aa[1, col1].Value.ToString().Contains("InviteeSegment"))
                    {
                        string rowValue = aa[1, col1].Value.ToString().Remove(0, 15).ToString();
                        InviteeSegment invSegment1 = new InviteeSegment(rowValue, col1.ToString());
                        segmentList = new SegmentColumn(col1, rowValue);
                        AllOf.Add(segmentList);
                        segInd++;

                    }

                    if (aa[1, col1].Value.ToString().Contains("Question_NPS")) { qnp = col1; }
                    if (aa[1, col1].Value.ToString().Contains("Question1Value")) { vl1 = col1; }
                    if (aa[1, col1].Value.ToString().Contains("Question_Comment")) { qcm = col1; }
                    if (aa[1, col1].Value.ToString().Contains("Question2Value")) { vl2 = col1; }
                    if (aa[1, col1].Value.ToString().Contains("AnsweredDate")) { asd = col1; }
                    if (aa[1, col1].Value.ToString().Contains("Location")) { lct = col1; }
                    if (aa[1, col1].Value.ToString()=="CustomData1") { cs1 = col1; }
                    if (aa[1, col1].Value.ToString() == "CustomData2") { cs2 = col1; }
                    if (aa[1, col1].Value.ToString() == "CustomData3") { cs3 = col1; }
                    if (aa[1, col1].Value.ToString() == "CustomData4") { cs4 = col1; }
                    if (aa[1, col1].Value.ToString() == "CustomData5") { cs5 = col1; }
                    if (aa[1, col1].Value.ToString() == "CustomData6") { cs6 = col1; }
                    if (aa[1, col1].Value.ToString() == "CustomData7") { cs7 = col1; }
                    if (aa[1, col1].Value.ToString() == "CustomData8") { cs8 = col1; }
                    if (aa[1, col1].Value.ToString() == "CustomData9") { cs9 = col1; }
                    if (aa[1, col1].Value.ToString()=="CustomData10") { cs10 = col1; }
                }

                while (string.IsNullOrWhiteSpace(aa[row, col].Value?.ToString()) == false)
                {
                    Invitee p = new();
                    //p.InviteeId = aa[row, col + 1].Value.ToString();
                    // p.InviteeEmail = aa[row, col + 2].Value.ToString();
                    // p.InviteeMSISDN = aa[row, col + 3].Value.ToString();
                    // p.InviteeFullName = aa[row, col + 4].Value.ToString();
                    p.InviteeLanguage = inl > 0 && aa[row, inl].Value != null ? aa[row, inl].Value.ToString() : "tr";    
                    p.TransactionChannel = txc > 0 && aa[row, txc].Value != null ? aa[row, txc].Value?.ToString() : "";
                    p.TransactionType = txt > 0 && aa[row, txt].Value != null ? aa[row, txt].Value?.ToString() : "";
                    p.TransactionDate = txd > 0 && aa[row, txd].Value != null ? aa[row, txd].Value.ToString() : DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss");
                    p.InteractionChannel = inc > 0 && aa[row, inc].Value != null ? aa[row, inc].Value.ToString() : "8";

                    if (AllOf.Count > 0)
                    {
                        {
                            foreach (var item in AllOf) //createNewColumnHeadClass
                            {
                                if (aa[row, item.ColumnId].Value!=null)
                                {
                                    try
                                    {
                                        InviteeSegment newsegments = new InviteeSegment(
                                        item.ColumnTitle, aa[row, item.ColumnId].Value.ToString());
                                        newSegments.Add(newsegments);
                                    }
                                    catch (Exception)
                                    {
                                        continue;
                                    }

                                }

                            };
                        };
                    }
                    p.InviteeSegments=new List<InviteeSegment>();

                    foreach (var newsegment in newSegments)
                    {
                        int i = 0;

                        try
                        {
                            p.InviteeSegments.Add(newsegment);
                        }
                        catch (Exception)
                        {

                            continue;
                        }
                        i++;
                    }

                    p.InviteeAnswers = new()
                    {
                        new()
                        {
                            QuestionIntegrationCode = qnp > 0 && aa[row, qnp].Value != null ? aa[row, qnp].Value.ToString() : "STGNPS",
                            Value = vl1 > 0 && aa[row, vl1].Value != null ? aa[row, vl1].Value.ToString() : "10",
                            AnsweredDate = asd > 0 && aa[row, asd].Value != null ? aa[row, asd].Value.ToString() : DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss")
                        },
                        new()
                        {
                            QuestionIntegrationCode = qcm > 0 && aa[row, qcm].Value != null ? aa[row, qcm].Value.ToString() : "STGCOMMENT",
                            Value = vl2 > 0 && aa[row, vl2].Value != null ? aa[row, vl2].Value?.ToString() : "",
                            AnsweredDate = asd > 0 && aa[row, asd].Value != null ? aa[row, asd].Value.ToString() : DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss")
                        }
                    };
                    p.InviteeLocation = lct > 0 && aa[row, lct].Value != null ? aa[row, lct].Value.ToString() : "";
                    p.CustomData1 = cs1 > 0 && aa[row, cs1].Value != null ? aa[row, cs1].Value?.ToString() : "";
                    p.CustomData2 = cs2 > 0 && aa[row, cs2].Value != null ? aa[row, cs2].Value?.ToString() : "";
                    p.CustomData3 = cs3 > 0 && aa[row, cs3].Value != null ?  aa[row, cs3].Value.ToString() : "";
                    p.CustomData4 = cs4 > 0 && aa[row, cs4].Value != null  ? aa[row, cs4].Value?.ToString() : "";
                    p.CustomData5 = cs5 > 0 && aa[row, cs5].Value != null ? aa[row, cs5].Value?.ToString() : "";
                    p.CustomData6 = cs6 > 0 && aa[row, cs6].Value != null ? aa[row, cs6].Value?.ToString() : "";
                    p.CustomData7 = cs7 > 0 && aa[row, cs7].Value != null ? aa[row, cs7].Value?.ToString() : "";
                    p.CustomData8 = cs8 > 0 && aa[row, cs8].Value != null ? aa[row, cs8].Value?.ToString() : "";
                    p.CustomData9 = cs9 > 0 && aa[row, cs9].Value != null ? aa[row, cs9].Value?.ToString() : "";
                    p.CustomData10 = cs10 > 0 && aa[row, cs10].Value != null ? aa[row, cs10].Value?.ToString() : "";

                    output.Add(p);
                    newSegments.Clear();
                    row++;

                }
            }
            catch (NullReferenceException e)
            {
                throw e;
            }
            return output;
        }

        private static void GetToken()
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, TokenPath)
            {
                Content = new FormUrlEncodedContent(new Dictionary<string, string>
                {
                    {"client_id", clientid},
                    {"client_secret", clientsecret},
                    {"grant_type", "client_credentials"}
                })
            };
            HttpClient _client = new HttpClient();
            HttpResponseMessage response = _client.SendAsync(request).Result;
            response.EnsureSuccessStatusCode();
            JObject payload = JObject.Parse(response.Content.ReadAsStringAsync().Result);
            _token = payload.Value<string>("access_token");
        }
    }
}
