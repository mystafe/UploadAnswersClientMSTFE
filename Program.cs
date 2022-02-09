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

        static async Task Main(string[] args)
        {
            Console.WriteLine("Please type your environment:       (staging / pilot / pilot2)");
            string env = Console.ReadLine().ToLower();
            Environment environment=new(env);
            env=environment.environmentName;
            Console.WriteLine(env+" is selected.");
            if (string.IsNullOrEmpty(environment.bearer_token))
            {
                GetToken(environment);
            }
            List<Invitee> invitees = new List<Invitee>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var file = new FileInfo(fileName: @"C:\Users\Mustafa\Documents\answer1.xlsx");
            using var package = new ExcelPackage(file);
            List<Invitee> peopleFromExcel = await LoadExcelFile(file);

            foreach (Invitee p in peopleFromExcel)
            {
                Invitee invitee = new (p)
                {

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
                            var url = environment.uploadPath;
                            client.DefaultRequestHeaders.Add("Authorization", "Bearer " + _token);
                            var response = await client.PostAsync(url, new StringContent(data, Encoding.UTF8, "application/json"));
                            var result = response.Content.ReadAsStringAsync().Result;
                            Console.WriteLine((int)response.StatusCode);                                               
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

        }
        private static async Task<List<Invitee>> LoadExcelFile(FileInfo file)
        {
            List<Invitee> output = new();
            using var package = new ExcelPackage(file);
            await package.LoadAsync(file);
            var ws = package.Workbook.Worksheets[PositionID: 0];
            int row = 2;
            int col,iid,ims,ifn,iem, txc, txt, inl, txd, inc, sg1, sg2, qnp, vl1, qcm, vl2, asd, lct, cs1, cs2, cs3, cs4, cs5, cs6, cs7, cs8, cs9, cs10;
            col = 1;
            iid=ims=ifn=iem =txc = txt = inl = txd = inc = sg1 = sg2 = qnp = vl1 = qcm = vl2 = asd = lct = cs1 = cs2 = cs3 = cs4 = cs5 = cs6 = cs7 = cs8 = cs9 = cs10 = 0;
            List<object> AllSegments = new List<object>();

            List<InviteeSegment> tempSegments = new List<InviteeSegment>();

            List<SegmentColumn> AllOf = new List<SegmentColumn>();
            SegmentColumn segmentList = new SegmentColumn();
            List<InviteeSegment> newSegments = new List<InviteeSegment>();
            var aa = ws.Cells;
            try
            {
                int rowCount = aa.Worksheet.Columns.EndColumn;
                Random random=new();
                int number=random.Next(10000,1000000);
                for (int col1 = 1; col1 <= rowCount; col1++)
                {
                    if (aa[1, col1].Value.ToString().Trim().ToLower()=="inviteeid") { iid = col1; }
                    if ((aa[1, col1].Value.ToString().Trim().ToLower()=="inviteemsisdn")|| (aa[1, col1].Value.ToString().Trim().ToLower()=="msisdn"))
                     { ims = col1; }
                    if ((aa[1, col1].Value?.ToString().Trim().ToLower()=="inviteefullname")|| (aa[1, col1].Value.ToString().Trim().ToLower()=="fullname"))
                     { ifn = col1; }
                    if ((aa[1, col1].Value.ToString().Trim().ToLower()=="inviteeemail")|| (aa[1, col1].Value.ToString().Trim().ToLower()=="email"))
                     { iem = col1; }
                    if ((aa[1, col1].Value.ToString().Trim().ToLower()=="inviteelanguage")|| (aa[1, col1].Value.ToString().Trim().ToLower()=="language"))
                     { inl = col1; }
                    if ((aa[1, col1].Value.ToString().Trim().ToLower()=="inviteelocation")|| (aa[1, col1].Value.ToString().Trim().ToLower()=="location"))
                     { lct = col1; }
                    if (aa[1, col1].Value?.ToString().Trim().ToLower()=="transactionchannel") { txc = col1; }
                    if (aa[1, col1].Value?.ToString().Trim().ToLower()=="transactiontype") { txt = col1; }
                    if (aa[1, col1].Value?.ToString().Trim().ToLower()=="transactiondate") { txd = col1; }
                    if (aa[1, col1].Value?.ToString().Trim().ToLower()=="interactionchannel") { inc = col1; }
                    if (aa[1, col1].Value?.ToString().Trim().ToLower()=="answereddate") { asd = col1; }
                    if (aa[1, col1].Value?.ToString().Trim().ToLower()=="customdata1") { cs1 = col1; }
                    if (aa[1, col1].Value.ToString().Trim().ToLower() == "customdata2") { cs2 = col1; }
                    if (aa[1, col1].Value.ToString().Trim().ToLower() == "customdata3") { cs3 = col1; }
                    if (aa[1, col1].Value.ToString().Trim().ToLower() == "customdata4") { cs4 = col1; }
                    if (aa[1, col1].Value.ToString().Trim().ToLower() == "customdata5") { cs5 = col1; }
                    if (aa[1, col1].Value.ToString().Trim().ToLower() == "customdata6") { cs6 = col1; }
                    if (aa[1, col1].Value.ToString().Trim().ToLower() == "customdata7") { cs7 = col1; }
                    if (aa[1, col1].Value.ToString().Trim().ToLower() == "customdata8") { cs8 = col1; }
                    if (aa[1, col1].Value.ToString().Trim().ToLower() == "customdata9") { cs9 = col1; }
                    if (aa[1, col1].Value.ToString().Trim().ToLower()=="customData10") { cs10 = col1; }
                    if (aa[1, col1].Value.ToString().Contains("Question_NPS")) { qnp = col1; }
                    if (aa[1, col1].Value.ToString().Contains("Question1Value")) { vl1 = col1; }
                    if (aa[1, col1].Value.ToString().Contains("Question_Comment")) { qcm = col1; }
                    if (aa[1, col1].Value.ToString().Contains("Question2Value")) { vl2 = col1; }
                    if (aa[1, col1].Value.ToString().Contains("InviteeSegment"))
                    {
                        string rowValue = aa[1, col1].Value.ToString().Remove(0, 15).ToString();
                        InviteeSegment invSegment1 = new InviteeSegment(rowValue, col1.ToString());
                        segmentList = new SegmentColumn(col1, rowValue);
                        AllOf.Add(segmentList);
                    }

                }

                while (string.IsNullOrWhiteSpace(aa[row, col].Value?.ToString()) == false)
                {
                    Invitee p = new();
                    int r=number+row;
                    p.InviteeId = iid > 0 && aa[row, iid].Value != null ? aa[row, iid].Value?.ToString() : "TestUser_"+r.ToString();
                    p.InviteeEmail = iem > 0 && aa[row, iem].Value != null ? aa[row, iem].Value?.ToString() : r.ToString()+"@alternatest.com";
                    p.InviteeMSISDN = ims > 0 && aa[row, ims].Value != null ? aa[row, ims].Value?.ToString() : "057000"+r.ToString();
                    p.InviteeFullName = ifn > 0 && aa[row, ifn].Value != null ? aa[row, ifn].Value?.ToString() : "TestUser "+r.ToString();
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
                            Value = vl2 > 0 && aa[row, vl2].Value != null ? aa[row, vl2].Value?.ToString() : "Test",
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

        private static void GetToken(Environment env)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, env.tokenPath)
            {
                Content = new FormUrlEncodedContent(new Dictionary<string, string>
                {
                    {"client_id", env.client_id},
                    {"client_secret", env.client_secret},
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
