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
        static async Task Main(string[] args)
        {
            //application is capable of creating different berarer token according to users desired enviroment, the default one is staging just in case any wrong input
            Console.WriteLine("Please type your environment:       (staging / pilot / pilot2)");
            string env = Console.ReadLine().ToLower();
            Environment environment = new(env);
            env = environment.environmentName;
            Console.WriteLine(env + " is selected.");
            if (string.IsNullOrEmpty(environment.bearer_token))
            {
                GetToken(environment);
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //upload file path
            var file = new FileInfo(fileName: @"C:\Users\Mustafa\Documents\answer1.xlsx");
            using var package = new ExcelPackage(file);
            List<Invitee> invitees = new();
            List<Invitee> peopleFromExcel = await LoadExcelFile(file);

            foreach (Invitee p in peopleFromExcel)
            {
                Invitee invitee = new(p)
                {
                    InviteeSegments = new List<InviteeSegment>(p.InviteeSegments),
                    InviteeAnswers= new List<InviteeAnswer>(p.InviteeAnswers)
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
                            client.DefaultRequestHeaders.Add("Authorization", "Bearer " + environment.bearer_token);
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
            //following variables are to dedect the order of the columns.
            int col, iid, ims, ifn, iem, txc, txt, inl, txd, inc, sg1, sg2, qnp, vl1, qcm, vl2, asd, lct, cs1, cs2, cs3, cs4, cs5, cs6, cs7, cs8, cs9, cs10;
            col = 1;
            iid = ims = ifn = iem = txc = txt = inl = txd = inc = sg1 = sg2 = qnp = vl1 = qcm = vl2 = asd = lct = cs1 = cs2 = cs3 = cs4 = cs5 = cs6 = cs7 = cs8 = cs9 = cs10 = 0;

            //to understand is there more than one segment or question, need to create list including their parameters
            List<SegmentColumn> allSegments = new();
            SegmentColumn segmentcolumn = new();
            List<InviteeSegment> newSegments = new();

            List<QuestionColumn> allQuestions = new();
            QuestionColumn questioncolumn = new();
            List<InviteeAnswer> newAnswers = new();

            var aa = ws.Cells;
            try
            {
                int rowCount = aa.Worksheet.Columns.EndColumn;

                //for the case if there is no invitee information such as invitee, random number will be used and it will go incremently.

                Random random = new();
                int number = random.Next(10000, 1000000);
                for (int c = 1; c <= rowCount; c++)
                {
                    if (aa[1, c].Value.ToString().Trim().ToLower() == "inviteeid") { iid = c; }

                    //to check whether inviteensisdn user or only msisdn
                    if ((aa[1, c].Value.ToString().Trim().ToLower() == "inviteemsisdn") || (aa[1, c].Value.ToString().Trim().ToLower() == "msisdn"))
                    { ims = c; }
                    if ((aa[1, c].Value?.ToString().Trim().ToLower() == "inviteefullname") || (aa[1, c].Value.ToString().Trim().ToLower() == "fullname"))
                    { ifn = c; }
                    if ((aa[1, c].Value.ToString().Trim().ToLower() == "inviteeemail") || (aa[1, c].Value.ToString().Trim().ToLower() == "email"))
                    { iem = c; }
                    if ((aa[1, c].Value.ToString().Trim().ToLower() == "inviteelanguage") || (aa[1, c].Value.ToString().Trim().ToLower() == "language"))
                    { inl = c; }
                    if ((aa[1, c].Value.ToString().Trim().ToLower() == "inviteelocation") || (aa[1, c].Value.ToString().Trim().ToLower() == "location"))
                    { lct = c; }
                    if (aa[1, c].Value?.ToString().Trim().ToLower() == "transactionchannel") { txc = c; }
                    if (aa[1, c].Value?.ToString().Trim().ToLower() == "transactiontype") { txt = c; }
                    if (aa[1, c].Value?.ToString().Trim().ToLower() == "transactiondate") { txd = c; }
                    if (aa[1, c].Value?.ToString().Trim().ToLower() == "interactionchannel") { inc = c; }
                    if (aa[1, c].Value?.ToString().Trim().ToLower() == "answereddate") { asd = c; }
                    if (aa[1, c].Value?.ToString().Trim().ToLower() == "customdata1") { cs1 = c; }
                    if (aa[1, c].Value.ToString().Trim().ToLower() == "customdata2") { cs2 = c; }
                    if (aa[1, c].Value.ToString().Trim().ToLower() == "customdata3") { cs3 = c; }
                    if (aa[1, c].Value.ToString().Trim().ToLower() == "customdata4") { cs4 = c; }
                    if (aa[1, c].Value.ToString().Trim().ToLower() == "customdata5") { cs5 = c; }
                    if (aa[1, c].Value.ToString().Trim().ToLower() == "customdata6") { cs6 = c; }
                    if (aa[1, c].Value.ToString().Trim().ToLower() == "customdata7") { cs7 = c; }
                    if (aa[1, c].Value.ToString().Trim().ToLower() == "customdata8") { cs8 = c; }
                    if (aa[1, c].Value.ToString().Trim().ToLower() == "customdata9") { cs9 = c; }
                    if (aa[1, c].Value.ToString().Trim().ToLower() == "customData10") { cs10 = c; }
                    //detect all the segment columns start with InviteeSegment_ by their order and title and adds to the list. Non-removed part is accepted as SegmentGroupName
                    if (aa[1, c].Value.ToString().Contains("InviteeSegment"))
                    {
                        string SegmentRowValue = aa[1, c].Value.ToString().Remove(0, 15).ToString();
                        segmentcolumn = new SegmentColumn(c, SegmentRowValue);
                        allSegments.Add(segmentcolumn);
                    }
                    //detect all the question columns start with Question_ by their order and title and adds to the list. Non-removed part is accepted as Question Integration Code.

                    if ((aa[1, c].Value.ToString().Contains("Question_")))
                    {
                        string QuestionRowValue = aa[1, c].Value.ToString().Remove(0, 9).ToString();
                        questioncolumn = new QuestionColumn(c, QuestionRowValue);
                        allQuestions.Add(questioncolumn);
                    }

                }
                //journey starts
                while (string.IsNullOrWhiteSpace(aa[row, col].Value?.ToString()) == false)
                {
                    Invitee p = new();

                    int r = number + row;

                    //if there is any null value on any excel cell, no error occurs, the following value will be added to the related parameters.
                    //--> next step check all the parameters length accordng to data types
                    p.InviteeId = iid > 0 && aa[row, iid].Value != null ? aa[row, iid].Value?.ToString() : "TestUser_" + r.ToString();
                    p.InviteeEmail = iem > 0 && aa[row, iem].Value != null ? aa[row, iem].Value?.ToString() : r.ToString() + "@alternatest.com";
                    p.InviteeMSISDN = ims > 0 && aa[row, ims].Value != null ? aa[row, ims].Value?.ToString() : "057000" + r.ToString();
                    p.InviteeFullName = ifn > 0 && aa[row, ifn].Value != null ? aa[row, ifn].Value?.ToString() : "TestUser " + r.ToString();
                    p.InviteeLanguage = inl > 0 && aa[row, inl].Value != null ? aa[row, inl].Value.ToString() : "tr";
                    p.TransactionChannel = txc > 0 && aa[row, txc].Value != null ? aa[row, txc].Value?.ToString() : "";
                    p.TransactionType = txt > 0 && aa[row, txt].Value != null ? aa[row, txt].Value?.ToString() : "";
                    p.TransactionDate = txd > 0 && aa[row, txd].Value != null ? aa[row, txd].Value.ToString() : DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss");
                    p.InteractionChannel = inc > 0 && aa[row, inc].Value != null ? aa[row, inc].Value.ToString() : "8";
                    p.InviteeLocation = lct > 0 && aa[row, lct].Value != null ? aa[row, lct].Value.ToString() : "";
                    p.CustomData1 = cs1 > 0 && aa[row, cs1].Value != null ? aa[row, cs1].Value?.ToString() : "";
                    p.CustomData2 = cs2 > 0 && aa[row, cs2].Value != null ? aa[row, cs2].Value?.ToString() : "";
                    p.CustomData3 = cs3 > 0 && aa[row, cs3].Value != null ? aa[row, cs3].Value.ToString() : "";
                    p.CustomData4 = cs4 > 0 && aa[row, cs4].Value != null ? aa[row, cs4].Value?.ToString() : "";
                    p.CustomData5 = cs5 > 0 && aa[row, cs5].Value != null ? aa[row, cs5].Value?.ToString() : "";
                    p.CustomData6 = cs6 > 0 && aa[row, cs6].Value != null ? aa[row, cs6].Value?.ToString() : "";
                    p.CustomData7 = cs7 > 0 && aa[row, cs7].Value != null ? aa[row, cs7].Value?.ToString() : "";
                    p.CustomData8 = cs8 > 0 && aa[row, cs8].Value != null ? aa[row, cs8].Value?.ToString() : "";
                    p.CustomData9 = cs9 > 0 && aa[row, cs9].Value != null ? aa[row, cs9].Value?.ToString() : "";
                    p.CustomData10 = cs10 > 0 && aa[row, cs10].Value != null ? aa[row, cs10].Value?.ToString() : "";

                    if (allSegments.Count > 0)
                    {
                        foreach (var item in allSegments) //add Multiple segments
                        {
                            if (aa[row, item.ColumnId].Value != null) //important check for null segment value
                            {
                                try
                                {
                                    InviteeSegment newsegment = new(item.ColumnTitle, aa[row, item.ColumnId].Value.ToString());
                                    newSegments.Add(newsegment);
                                }
                                catch (Exception) { continue; }
                            }
                        };
                    }


                    p.InviteeSegments = new List<InviteeSegment>();
                    foreach (var eachsegment in newSegments)
                    {

                        try
                        {
                            p.InviteeSegments.Add(eachsegment);
                        }
                        catch (Exception) { continue; }
                    }



                    if (allQuestions.Count>0)
                    {
                        foreach (var item in allQuestions)
                        {
                            if (aa[row, item.ColumnId].Value != null) //important check for null question value
                            {
                                try
                                {
                                    //max 500 characters are allowed
                                    string questionValue = aa[row, item.ColumnId].Value.ToString();
                                    questionValue= questionValue.Length>500 ? questionValue.Substring(0, 500) : questionValue;

                                    InviteeAnswer newanswer = new(item.ColumnTitle, questionValue);
                                    newanswer.AnsweredDate = asd > 0 && aa[row, asd].Value != null ? aa[row, asd].Value.ToString() : DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss");
                                    newAnswers.Add(newanswer);

                                }
                                catch (Exception) { continue; }
                            }
                        };
                    }



                    p.InviteeAnswers=new List<InviteeAnswer>();
                    foreach (var eachanswer in newAnswers)
                    {
                        try
                        {
                            p.InviteeAnswers.Add(eachanswer);
                        }
                        catch (Exception) {   continue; }                   
                    }

                    output.Add(p);
                    newSegments.Clear(); //drop the segments and ready to load new ones for new invitees
                    newAnswers.Clear(); //drop the questions and ready to load new ones for new invitees
                    row++;

                }
            }
            catch (Exception)
            {
                throw;
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
            env.bearer_token = payload.Value<string>("access_token");
        }
    }
}
