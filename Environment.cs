using System;
using System.Collections.Generic;

namespace UploadInviteewithAnswersClient
{
    public class Environment
    {

        public string environmentName { get; set; }
        public string tokenPath { get; set; }
        public string uploadPath { get; set; }
        public string client_id { get; set; }
        public string client_secret { get; set; }
        public string bearer_token { get; set; }

        public Environment(string env)
        {
            switch (env)
            {
                case "staging":
                    SetStage();
                    break;
                case "pilot":
                    SetPilot();
                    break;
                case "pilot2":
                    SetPilot2();
                    break;
                default:
                    SetStage();
                    break;
            }
        }

        public string SetStage()
        {
            environmentName = "staging";
            client_id = "e7f8ae6d-92ac-4443-8f6d-1210b4dd21c3";
            client_secret = "*****";
            tokenPath = "https://staging-api.alternacx.com/token";
            uploadPath = "https://staging-api.alternacx.com/api/upload/inviteesWithAnswers";
            bearer_token = "";

            return environmentName;
        }
        public string SetPilot()
        {
            environmentName = "pilot";
            client_id = "fa750f45-07d3-4549-b54f-45abda5491d6";
            client_secret = "*****";
            tokenPath = "https://pilot-api.alternacx.com/token";
            uploadPath = "https://piot-api.alternacx.com/api/upload/inviteesWithAnswers";
            bearer_token = "";

            return environmentName;
        }
        public string SetPilot2()
        {
            environmentName = "pilot2";
            client_id = "121d6c4a-fced-481f-8a9d-164f6094c0bb";
            client_secret = "*****";
            tokenPath = "https://pilot2-api.alternacx.com/token";
            uploadPath = "https://pilot2-api.alternacx.com/api/upload/inviteesWithAnswers";
            bearer_token = "";

            return environmentName;
        }




    }
}
