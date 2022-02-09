using System;

namespace UploadInviteewithAnswersClient
{
    public class InviteeAnswer
    {
        public InviteeAnswer(string questionIntegrationCode, string value)
        {
            QuestionIntegrationCode = questionIntegrationCode;
            Value = value;
            AnsweredDate = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss");
        }

        public string QuestionIntegrationCode { get; set; }
        public string Value { get; set; }
        public string AnsweredDate { get; set; }
        public InviteeAnswer(){
        }

        
        
        

    }
}