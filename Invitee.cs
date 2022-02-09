using System;
using System.Collections.Generic;

namespace UploadInviteewithAnswersClient
{
    public class InviteeSegment
    {
        public InviteeSegment(string segmentGroupName, string segmentId)
        {
            this.SegmentGroupName = segmentGroupName;
            this.SegmentId = segmentId;

        }
        public string SegmentGroupName { get; set; }
        public string SegmentId { get; set; }


    }
    public class SegmentColumn
    {
        public int ColumnId { get; set; }
        public string ColumnTitle { get; set; }
        public SegmentColumn()
        {

        }

        public SegmentColumn(int id, string tittle)
        {
            this.ColumnId = id;
            this.ColumnTitle = tittle;
        }
    }
    public class InviteeAnswer
    {
        public string QuestionIntegrationCode { get; set; }
        public string Value { get; set; }
        public string AnsweredDate { get; set; }
    }

    public class Invitee
    {

        public Invitee (Invitee p)
        {
           this.InviteeId = p.InviteeId;
            this.InviteeEmail = p.InviteeEmail;
            this.InviteeMSISDN = p.InviteeMSISDN;
            this.InviteeFullName = p.InviteeFullName;
            this.InviteeLanguage = p.InviteeLanguage;
            this.InviteeLocation = p.InviteeLocation;
            this.TransactionChannel = p.TransactionChannel;
            this.TransactionType = p.TransactionType;
            this.TransactionDate = p.TransactionDate;
            this.InteractionChannel = p.InteractionChannel;
            this.CustomData1 = p.CustomData1;
            this.CustomData2 = p.CustomData2;
            this.CustomData3 = p.CustomData3;
            this.CustomData4 = p.CustomData4;
            this.CustomData5 = p.CustomData5;
            this.CustomData6 = p.CustomData6;
            this.CustomData7 = p.CustomData7;
            this.CustomData8 = p.CustomData8;
            this.CustomData9 = p.CustomData9;
            this.CustomData10 =p.CustomData10;
        }
        public Invitee()
        {

        }
        public string InviteeId { get; set; }
        public string InviteeEmail { get; set; }
        public string InviteeMSISDN { get; set; }
        public string InviteeFullName { get; set; }
        public string InviteeLanguage { get; set; }
        public string InviteeLocation { get; set; }
        public string TransactionChannel { get; set; }
        public string TransactionType { get; set; }
        public string TransactionDate { get; set; }
        public string InteractionChannel { get; set; }
        public List<InviteeSegment> InviteeSegments { get; set; }
        public List<InviteeAnswer> InviteeAnswers { get; set; }
        public string CustomData1 { get; set; }
        public string CustomData2 { get; set; }
        public string CustomData3 { get; set; }
        public string CustomData4 { get; set; }
        public string CustomData5 { get; set; }
        public string CustomData6 { get; set; }
        public string CustomData7 { get; set; }
        public string CustomData8 { get; set; }
        public string CustomData9 { get; set; }
        public string CustomData10 { get; set; }



    }

}