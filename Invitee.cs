using System;
using System.Collections.Generic;

namespace UploadInviteewithAnswersClient
{
    public class InviteeSegment
    {
        public string SegmentGroupName { get; set; }
        public string SegmentId { get; set; }
        public InviteeSegment()
        {
        }
        public InviteeSegment(string sgr, string sid)
        {
            this.SegmentGroupName = sgr;
            this.SegmentId = sid;
        }

        public void AddSegment(string sgr, string sid)
        {
            this.SegmentGroupName = sgr;
            this.SegmentId = sid;
        }
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