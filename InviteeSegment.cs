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
}