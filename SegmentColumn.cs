namespace UploadInviteewithAnswersClient
{
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
}