namespace UploadInviteewithAnswersClient
{
    public class QuestionColumn
    {


        public QuestionColumn(int columnId, string columnTitle)
        {
            this.ColumnId = columnId;
            this.ColumnTitle = columnTitle;

        }
        public int ColumnId { get; set; }
        public string ColumnTitle { get; set; }
        public QuestionColumn()
        {
        }


    }
}