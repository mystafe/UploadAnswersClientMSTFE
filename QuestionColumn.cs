namespace UploadInviteewithAnswersClient
{
    public class QuestionColumn
    {
        public int ColumnId { get; set; }
        public string ColumnTitle { get; set; }
        public QuestionColumn()
        {
        }

        public QuestionColumn(int id, string tittle)
        {
            this.ColumnId = id;
            this.ColumnTitle = tittle;
        }
    }
}