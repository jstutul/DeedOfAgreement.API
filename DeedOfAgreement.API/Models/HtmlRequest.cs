namespace DeedOfAgreement.API.Models
{
    public class HtmlRequest
    {
        public string FileNo { get; set; }
        public int OrganizationId { get; set; }
        public int ProjectId { get; set; }
        public string FileName => $"{FileNo}.docx";
    }
}
