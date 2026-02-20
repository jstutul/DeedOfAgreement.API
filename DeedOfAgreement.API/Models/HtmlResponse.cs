namespace DeedOfAgreement.API.Models
{
    public class HtmlResponse
    {
        // Organization & Project
        public string OrganizationName { get; set; }
        public string ProjectName { get; set; }
        public string Location { get; set; }
        public string PS { get; set; }

        // Client Information
        public string ClientName { get; set; }
        public string FileNo { get; set; }
        public string FatherName { get; set; }
        public string MotherName { get; set; }
        public string Gender { get; set; }
        public string OccupationName { get; set; }
        public string PresentAddress { get; set; }
        public string PermanentAddress { get; set; }
        public string ClientNid { get; set; }
        public string Religion { get; set; }
        public string Btext { get; set; }
        public int TotalInstallment { get; set; }

        public string SaleTypeShortName { get; set; }

        // Additional Text Fields
        public string ClientDetails { get; set; }
        public string NomineeList { get; set; }
        public string OptionOne { get; set; }
        public string OptionTwo { get; set; }
        public string OptionTwoA { get; set; }
        public string A { get; set; }
        public string B { get; set; }
        public string C { get; set; }
        public string ScheduleB { get; set; }
        public string BText { get; set; }
    }
}
