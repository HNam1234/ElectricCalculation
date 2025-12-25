using System.Collections.Generic;

namespace ElectricCalculation.Models
{
    public sealed class ProjectFile
    {
        public string PeriodLabel { get; set; } = string.Empty;

        public List<ProjectCustomer> Customers { get; set; } = new();
    }

    public sealed class ProjectCustomer
    {
        public int SequenceNumber { get; set; }
        public string Name { get; set; } = string.Empty;
        public string GroupName { get; set; } = string.Empty;
        public string Category { get; set; } = string.Empty;
        public string Address { get; set; } = string.Empty;
        public string RepresentativeName { get; set; } = string.Empty;
        public string HouseholdPhone { get; set; } = string.Empty;
        public string Phone { get; set; } = string.Empty;
        public string BuildingName { get; set; } = string.Empty;
        public string MeterNumber { get; set; } = string.Empty;
        public string Substation { get; set; } = string.Empty;
        public string Page { get; set; } = string.Empty;
        public string PerformedBy { get; set; } = string.Empty;
        public string Location { get; set; } = string.Empty;

        public decimal PreviousIndex { get; set; }
        public decimal? CurrentIndex { get; set; }
        public decimal Multiplier { get; set; } = 1;
        public decimal SubsidizedKwh { get; set; }
        public decimal UnitPrice { get; set; }
    }
}
