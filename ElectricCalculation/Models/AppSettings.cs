namespace ElectricCalculation.Models
{
    public sealed class AppSettings
    {
        // Optional: if set, snapshots will be saved/listed from this folder (can be a UNC/network share).
        public string SharedSavesDirectory { get; set; } = string.Empty;

        public decimal DefaultUnitPrice { get; set; }

        public decimal DefaultMultiplier { get; set; } = 1m;

        public decimal DefaultSubsidizedKwh { get; set; }

        public string DefaultPerformedBy { get; set; } = string.Empty;

        public bool ApplyDefaultsOnNewRow { get; set; } = true;

        public bool ApplyDefaultsOnImport { get; set; } = true;

        public bool OverrideExistingValues { get; set; } = true;
    }
}
