namespace ElectricCalculation.Models
{
    public class MeterReading
    {
        public decimal PreviousIndex { get; set; }

        public decimal CurrentIndex { get; set; }

        public decimal Multiplier { get; set; } = 1;

        public decimal SubsidizedKwh { get; set; }

        public decimal UnitPrice { get; set; }
    }
}

