using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using ElectricCalculation.Models;

namespace ElectricCalculation.Services
{
    public static class ProjectFileService
    {
        private static readonly JsonSerializerOptions JsonOptions = new()
        {
            WriteIndented = true,
            PropertyNameCaseInsensitive = true
        };

        public static void Save(string filePath, string periodLabel, IEnumerable<Customer> customers)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("File path is required.", nameof(filePath));
            }

            var list = customers?.ToList() ?? new List<Customer>();

            var payload = new ProjectFile
            {
                PeriodLabel = periodLabel ?? string.Empty,
                Customers = list.Select(c => new ProjectCustomer
                {
                    SequenceNumber = c.SequenceNumber,
                    Name = c.Name ?? string.Empty,
                    GroupName = c.GroupName ?? string.Empty,
                    Category = c.Category ?? string.Empty,
                    Address = c.Address ?? string.Empty,
                    RepresentativeName = c.RepresentativeName ?? string.Empty,
                    HouseholdPhone = c.HouseholdPhone ?? string.Empty,
                    Phone = c.Phone ?? string.Empty,
                    BuildingName = c.BuildingName ?? string.Empty,
                    MeterNumber = c.MeterNumber ?? string.Empty,
                    Substation = c.Substation ?? string.Empty,
                    Page = c.Page ?? string.Empty,
                    PerformedBy = c.PerformedBy ?? string.Empty,
                    Location = c.Location ?? string.Empty,
                    PreviousIndex = c.PreviousIndex,
                    CurrentIndex = c.CurrentIndex,
                    Multiplier = c.Multiplier,
                    SubsidizedKwh = c.SubsidizedKwh,
                    SubsidizedPercent = c.SubsidizedPercent,
                    UnitPrice = c.UnitPrice
                }).ToList()
            };

            var json = JsonSerializer.Serialize(payload, JsonOptions);
            File.WriteAllText(filePath, json);
        }

        public static (string PeriodLabel, List<Customer> Customers) Load(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("File path is required.", nameof(filePath));
            }

            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("Data file not found.", filePath);
            }

            var json = File.ReadAllText(filePath);
            var payload = JsonSerializer.Deserialize<ProjectFile>(json, JsonOptions)
                ?? new ProjectFile();

            var customers = (payload.Customers ?? new List<ProjectCustomer>())
                .Select(c => new Customer
                {
                    SequenceNumber = c.SequenceNumber,
                    Name = c.Name ?? string.Empty,
                    GroupName = c.GroupName ?? string.Empty,
                    Category = c.Category ?? string.Empty,
                    Address = c.Address ?? string.Empty,
                    RepresentativeName = c.RepresentativeName ?? string.Empty,
                    HouseholdPhone = c.HouseholdPhone ?? string.Empty,
                    Phone = c.Phone ?? string.Empty,
                    BuildingName = c.BuildingName ?? string.Empty,
                    MeterNumber = c.MeterNumber ?? string.Empty,
                    Substation = c.Substation ?? string.Empty,
                    Page = c.Page ?? string.Empty,
                    PerformedBy = c.PerformedBy ?? string.Empty,
                    Location = c.Location ?? string.Empty,
                    PreviousIndex = c.PreviousIndex,
                    CurrentIndex = c.CurrentIndex,
                    Multiplier = c.Multiplier <= 0 ? 1 : c.Multiplier,
                    SubsidizedKwh = c.SubsidizedKwh,
                    SubsidizedPercent = c.SubsidizedPercent,
                    UnitPrice = c.UnitPrice
                })
                .OrderBy(c => c.SequenceNumber)
                .ToList();

            return (payload.PeriodLabel ?? string.Empty, customers);
        }
    }
}
