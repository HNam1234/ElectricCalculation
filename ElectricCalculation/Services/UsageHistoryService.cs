using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using ElectricCalculation.Models;

namespace ElectricCalculation.Services
{
    public static class UsageHistoryService
    {
        private static readonly Regex MonthYearRegex = new(@"(\d{1,2})\s*/\s*(\d{4})", RegexOptions.Compiled);

        public static IReadOnlyDictionary<string, decimal> BuildAverageConsumptionByMeterKey(
            string? currentPeriodLabel,
            IEnumerable<Customer> currentCustomers,
            int periodsToAverage = 3,
            string? excludeSnapshotPath = null)
        {
            if (!TryParsePeriod(currentPeriodLabel, out var currentMonth, out var currentYear))
            {
                return new Dictionary<string, decimal>();
            }

            var currentPeriodKey = currentYear * 100 + currentMonth;

            var targetKeys = new HashSet<string>(
                currentCustomers
                    .Select(BuildMeterKey)
                    .Where(k => !string.IsNullOrWhiteSpace(k)),
                StringComparer.OrdinalIgnoreCase);

            if (targetKeys.Count == 0)
            {
                return new Dictionary<string, decimal>();
            }

            var snapshots = SaveGameService.ListSnapshots(maxCount: 500);

            var priorSnapshots = snapshots
                .Select(s =>
                {
                    var ok = TryParsePeriod(s.PeriodLabel, out var month, out var year);
                    return new
                    {
                        s.Path,
                        s.SavedAt,
                        PeriodKey = ok ? year * 100 + month : (int?)null
                    };
                })
                .Where(s => s.PeriodKey != null && s.PeriodKey.Value < currentPeriodKey)
                .Where(s => string.IsNullOrWhiteSpace(excludeSnapshotPath) ||
                            !string.Equals(s.Path, excludeSnapshotPath, StringComparison.OrdinalIgnoreCase))
                .GroupBy(s => s.PeriodKey!.Value)
                .Select(g => g.OrderByDescending(x => x.SavedAt).First())
                .OrderByDescending(x => x.PeriodKey)
                .Take(Math.Max(0, periodsToAverage))
                .ToList();

            if (priorSnapshots.Count == 0)
            {
                return new Dictionary<string, decimal>();
            }

            var totals = new Dictionary<string, (decimal Sum, int Count)>(StringComparer.OrdinalIgnoreCase);

            foreach (var snapshot in priorSnapshots)
            {
                List<Customer> customers;
                try
                {
                    (_, customers) = ProjectFileService.Load(snapshot.Path);
                }
                catch
                {
                    continue;
                }

                foreach (var customer in customers)
                {
                    var key = BuildMeterKey(customer);
                    if (string.IsNullOrWhiteSpace(key) || !targetKeys.Contains(key))
                    {
                        continue;
                    }

                    if (customer.CurrentIndex == null)
                    {
                        continue;
                    }

                    var delta = customer.CurrentIndex.Value - customer.PreviousIndex;
                    if (delta < 0)
                    {
                        continue;
                    }

                    var multiplier = customer.Multiplier <= 0 ? 1 : customer.Multiplier;
                    var consumption = delta * multiplier;

                    if (!totals.TryGetValue(key, out var total))
                    {
                        total = (0m, 0);
                    }

                    totals[key] = (total.Sum + consumption, total.Count + 1);
                }
            }

            var result = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
            foreach (var (key, total) in totals)
            {
                if (total.Count <= 0)
                {
                    continue;
                }

                result[key] = total.Sum / total.Count;
            }

            return result;
        }

        public static bool TryParsePeriod(string? periodLabel, out int month, out int year)
        {
            month = 0;
            year = 0;

            if (string.IsNullOrWhiteSpace(periodLabel))
            {
                return false;
            }

            var match = MonthYearRegex.Match(periodLabel);
            if (match.Success &&
                int.TryParse(match.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out month) &&
                int.TryParse(match.Groups[2].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out year))
            {
                return month is >= 1 and <= 12 && year >= 2000;
            }

            return false;
        }

        public static string BuildMeterKey(Customer customer)
        {
            if (!string.IsNullOrWhiteSpace(customer.MeterNumber))
            {
                return customer.MeterNumber.Trim();
            }

            if (!string.IsNullOrWhiteSpace(customer.Name) && !string.IsNullOrWhiteSpace(customer.Location))
            {
                return $"{customer.Name.Trim()}|{customer.Location.Trim()}";
            }

            if (!string.IsNullOrWhiteSpace(customer.Name))
            {
                return customer.Name.Trim();
            }

            return customer.SequenceNumber.ToString(CultureInfo.InvariantCulture);
        }
    }
}

