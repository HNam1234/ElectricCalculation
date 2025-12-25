using System;
using System.Collections;
using System.Globalization;
using ElectricCalculation.Models;

namespace ElectricCalculation.ViewModels
{
    internal sealed class CustomerRouteComparer : IComparer
    {
        public static CustomerRouteComparer Instance { get; } = new();

        private CustomerRouteComparer()
        {
        }

        public int Compare(object? x, object? y)
        {
            if (ReferenceEquals(x, y))
            {
                return 0;
            }

            if (x is not Customer a)
            {
                return -1;
            }

            if (y is not Customer b)
            {
                return 1;
            }

            var location = CompareRouteText(a.Location, b.Location);
            if (location != 0)
            {
                return location;
            }

            var page = ComparePage(a.Page, b.Page);
            if (page != 0)
            {
                return page;
            }

            var sequence = a.SequenceNumber.CompareTo(b.SequenceNumber);
            if (sequence != 0)
            {
                return sequence;
            }

            return CompareRouteText(a.Name, b.Name);
        }

        private static int CompareRouteText(string? a, string? b)
        {
            var left = (a ?? string.Empty).Trim();
            var right = (b ?? string.Empty).Trim();

            var leftEmpty = string.IsNullOrWhiteSpace(left);
            var rightEmpty = string.IsNullOrWhiteSpace(right);

            if (leftEmpty && rightEmpty)
            {
                return 0;
            }

            if (leftEmpty)
            {
                return 1;
            }

            if (rightEmpty)
            {
                return -1;
            }

            return string.Compare(left, right, CultureInfo.CurrentCulture, CompareOptions.IgnoreCase);
        }

        private static int ComparePage(string? a, string? b)
        {
            if (TryExtractFirstInt(a, out var pageA) && TryExtractFirstInt(b, out var pageB))
            {
                return pageA.CompareTo(pageB);
            }

            return CompareRouteText(a, b);
        }

        private static bool TryExtractFirstInt(string? text, out int value)
        {
            value = 0;

            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            var span = text.AsSpan();
            var i = 0;
            while (i < span.Length && !char.IsDigit(span[i]))
            {
                i++;
            }

            if (i >= span.Length)
            {
                return false;
            }

            var start = i;
            while (i < span.Length && char.IsDigit(span[i]))
            {
                i++;
            }

            return int.TryParse(span.Slice(start, i - start), NumberStyles.Integer, CultureInfo.InvariantCulture, out value);
        }
    }
}

