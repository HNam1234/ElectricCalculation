using System;
using System.Text.RegularExpressions;

namespace ElectricCalculation.Helpers
{
    public static class TextNormalization
    {
        private const string NoGroupKey = "\u0000NO_GROUP\u0000";

        private static readonly Regex MultiWhitespaceRegex = new(@"\s+", RegexOptions.Compiled);
        private static readonly Regex DashSpacingRegex = new(@"\s*[-–—]\s*", RegexOptions.Compiled);
        private static readonly Regex SlashSpacingRegex = new(@"\s*/\s*", RegexOptions.Compiled);

        public static string NormalizeForDisplay(string? text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            var value = text
                .Replace('\u00A0', ' ')
                .Replace('\u2007', ' ')
                .Replace('\u202F', ' ');

            value = MultiWhitespaceRegex.Replace(value, " ").Trim();
            value = DashSpacingRegex.Replace(value, " - ");
            value = SlashSpacingRegex.Replace(value, "/");
            value = MultiWhitespaceRegex.Replace(value, " ").Trim();

            return value;
        }

        public static string BuildGroupKey(string? groupName)
        {
            var normalized = NormalizeForDisplay(groupName);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return NoGroupKey;
            }

            return normalized.ToUpperInvariant();
        }

        public static bool IsNoGroupKey(string? groupKey)
        {
            return string.Equals(groupKey, NoGroupKey, StringComparison.Ordinal);
        }
    }
}
