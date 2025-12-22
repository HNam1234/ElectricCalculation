using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ElectricCalculation.Services
{
    public static class VietnameseNumberTextService
    {
        public static string ConvertAmountToText(decimal amount)
        {
            var rounded = Math.Round(amount, 0, MidpointRounding.AwayFromZero);
            if (rounded <= 0)
            {
                return "Không đồng";
            }

            var value = (long)rounded;
            if (value < 0)
            {
                value = -value;
            }

            string[] unitNumbers =
            {
                "không", "một", "hai", "ba", "bốn",
                "năm", "sáu", "bảy", "tám", "chín"
            };

            string[] placeValues =
            {
                string.Empty,
                "nghìn",
                "triệu",
                "tỷ",
                "nghìn tỷ",
                "triệu tỷ"
            };

            string ReadThreeDigits(int number, bool isMostSignificantGroup)
            {
                int hundreds = number / 100;
                int tens = (number % 100) / 10;
                int ones = number % 10;

                var sb = new StringBuilder();

                if (hundreds > 0 || !isMostSignificantGroup)
                {
                    if (hundreds > 0)
                    {
                        sb.Append(unitNumbers[hundreds]).Append(" trăm");
                    }
                    else if (!isMostSignificantGroup)
                    {
                        sb.Append("không trăm");
                    }
                }

                if (tens > 1)
                {
                    if (sb.Length > 0)
                    {
                        sb.Append(' ');
                    }

                    sb.Append(unitNumbers[tens]).Append(" mươi");

                    if (ones == 1)
                    {
                        sb.Append(" mốt");
                    }
                    else if (ones == 4)
                    {
                        sb.Append(" tư");
                    }
                    else if (ones == 5)
                    {
                        sb.Append(" lăm");
                    }
                    else if (ones > 0)
                    {
                        sb.Append(' ').Append(unitNumbers[ones]);
                    }
                }
                else if (tens == 1)
                {
                    if (sb.Length > 0)
                    {
                        sb.Append(' ');
                    }

                    sb.Append("mười");

                    if (ones == 1)
                    {
                        sb.Append(" một");
                    }
                    else if (ones == 4)
                    {
                        sb.Append(" bốn");
                    }
                    else if (ones == 5)
                    {
                        sb.Append(" lăm");
                    }
                    else if (ones > 0)
                    {
                        sb.Append(' ').Append(unitNumbers[ones]);
                    }
                }
                else if (tens == 0 && ones > 0)
                {
                    if (sb.Length > 0)
                    {
                        sb.Append(" lẻ");
                    }

                    if (ones == 5 && sb.Length > 0)
                    {
                        sb.Append(" năm");
                    }
                    else
                    {
                        sb.Append(' ').Append(unitNumbers[ones]);
                    }
                }

                return sb.ToString().Trim();
            }

            var groups = new List<int>(capacity: placeValues.Length);
            while (value > 0 && groups.Count < placeValues.Length)
            {
                groups.Add((int)(value % 1000));
                value /= 1000;
            }

            var highestGroupIndex = -1;
            for (var i = groups.Count - 1; i >= 0; i--)
            {
                if (groups[i] > 0)
                {
                    highestGroupIndex = i;
                    break;
                }
            }

            var resultBuilder = new StringBuilder();

            for (var groupIndex = highestGroupIndex; groupIndex >= 0; groupIndex--)
            {
                var groupNumber = groups[groupIndex];
                if (groupNumber <= 0)
                {
                    continue;
                }

                var groupText = ReadThreeDigits(groupNumber, isMostSignificantGroup: groupIndex == highestGroupIndex);
                if (string.IsNullOrEmpty(groupText))
                {
                    continue;
                }

                if (resultBuilder.Length > 0)
                {
                    resultBuilder.Append(' ');
                }

                resultBuilder.Append(groupText);

                var unitText = placeValues[groupIndex];
                if (!string.IsNullOrEmpty(unitText))
                {
                    resultBuilder.Append(' ').Append(unitText);
                }
            }

            var result = resultBuilder.ToString().Trim();
            if (result.Length == 0)
            {
                result = "không";
            }

            result = char.ToUpper(result[0], CultureInfo.CurrentCulture) + result[1..] + " đồng";
            return result;
        }
    }
}

