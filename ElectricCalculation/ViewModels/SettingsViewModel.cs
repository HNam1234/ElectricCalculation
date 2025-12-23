using System;
using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ElectricCalculation.Models;

namespace ElectricCalculation.ViewModels
{
    public partial class SettingsViewModel : ObservableObject
    {
        [ObservableProperty]
        private bool? dialogResult;

        [ObservableProperty]
        private string defaultUnitPrice = "0";

        [ObservableProperty]
        private string defaultMultiplier = "1";

        [ObservableProperty]
        private string defaultSubsidizedKwh = "0";

        [ObservableProperty]
        private string defaultSubsidizedPercent = "0";

        [ObservableProperty]
        private string defaultPerformedBy = string.Empty;

        [ObservableProperty]
        private bool applyDefaultsOnNewRow = true;

        [ObservableProperty]
        private bool applyDefaultsOnImport = true;

        [ObservableProperty]
        private bool overrideExistingValues;

        [ObservableProperty]
        private string errorMessage = string.Empty;

        public SettingsViewModel(AppSettings settings)
        {
            var s = settings ?? new AppSettings();
            DefaultUnitPrice = s.DefaultUnitPrice.ToString("0.##", CultureInfo.CurrentCulture);
            DefaultMultiplier = s.DefaultMultiplier.ToString("0.##", CultureInfo.CurrentCulture);
            DefaultSubsidizedKwh = s.DefaultSubsidizedKwh.ToString("0.##", CultureInfo.CurrentCulture);
            DefaultSubsidizedPercent = s.DefaultSubsidizedPercent.ToString("0.##", CultureInfo.CurrentCulture);
            DefaultPerformedBy = s.DefaultPerformedBy ?? string.Empty;
            ApplyDefaultsOnNewRow = s.ApplyDefaultsOnNewRow;
            ApplyDefaultsOnImport = s.ApplyDefaultsOnImport;
            OverrideExistingValues = s.OverrideExistingValues;
        }

        public AppSettings BuildSettings()
        {
            ErrorMessage = string.Empty;

            if (!TryParseDecimal(DefaultUnitPrice, out var unitPrice, allowNegative: false) ||
                !TryParseDecimal(DefaultMultiplier, out var multiplier, allowNegative: false) ||
                !TryParseDecimal(DefaultSubsidizedKwh, out var subsidizedKwh, allowNegative: false) ||
                !TryParseDecimal(DefaultSubsidizedPercent, out var subsidizedPercent, allowNegative: false))
            {
                throw new InvalidOperationException("Invalid numeric input.");
            }

            if (multiplier <= 0)
            {
                multiplier = 1;
            }

            return new AppSettings
            {
                DefaultUnitPrice = unitPrice,
                DefaultMultiplier = multiplier,
                DefaultSubsidizedKwh = subsidizedKwh,
                DefaultSubsidizedPercent = subsidizedPercent,
                DefaultPerformedBy = DefaultPerformedBy ?? string.Empty,
                ApplyDefaultsOnNewRow = ApplyDefaultsOnNewRow,
                ApplyDefaultsOnImport = ApplyDefaultsOnImport,
                OverrideExistingValues = OverrideExistingValues
            };
        }

        [RelayCommand]
        private void Ok()
        {
            try
            {
                _ = BuildSettings();
                DialogResult = true;
            }
            catch
            {
                if (string.IsNullOrWhiteSpace(ErrorMessage))
                {
                    ErrorMessage = "Gia tri nhap khong hop le. Hay kiem tra lai cac o so.";
                }
            }
        }

        [RelayCommand]
        private void Cancel()
        {
            DialogResult = false;
        }

        private bool TryParseDecimal(string? text, out decimal value, bool allowNegative)
        {
            value = 0m;

            if (string.IsNullOrWhiteSpace(text))
            {
                return true;
            }

            if (!decimal.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out value) &&
                !decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out value))
            {
                ErrorMessage = $"Khong doc duoc gia tri: '{text}'.";
                return false;
            }

            if (!allowNegative && value < 0)
            {
                ErrorMessage = "Gia tri phai >= 0.";
                return false;
            }

            return true;
        }
    }
}

