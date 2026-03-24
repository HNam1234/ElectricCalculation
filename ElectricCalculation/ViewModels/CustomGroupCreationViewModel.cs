using System;
using System.Collections.Generic;
using System.Linq;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace ElectricCalculation.ViewModels
{
    public partial class CustomGroupCreationViewModel : ObservableObject
    {
        public IReadOnlyList<string> SourceGroups { get; }

        public int SourceGroupCount => SourceGroups.Count;

        public int SelectedCustomerCount { get; }

        public bool ShowSelectionSummary => SelectedCustomerCount > 0;

        public bool ShowGroupList => SourceGroupCount > 0;

        [ObservableProperty]
        private bool? dialogResult;

        [ObservableProperty]
        private string groupName = string.Empty;

        public CustomGroupCreationViewModel(IEnumerable<string> sourceGroups, string? defaultGroupName = null)
            : this(sourceGroups, selectedCustomerCount: 0, defaultGroupName)
        {
        }

        public CustomGroupCreationViewModel(IEnumerable<string> sourceGroups, int selectedCustomerCount, string? defaultGroupName = null)
        {
            SelectedCustomerCount = Math.Max(0, selectedCustomerCount);
            SourceGroups = (sourceGroups ?? Array.Empty<string>())
                .Select(value => (value ?? string.Empty).Trim())
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .OrderBy(value => value, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            GroupName = (defaultGroupName ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(GroupName))
            {
                GroupName = "Nhóm custom";
            }
        }

        [RelayCommand]
        private void Ok()
        {
            DialogResult = true;
        }

        [RelayCommand]
        private void Cancel()
        {
            DialogResult = false;
        }
    }
}
