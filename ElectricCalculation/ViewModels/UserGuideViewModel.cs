using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace ElectricCalculation.ViewModels
{
    public sealed record UserGuideStepItem(
        string StepTitle,
        string Description,
        BitmapSource Screenshot);

    public sealed record UserGuideSectionItem(
        string TabTitle,
        string Heading,
        string Description,
        IReadOnlyList<UserGuideStepItem> Steps);

    public partial class UserGuideViewModel : ObservableObject
    {
        [ObservableProperty]
        private bool? dialogResult;

        [ObservableProperty]
        private UserGuideSectionItem? selectedSection;

        public ObservableCollection<UserGuideSectionItem> Sections { get; } = new();

        public string GeneratedAtText { get; }

        public UserGuideViewModel(IEnumerable<UserGuideSectionItem> sections)
        {
            foreach (var section in sections)
            {
                Sections.Add(section);
            }

            SelectedSection = Sections.Count > 0 ? Sections[0] : null;
            GeneratedAtText = $"Hướng dẫn được tạo lúc: {DateTime.Now:dd/MM/yyyy HH:mm:ss}";
        }

        [RelayCommand]
        private void Close()
        {
            DialogResult = true;
        }
    }
}
