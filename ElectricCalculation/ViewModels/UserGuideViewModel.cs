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

    public partial class UserGuideViewModel : ObservableObject
    {
        [ObservableProperty]
        private bool? dialogResult;

        public ObservableCollection<UserGuideStepItem> Steps { get; } = new();

        public string GeneratedAtText { get; }

        public UserGuideViewModel(IEnumerable<UserGuideStepItem> steps)
        {
            foreach (var step in steps)
            {
                Steps.Add(step);
            }

            GeneratedAtText = $"Hướng dẫn được tạo lúc: {DateTime.Now:dd/MM/yyyy HH:mm:ss}";
        }

        [RelayCommand]
        private void Close()
        {
            DialogResult = true;
        }
    }
}
