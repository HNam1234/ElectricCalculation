using CommunityToolkit.Mvvm.ComponentModel;

namespace ElectricCalculation.ViewModels
{
    public partial class BusyDialogViewModel : ObservableObject
    {
        [ObservableProperty]
        private string title = string.Empty;

        [ObservableProperty]
        private string message = string.Empty;

        public BusyDialogViewModel()
        {
        }

        public BusyDialogViewModel(string title, string message)
        {
            this.title = title;
            this.message = message;
        }
    }
}

