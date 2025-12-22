using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace ElectricCalculation.ViewModels
{
    public partial class MessageDialogViewModel : ObservableObject
    {
        [ObservableProperty]
        private bool? dialogResult;

        [ObservableProperty]
        private string title = string.Empty;

        [ObservableProperty]
        private string message = string.Empty;

        public MessageDialogViewModel()
        {
        }

        public MessageDialogViewModel(string title, string message)
        {
            this.title = title;
            this.message = message;
        }

        [RelayCommand]
        private void Ok()
        {
            DialogResult = true;
        }
    }
}

