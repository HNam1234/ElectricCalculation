using CommunityToolkit.Mvvm.ComponentModel;

namespace ElectricCalculation.ViewModels
{
    public partial class MessageDialogViewModel : ObservableObject
    {
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
    }
}

