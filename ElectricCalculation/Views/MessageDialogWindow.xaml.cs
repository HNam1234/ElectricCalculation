using System.Windows;

namespace ElectricCalculation.Views
{
    public partial class MessageDialogWindow : Window
    {
        public MessageDialogWindow()
        {
            InitializeComponent();
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }
    }
}

