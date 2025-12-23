using System.Windows;

namespace ElectricCalculation.Views
{
    public partial class SaveSnapshotPromptWindow : Window
    {
        public SaveSnapshotPromptWindow()
        {
            InitializeComponent();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}

