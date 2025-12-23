using System.Windows;

namespace ElectricCalculation.Views
{
    public partial class SaveSnapshotWindow : Window
    {
        public SaveSnapshotWindow()
        {
            InitializeComponent();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
