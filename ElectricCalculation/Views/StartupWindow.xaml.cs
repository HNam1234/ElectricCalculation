using System.Windows;
using System.Windows.Threading;
using ElectricCalculation.Services;

namespace ElectricCalculation.Views
{
    public partial class StartupWindow : Window
    {
        public StartupWindow()
        {
            InitializeComponent();
            ContentRendered += OnContentRendered;
        }

        private void OnContentRendered(object? sender, System.EventArgs e)
        {
            ContentRendered -= OnContentRendered;
            Dispatcher.BeginInvoke(() => new UiService().PreloadUserGuide(this), DispatcherPriority.ApplicationIdle);
        }
    }
}
