using System.Windows;
using System.Windows.Threading;
using ElectricCalculation.Services;
using ElectricCalculation.ViewModels;

namespace ElectricCalculation.Views
{
    public partial class StartupWindow : Window
    {
        public StartupWindow()
        {
            InitializeComponent();
            ContentRendered += OnContentRendered;
        }

        private async void OnContentRendered(object? sender, System.EventArgs e)
        {
            ContentRendered -= OnContentRendered;

            if (DataContext is StartupViewModel vm)
            {
                vm.LoadingResourcesText = "Loading resources…";
                vm.IsLoadingResources = true;
            }

            await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.Render);
            await Dispatcher.InvokeAsync(() => new UiService().PreloadUserGuide(this), DispatcherPriority.ApplicationIdle);

            if (DataContext is StartupViewModel vmDone)
            {
                vmDone.IsLoadingResources = false;
            }
        }
    }
}
