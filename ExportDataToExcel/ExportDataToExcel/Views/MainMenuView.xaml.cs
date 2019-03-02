using ExportDataToExcel.ViewModels;

using Xamarin.Forms;
using Xamarin.Forms.Xaml;

namespace ExportDataToExcel.Views
{
    [XamlCompilation(XamlCompilationOptions.Compile)]
    public partial class MainMenuView : ContentPage
    {
        private MainMenuViewModel viewModel;
        public MainMenuView()
        {
            InitializeComponent();
            BindingContext = viewModel = new MainMenuViewModel();
        }
     
    }
}