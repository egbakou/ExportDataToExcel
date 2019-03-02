using ExportDataToExcel.Helpers;
using ExportDataToExcel.ViewModels.Base;
using System.Threading.Tasks;
using ExportDataToExcel.Services.Navigation;

namespace ExportDataToExcel.ViewModels
{
    public class BaseViewModel : ObservableObject
    {
        
        protected readonly INavigationService NavigationService;

        public BaseViewModel()
        {

            NavigationService = ViewModelLocator.Resolve<INavigationService>();
           
        }

        bool isBusy = false;
        public bool IsBusy
        {
            get { return isBusy; }
            set { SetProperty(ref isBusy, value); }
        }

        bool syncInProcess = false;
        public bool SyncInProcess
        {
            get { return syncInProcess; }
            set { SetProperty(ref syncInProcess, value); }
        }



        /// <summary>
        /// Private backing field to hold the title
        /// </summary>
        string title = string.Empty;
        /// <summary>
        /// Public property to set and get the title of the item
        /// </summary>
        public string Title
        {
            get { return title; }
            set { SetProperty(ref title, value); }
        }

        public virtual Task InitializeAsync(object navigationData)
        {
            return Task.FromResult(false);
        }

    }
}
