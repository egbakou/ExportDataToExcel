using ExportDataToExcel.Interfaces;
using ExportDataToExcel.iOS;
using Xamarin.Forms;

[assembly: Dependency(typeof(ExportFilesToLocation))]
namespace ExportDataToExcel.iOS
{
    public class ExportFilesToLocation : IExportFilesToLocation
    {
        public ExportFilesToLocation()
        {
        }

        public string GetFolderLocation()
        {
            string root = System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments);
            return root + "/";
        }
    }
}