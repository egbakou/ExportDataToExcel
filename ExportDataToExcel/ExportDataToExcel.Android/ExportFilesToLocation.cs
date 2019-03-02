using ExportDataToExcel.Droid;
using ExportDataToExcel.Interfaces;
using Java.IO;
using Xamarin.Forms;

[assembly: Dependency(typeof(ExportFilesToLocation))]
namespace ExportDataToExcel.Droid
{
    public class ExportFilesToLocation : IExportFilesToLocation
    {
        public ExportFilesToLocation()
        {
        }

        public string GetFolderLocation()
        {
            string root = null;
            if (Android.OS.Environment.IsExternalStorageEmulated)
            {
                root = Android.OS.Environment.ExternalStorageDirectory.AbsolutePath;
            }
            else
                root = System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments);

            File myDir = new File(root + "/Exports");
            if(!myDir.Exists())
                myDir.Mkdir();

            return root + "/Exports/";
        }
    }
}