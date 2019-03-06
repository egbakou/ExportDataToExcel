using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExportDataToExcel.Interfaces;
using ExportDataToExcel.Models;
using ExportDataToExcel.Services;
using Plugin.Permissions;
using Plugin.Permissions.Abstractions;
using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Windows.Input;
using Xamarin.Forms;
using Cell = DocumentFormat.OpenXml.Spreadsheet.Cell;

namespace ExportDataToExcel.ViewModels
{
    public class MainMenuViewModel : BaseViewModel
    {
       
        public MainMenuViewModel()
        {
            Title = "Xamarin Developers";
            LoadData();
            ExportToExcelCommand = new Command(async () => await ExportDataToExcelAsync());
        }

        /* Get Xamarin developers list from Service*/
        private void LoadData()
        {
            Developers = new ObservableCollection<XFDeveloper>(XFDeveloperService.GetAllXamarinDevelopers());
        }


        /* Export the list to excel file at the location provide by DependencyService */
        public async System.Threading.Tasks.Task ExportDataToExcelAsync()
        {
            // Granted storage permission
            var storageStatus = await CrossPermissions.Current.CheckPermissionStatusAsync(Permission.Storage);

            if (storageStatus != PermissionStatus.Granted)
            {
                var results = await CrossPermissions.Current.RequestPermissionsAsync(new[] { Permission.Storage });
                storageStatus = results[Permission.Storage];
            }

            if (Developers.Count() > 0)
            {
                try
                {
                    string date = DateTime.Now.ToShortDateString();
                    date = date.Replace("/", "_");

                    var path = DependencyService.Get<IExportFilesToLocation>().GetFolderLocation() + "XFDevelopers" + date + ".xlsx";
                    FilePath = path;
                    using (SpreadsheetDocument document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
                    {
                        WorkbookPart workbookPart = document.AddWorkbookPart();
                        workbookPart.Workbook = new Workbook();

                        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                        worksheetPart.Worksheet = new Worksheet();

                        Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                        Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Xamarin Forms developers list" };
                        sheets.Append(sheet);

                        workbookPart.Workbook.Save();

                        SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

                        // Constructing header
                        Row row = new Row();

                        row.Append(
                            ConstructCell("No", CellValues.String),
                            ConstructCell("FullName", CellValues.String),
                            ConstructCell("Phone", CellValues.String)
                            );

                        // Insert the header row to the Sheet Data
                        sheetData.AppendChild(row);

                        // Add each product
                        foreach (var d in Developers)
                        {
                            row = new Row();
                            row.Append(
                                ConstructCell(d.ID.ToString(), CellValues.String),
                                ConstructCell(d.FullName, CellValues.String),
                                ConstructCell(d.Phone, CellValues.String));
                            sheetData.AppendChild(row);
                        }

                        worksheetPart.Worksheet.Save();
                        MessagingCenter.Send(this, "DataExportedSuccessfully");
                    }

                }
                catch (Exception e)
                {
                    Debug.WriteLine("ERROR: "+ e.Message);
                }
            }
            else
            {
                MessagingCenter.Send(this, "NoDataToExport");
            }
        }


        /* To create cell in Excel */
        private Cell ConstructCell(string value, CellValues dataType)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType)
            };
        }


        public ICommand ExportToExcelCommand { get; set; }

        private ObservableCollection<XFDeveloper> _developers;
        public ObservableCollection<XFDeveloper> Developers
        {
            get { return _developers; }
            set { SetProperty(ref _developers, value); }
        }

        private string _filePath;
        public string FilePath
        {
            get { return _filePath; }
            set { SetProperty(ref _filePath, value); }
        }

    }   
}
