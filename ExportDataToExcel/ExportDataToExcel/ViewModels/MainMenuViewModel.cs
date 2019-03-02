using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExportDataToExcel.Interfaces;
using ExportDataToExcel.Models;
using ExportDataToExcel.Services;
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
            Title = "Xamarin Developpers";
            LoadData();
            ExportToExcelCommand = new Command(ExportDataToExcelAsync);
        }

        /* Get Xamarin developpers list from Service*/
        private void LoadData()
        {
            Developpers = new ObservableCollection<XFDevelopper>(XFDevelopperService.GetAllXamarinDeveloppers());
        }


        /* Export the liste to excel file at the location*/
        public void ExportDataToExcelAsync()
        {
            if (Developpers.Count() > 0)
            {
                try
                {
                    string date = DateTime.Now.ToShortDateString();
                    date = date.Replace("/", "_");

                    FilePath = DependencyService.Get<IExportFilesToLocation>().GetFolderLocation() + "xfDeveloppers" + date + ".xlsx";

                    using (SpreadsheetDocument document = SpreadsheetDocument.Create(FilePath, SpreadsheetDocumentType.Workbook))
                    {
                        WorkbookPart workbookPart = document.AddWorkbookPart();
                        workbookPart.Workbook = new Workbook();

                        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                        worksheetPart.Worksheet = new Worksheet();

                        Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                        Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Liste des produits" };
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

                        // Inserting each product
                        foreach (var d in Developpers)
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

        private ObservableCollection<XFDevelopper> _developpers;
        public ObservableCollection<XFDevelopper> Developpers
        {
            get { return _developpers; }
            set { SetProperty(ref _developpers, value); }
        }

        private string _filePath;
        public string FilePath
        {
            get { return _filePath; }
            set { SetProperty(ref _filePath, value); }
        }

    }   
}
