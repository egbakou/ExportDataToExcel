using ExportDataToExcel.Models;
using System.Collections.Generic;

namespace ExportDataToExcel.Services
{
    public class XFDeveloperService
    {
        public static List<XFDeveloper> GetAllXamarinDevelopers()
        {
            List<XFDeveloper> developpers = new List<XFDeveloper>
            {
                new XFDeveloper
                {
                    ID = 1,
                    FullName = "James Montemagno",
                    Phone = "+00 0000 0001"
                },
                new XFDeveloper
                {
                    ID = 2,
                    FullName = "Leomaris Rayes",
                    Phone = "+00 0000 0002"
                },
                new XFDeveloper
                {
                    ID = 3,
                    FullName = "K. Laurent Egbakou",
                    Phone = "+00 0000 0003"
                },
                new XFDeveloper
                {
                    ID = 4,
                    FullName = "Houssem Dellai",
                    Phone = "+00 0000 0004"
                },
                new XFDeveloper
                {
                    ID = 5,
                    FullName = "Yves Gaston",
                    Phone = "+00 0000 0005"
                },
                new XFDeveloper
                {
                    ID = 6,
                    FullName = "John Doe",
                    Phone = "+00 0000 0006"
                },
                new XFDeveloper
                {
                    ID = 7,
                    FullName = "Marcel Adama",
                    Phone = "+00 0000 0007"
                },
                new XFDeveloper
                {
                    ID = 8,
                    FullName = "Carlos Ognankotan",
                    Phone = "+00 0000 0008"
                }
            };

            return developpers;
        }                        
    }
}
