using ExportDataToExcel.Models;
using System.Collections.Generic;

namespace ExportDataToExcel.Services
{
    public class XFDevelopperService
    {
        public static List<XFDevelopper> GetAllXamarinDeveloppers()
        {
            List<XFDevelopper> developpers = new List<XFDevelopper>
            {
                new XFDevelopper
                {
                    ID = 1,
                    FullName = "James Montemagno",
                    Phone = "+00 0000 0001"
                },
                new XFDevelopper
                {
                    ID = 2,
                    FullName = "Leomaris Rayes",
                    Phone = "+00 0000 0002"
                },
                new XFDevelopper
                {
                    ID = 3,
                    FullName = "K. Laurent egbakou",
                    Phone = "+00 0000 0003"
                },
                new XFDevelopper
                {
                    ID = 4,
                    FullName = "Houssem Dellai",
                    Phone = "+00 0000 0004"
                },
                new XFDevelopper
                {
                    ID = 5,
                    FullName = "Houssem Dellai",
                    Phone = "+00 0000 0005"
                },
                new XFDevelopper
                {
                    ID = 6,
                    FullName = "John Doe",
                    Phone = "+00 0000 0006"
                },
                new XFDevelopper
                {
                    ID = 7,
                    FullName = "Marcel Adama",
                    Phone = "+00 0000 0007"
                },
                new XFDevelopper
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
