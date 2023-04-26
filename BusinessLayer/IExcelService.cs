using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using System.Data;
using System.IO;

namespace BusinessLayer
{
    public interface IExcelService
    {
        DataTable MergeExcel(IFormFile mainFile, IFormFile secondFile, string primaryKey1, string primaryKey2, string selectedColumn);
    }
}