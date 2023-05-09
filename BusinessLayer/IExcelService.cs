using System.Collections.Generic;
using System.Data;
using Microsoft.AspNetCore.Http;

namespace BusinessLayer
{
    public interface IExcelService
    {
        DataTable MergeExcel(IFormFile mainFile, IFormFile secondFile, string primaryKey1, string primaryKey2, List<string> selectedColumns);
    }

}


