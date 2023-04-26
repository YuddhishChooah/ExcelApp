// ExcelApp/Models/ExcelDataModel.cs
using Microsoft.AspNetCore.Http;
using System.Data;

namespace ExcelApp.Models
{
    public class ExcelDataModel
    {
        public int Id { get; set; }
        public IFormFile MainFile { get; set; }
        public IFormFile SecondFile { get; set; }
        public string PrimaryKey1 { get; set; }
        public string PrimaryKey2 { get; set; }
        public string SelectedColumn { get; set; }
        public DataTable MergedData { get; set; }
    }
}
