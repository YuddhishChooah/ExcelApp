// ExcelApp/Models/ExcelDataModel.cs
using Microsoft.AspNetCore.Http;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Xml.Linq;

namespace ExcelApp.Models
{
    public class ExcelDataModel
    {
        [Display(Name = "Primary Key 1")]
        public string PrimaryKey1 { get; set; }

        [Display(Name = "Primary Key 2")]
        public string PrimaryKey2 { get; set; }

        [Display(Name = "Selected Columns (comma-separated)")]
        public string SelectedColumns { get; set; }

        public DataTable MergedData { get; set; }

        public List<string> GetSelectedColumnsList()
        {
            if (string.IsNullOrEmpty(SelectedColumns))
                return new List<string>();

            return new List<string>(SelectedColumns.Split(',').Select(s => s.Trim()));
        }
    }
}

