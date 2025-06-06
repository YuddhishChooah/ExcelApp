﻿using System;
using System.Data;
using System.IO;
using System.Linq;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml;

namespace BusinessLayer
{
    public class ExcelService : IExcelService
    {
        public ExcelService() { }

        public DataTable MergeExcel(IFormFile mainFile, IFormFile secondFile, string primaryKey1, string primaryKey2, List<string> selectedColumns)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            DataTable mergedData;


            using (var mainStream = new MemoryStream())
            using (var secondStream = new MemoryStream())
            {
                mainFile.CopyTo(mainStream);
                secondFile.CopyTo(secondStream);

                using (ExcelPackage mainExcelPackage = new ExcelPackage(mainStream))
                using (ExcelPackage secondExcelPackage = new ExcelPackage(secondStream))
                {
                    ExcelWorksheet mainWorksheet = mainExcelPackage.Workbook.Worksheets["Issue List KA"];
                    ExcelWorksheet secondWorksheet = secondExcelPackage.Workbook.Worksheets[0];

                    var mainDataTable = WorksheetToDataTable(mainWorksheet);
                    var secondDataTable = WorksheetToDataTable(secondWorksheet);

                    List<string> selectedColumnsList = selectedColumns;

                    mergedData = MergeDataTables(mainDataTable, secondDataTable, primaryKey1, primaryKey2, selectedColumnsList);
                }
            }

            return mergedData;
        }

        private DataTable WorksheetToDataTable(ExcelWorksheet worksheet)
        {
            DataTable dt = new DataTable();

            foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
            {
                dt.Columns.Add(firstRowCell.Text);
            }

            for (int rowNum = 2; rowNum <= worksheet.Dimension.End.Row; rowNum++)
            {
                var wsRow = worksheet.Cells[rowNum, 1, rowNum, worksheet.Dimension.End.Column];
                DataRow row = dt.Rows.Add();
                foreach (var cell in wsRow)
                {
                    if (cell.Start.Column - 1 >= 0 && cell.Start.Column - 1 < row.ItemArray.Length)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                    else
                    {
                        // Handle the case when the index is out of range
                        // For example, log a warning or throw a custom exception
                    }
                }
            }

            return dt;
        }

        private DataTable MergeDataTables(DataTable mainDataTable, DataTable secondDataTable, string primaryKey1, string primaryKey2, List<string> selectedColumns)
        {
            if (mainDataTable == null || secondDataTable == null)
                throw new ArgumentNullException("Both DataTables should not be null.");

            if (!mainDataTable.Columns.Contains(primaryKey1) || !secondDataTable.Columns.Contains(primaryKey2))
                throw new ArgumentException("The primary key columns must be present in both DataTables.");

            foreach (var selectedColumn in selectedColumns)
            {
                if (!mainDataTable.Columns.Contains(selectedColumn) || !secondDataTable.Columns.Contains(selectedColumn))
                    throw new ArgumentException($"The selected column '{selectedColumn}' must be present in both DataTables.");
            }

            DataTable mergedDataTable = mainDataTable.Copy();
            var secondTableRows = secondDataTable.AsEnumerable();

            foreach (DataRow mainRow in mergedDataTable.Rows)
            {
                var matchingRow = secondTableRows.FirstOrDefault(secondRow => secondRow[primaryKey2].ToString() == mainRow[primaryKey1].ToString());

                if (matchingRow != null)
                {
                    foreach (var selectedColumn in selectedColumns)
                    {
                        mainRow[selectedColumn] = matchingRow[selectedColumn];
                    }
                }
            }

            return mergedDataTable;
        }


    }
}