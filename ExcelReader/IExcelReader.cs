using System;
using System.Data;
using System.Collections.Generic;

namespace ExcelReader
{
   /// <summary>
   /// Generic interface handling the transformation from Excel files to dataset.
   /// </summary>
   interface IExcelReader : IDatasetReader
   {
      /// <summary>
      /// Retrieve excel workbook and converts it to datatable
      /// </summary>
      /// <param name="WorksheetName">Worksheet name</param>
      /// <param name="ExcelBinary">Excel binary file</param>
      /// <returns>Datatable containing the data of the worksheet</returns>
      DataTable GetExcelData(string WorksheetName, byte[] ExcelBinary);
      /// <summary>
      /// Retrieve excel workbooks as a dataset.
      /// </summary>
      /// <param name="WorksheetNames">Collection of worksheet names</param>
      /// <param name="ExcelBinary">Excel binary file</param>
      /// <returns>Dataset containing the datatables of the passed worksheets</returns>
      DataSet GetExcelData(ICollection<string> WorksheetNames, byte[] ExcelBinary);
   }
}
