namespace ExcelReader
{
   using System.Data;
   using System.Collections.Generic;

   /// <summary>
   /// Generic interface handling the transformation from Excel files to dataset.
   /// </summary>
   interface IExcelReader : IDataSetReader
   {
      /// <summary>
      /// Retrieve excel workbook and converts it to datatable
      /// </summary>
      /// <param name="worksheetName">Worksheet name</param>
      /// <param name="file">Excel binary file</param>
      /// <returns>Datatable containing the data of the worksheet</returns>
      DataTable GetExcelData(string worksheetName, byte[] file);
      /// <summary>
      /// Retrieve excel workbooks as a dataset.
      /// </summary>
      /// <param name="worksheetNames">Collection of worksheet names</param>
      /// <param name="file">Excel binary file</param>
      /// <returns>Dataset containing the datatables of the passed worksheets</returns>
      DataSet GetExcelData(ICollection<string> worksheetNames, byte[] file);
   }
}
