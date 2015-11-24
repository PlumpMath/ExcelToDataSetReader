using Microsoft.Office.Interop.Excel;

namespace ExcelReader
{
   using System;
   using System.Collections.Generic;
   using System.Linq;
   using System.Text;
   using System.Data;

   /// <summary>
   /// Object containins logic for reading excel files using Excel API.
   /// </summary>
   public class ExcelToDataSetReader : ExcelReaderBase, IExcelReader
   {
      /// <summary>
      /// Default excel first sheet name.
      /// </summary>
      public const string DefaultExcelSheetName = "Sheet1";

      /// <summary>
      /// Initialize the base class, without keeping the generated files
      /// </summary>
      public ExcelToDataSetReader()
         : base(false)
      {
      }

      /// <summary>
      /// Retrieve excel data and convert it to dataset
      /// </summary>
      /// <param name="file">Excel binary file.</param>
      /// <returns>Dataset containing the data of the sheet pertaining to the position specified</returns>
      public DataSet GetDataSet(byte[] file)
      {
         try
         {
            //try to create an excel application
            CreateExcelApplication();
            KeepFiles = false;
            OpenWorkbook(file);
            List<string> sheets = new List<string>();
            foreach (Worksheet s in Application.Worksheets)
            {
               //ignore hidden sheets
               if (s.Visible != XlSheetVisibility.xlSheetVisible)
               {
                  continue;
               }
               sheets.Add(s.Name);
               //release com obj
               ReleaseComObject(s);
            }
            //read respective sheets
            DataSet ds = new DataSet();
            foreach (string s in sheets)
            {
               ds.Tables.Add(GetExcelData(s));
            }
            ds.AcceptChanges();
            return ds;
         }
         finally
         {
            this.Dispose();
         }
      }

      /// <summary>
      /// Retrieve excel workbooks as a dataset.
      /// </summary>
      /// <param name="worksheetNames">Collection of worksheet names</param>
      /// <param name="file">Excel binary file</param>
      /// <returns>Dataset containing the datatables of the passed worksheets</returns>
      public DataSet GetExcelData(ICollection<string> worksheetNames, byte[] file)
      {
         //get the whole workbook content
         DataSet ds = GetDataSet(file);

         DataSet dsRes = new DataSet();
         foreach (System.Data.DataTable t in ds.Tables)
         {
            if (worksheetNames.Contains(t.TableName))
            {
               System.Data.DataTable Table = t.Copy();
               dsRes.Tables.Add(Table);
            }
         }
         return dsRes;
      }

      /// <summary>
      /// Retrieve excel workbook and converts it to datatable
      /// </summary>
      /// <param name="worksheetNames">Worksheet name</param>
      /// <param name="file">Excel binary file</param>
      /// <returns>Datatable containing the data of the workbook</returns>
      public DataTable GetExcelData(string worksheetNames, byte[] file)
      {
         try
         {
            //try to create an excel application
            this.CreateExcelApplication();
            this.KeepFiles = false;
            this.OpenWorkbook(file);

            //read respective sheets
            DataTable wbTable = new DataTable();
            wbTable = GetExcelData(worksheetNames);
            wbTable.AcceptChanges();
            return wbTable;
         }
         finally
         {
            this.Dispose();
         }
      }

      /// <summary>
      /// Gets excel data from the application where workbook is opened.
      /// </summary>
      /// <param name="sheet">Sheet to get the data</param>
      /// <returns>Datatable containins workbook sheet data</returns>
      private DataTable GetExcelData(string sheet)
      {
         //top corner/first column for each excel sheet
         const string TopAbsoluteCorner = "A1";
         const string AbsoluteFirstColumn = "A";

         var oS = (Worksheet)Application.Worksheets[sheet];
         var oR = oS.UsedRange;

         //creating table
         System.Data.DataTable t = new System.Data.DataTable(oS.Name);

         try
         {
            //get the address of the bottom, right cell
            string downAddress = oR.get_Address(false, false, XlReferenceStyle.xlA1);

            //get the full range
            oR = oS.get_Range(TopAbsoluteCorner, downAddress);
            object[,] sheetMatrix = oR.Value as object[,];

            if (sheetMatrix != null)
            {
               for (int colIdx = 0; colIdx <= sheetMatrix.GetUpperBound(1) - 1; colIdx++)
               {
                  DataColumn Column = new DataColumn(GetColumnNameFromIndex(colIdx), typeof(object));
                  t.Columns.Add(Column);
               }

               for (int rowIdx = 0; rowIdx <= sheetMatrix.GetUpperBound(0) - 1; rowIdx++)
               {
                  DataRow Row = t.NewRow();

                  for (int cellIdx = 0; cellIdx <= sheetMatrix.GetUpperBound(1) - 1; cellIdx++)
                  {
                     Row[cellIdx] = sheetMatrix[rowIdx + 1, cellIdx + 1];

                  }
                  t.Rows.Add(Row);
               }
            }
            else
            {
               //only one cell has contents
               DataColumn col = new DataColumn(AbsoluteFirstColumn, typeof(object));
               t.Columns.Add(col);
               DataRow Row = t.NewRow();
               Row[AbsoluteFirstColumn] = oR.Value;
               t.Rows.Add(Row);
            }

            return t;

         }
         finally
         {
            //release all COM components for the used object sheet and range
            ReleaseComObject(oR);
            ReleaseComObject(oS);
         }
      }

      /// <summary>
      /// Gets the Excel column name from a given absolute index.
      /// </summary>
      /// <param name="i">Absolute index</param>
      /// <returns>Excel column name</returns>
      static internal string GetColumnNameFromIndex(int i)
      {
         char[] baseChars = new char[] {
         'A',
         'B',
         'C',
         'D',
         'E',
         'F',
         'G',
         'H',
         'I',
         'J',
         'K',
         'L',
         'M',
         'N',
         'O',
         'P',
         'Q',
         'R',
         'S',
         'T',
         'U',
         'V',
         'W',
         'X',
         'Y',
         'Z'};
         string name = string.Empty;
         int targetBase = baseChars.Length;
         int correction = 0;
         // This is necessary since Excel column naming consider is A,..,Z,AA,..,AZ,BA,... which is
         // equivalent, in decimal to 0,...,9,00,..,09,10,...,19. If this is not used the columns name would be
         // A,..,Z,BA,...,BZ,CA,..
         do
         {
            name = baseChars[(i - correction) % targetBase] + name;
            i = Convert.ToInt32(Math.Floor((double)i / targetBase));
            correction = 1;
         } while ((i > 0));

         return name;
      }
   }
}
