namespace ExcelReader
{
   using System;
   using System.Collections.Generic;
   using System.Linq;
   using System.Text;
   using System.Data;

   /// <summary>
   /// Contains logic for reading CSV files and translating them into a dataset representing the file opened in Excel.
   /// </summary>
   public class CsvReader : IDataSetReader
   {

      /// <summary>
      /// Default RFC-4180 separator
      /// </summary>
      public const char DefRfc4180Sep = ',';

      /// <summary>
      /// Reads a CSV file and translates that to a dataset.
      /// Column names reflect the Excel columns.
      /// </summary>
      /// <param name="file">CSV file to read</param>
      /// <returns>Dataset</returns>
      /// <remarks>CSV file can be read if specified as in RFC-4180, with the additional formatting of European CSV (semi-column separator instead of comma)</remarks>
      public DataSet GetDataSet(byte[] file)
      {
         char[] candidateSeparators = new char[] { DefRfc4180Sep, ';', '\t' };
         //US format, European format and tab
         DataSet ds = new DataSet();

         DataTable CsvTable = new DataTable(ExcelToDataSetReader.DefaultExcelSheetName);
         Encoding enc = GetSafeEncodingFromBom(file);
         string CSVFile = enc.GetString(file);

         //get rows (vbCrLF is the standard used by CSV RFC)
         string[] CsvRows = CSVFile.Split(Environment.NewLine.ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

         //check separator and replace
         char CsvSep = DetectCsvSeparator(CsvRows, candidateSeparators, DefRfc4180Sep);

         foreach (string rowsWithEnd in (from o in CsvRows select o + CsvSep))
         {
            List<string> values = new List<string>();
            List<char> currValue = new List<char>();
            bool skip = false;

            char[] chars = rowsWithEnd.ToCharArray();
            for (int i = 0; i <= chars.Length - 1; i++)
            {
               char current = chars[i];

               if (current == CsvSep)
               {
                  if (!skip)
                  {
                     values.Add(new string(currValue.ToArray()));
                     currValue.Clear();
                  }
                  else
                  {
                     currValue.Add(chars[i]);
                  }
               }
               else if (current == '"')
               {
                  skip = !skip;
                  if (!skip && i + 1 < chars.Length - 1 && chars[i + 1] == CsvSep)
                  {
                     continue;
                  }
                  if (i > 0 && chars[i - 1] == '"')
                  {
                     currValue.Add('"');
                  }
               }
               else
               {
                  currValue.Add(chars[i]);
               }
            }
            //creating the columns
            if (CsvTable.Columns.Count < values.Count)
            {
               int upperBound = values.Count - CsvTable.Columns.Count - 1;
               for (int colIdx = CsvTable.Columns.Count; colIdx <= upperBound; colIdx++)
               {
                  CsvTable.Columns.Add(ExcelToDataSetReader.GetColumnNameFromIndex(colIdx), typeof(System.Object));
               }
            }
            //add row
            CsvTable.Rows.Add(values.ToArray());
         }
         ds.Tables.Add(CsvTable);

         return ds;
      }

      /// <summary>
      /// Safely get encoding reading the BOM (Byte Order Mark) or preamble of a byte array.
      /// </summary>
      /// <param name="file">File byte array</param>
      /// <returns>Default Encoding <v>ANSI</v> if no BOM found, <v>Encoding</v> otherwise</returns>
      private static Encoding GetSafeEncodingFromBom(byte[] file)
      {
         foreach (EncodingInfo EncodingInfo in Encoding.GetEncodings())
         {
            Encoding lookupEnc = EncodingInfo.GetEncoding();
            byte[] bom = lookupEnc.GetPreamble();
            bool eureka = true;
            if ((bom.Length > 0) && (bom.Length <= file.Length))
            {
               for (int bomIdx = 0; bomIdx <= bom.Length - 1; bomIdx++)
               {
                  if (bom[bomIdx] != file[bomIdx])
                  {
                     eureka = false;
                     break;
                  }
               }
            }
            else
            {
               eureka = false;
            }
            if (eureka)
            {
               return lookupEnc;
            }
         }

         //return default encoding if nothing has been found
         return Encoding.Default;
      }

      /// <summary>
      /// Detects the CSV separator from a set of possible separators.
      /// </summary>
      /// <param name="CsvRows">CSV rows as set of strings</param>
      /// <param name="candidateSeparators">Hint of candidate separators</param>
      /// <param name="defSep">Default value to return if no other separators are detected</param>
      /// <returns>Separator, if detected, <paramref name="defSep">Default separator</paramref> if nothing else has been found</returns>
      private static char DetectCsvSeparator(string[] CsvRows, IList<char> candidateSeparators, char defSep)
      {
         //copy the list of candidate separators
         //ControlChars.Tab
         if (candidateSeparators.Count == 0 || CsvRows.Length == 0)
         {
            return defSep;
         }

         List<char> seps = (from chSep in candidateSeparators select chSep).Distinct().ToList();
         //read the first row and exclude separators from the array
         for (int sepIdx = candidateSeparators.Count - 1; sepIdx >= 0; sepIdx += -1)
         {
            if (!CsvRows[0].Contains(seps[sepIdx]))
            {
               seps.RemoveAt(sepIdx);
            }
         }

         //optimize, return separator if just one has been found
         if (seps.Count == 1)
         {
            return seps[0];
         }

         //initialize a counter for separators 
         //count the occurrences and compare them with the previous row
         int[,] sepCounts = new int[seps.Count, CsvRows.Length];
         for (int i = 0; i <= CsvRows.Length - 1; i++)
         {
            char[] arr = CsvRows[i].ToCharArray();
            for (int j = 0; j <= seps.Count - 1; j++)
            {
               bool skip = false;
               for (int iCharIndex = 0; iCharIndex <= arr.Length - 1; iCharIndex++)
               {
                  char current = arr[iCharIndex];
                  if (current == '"')
                  {
                     skip = !skip;
                  }
                  else if (current == seps[j])
                  {
                     if (!skip)
                     {
                        sepCounts[j, i] += 1;
                     }
                  }
               }
            }
            if (i > 0)
            {
               for (int k = 0; k <= seps.Count - 1; k++)
               {
                  if (sepCounts[k, i - 1] == sepCounts[k, i])
                  {
                     return seps[k];
                  }
               }
            }
         }
         return defSep;
      }
   }
}
