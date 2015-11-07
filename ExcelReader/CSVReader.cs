using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace ExcelReader
{
   /// <summary>
   /// Contains logic for reading CSV files and translating them into a dataset representing the file opened in Excel.
   /// </summary>
   public class CSVReader : IDatasetReader
   {

      /// <summary>
      /// Default RFC-4180 separator
      /// </summary>
      public const char DefRFC4180Sep = ',';

      /// <summary>
      /// Reads a CSV file and translates that to a dataset.
      /// Column names reflect the Excel columns.
      /// </summary>
      /// <param name="FileData">CSV file to read</param>
      /// <returns>Dataset</returns>
      /// <remarks>CSV file can be read if specified as in RFC-4180, with the additional formatting of European CSV (semi-column separator instead of comma)</remarks>
      public DataSet GetDataSet(byte[] FileData)
      {
         char[] PossibleSeparator = new char[] {
         DefRFC4180Sep,
         ';',
         '\t'
      };
         //US format, European format and tab
         DataSet ExcelDataset = new DataSet();

         DataTable CSVTable = new DataTable(ExcelReader.DefaultExcelSheetName);
         Encoding Encoding = GetSafeEncodingFromBOM(FileData);
         string CSVFile = Encoding.GetString(FileData);

         //get rows (vbCrLF is the standard used by CSV RFC)
         string[] CSVRows = CSVFile.Split(Environment.NewLine.ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

         //check separator and replace
         char CSVSep = DetectCSVSeparator(CSVRows, PossibleSeparator, DefRFC4180Sep);

         foreach (string rowsWithEnd in (from o in CSVRows select o + CSVSep))
         {
            List<string> values = new List<string>();
            List<char> currValue = new List<char>();
            bool skip = false;

            char[] chars = rowsWithEnd.ToCharArray();
            for (int i = 0; i <= chars.Length - 1; i++)
            {
               char current = chars[i];

               if (current == CSVSep)
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
                  if (!skip && i + 1 < chars.Length - 1 && chars[i + 1] == CSVSep)
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
            if (CSVTable.Columns.Count < values.Count)
            {
               int upperBound = values.Count - CSVTable.Columns.Count - 1;
               for (int colIdx = CSVTable.Columns.Count; colIdx <= upperBound; colIdx++)
               {
                  CSVTable.Columns.Add(ExcelReader.GetColumnNameFromIndex(colIdx), typeof(System.Object));
               }
            }
            //add row
            CSVTable.Rows.Add(values.ToArray());
         }
         ExcelDataset.Tables.Add(CSVTable);

         return ExcelDataset;
      }

      /// <summary>
      /// Safely get encoding reading the BOM (Byte Order Mark) or preamble of a byte array.
      /// </summary>
      /// <param name="file">File byte array</param>
      /// <returns>Default Encoding <v>ANSI</v> if no BOM found, <v>Encoding</v> otherwise</returns>
      private Encoding GetSafeEncodingFromBOM(byte[] file)
      {
         foreach (EncodingInfo EncodingInfo in Encoding.GetEncodings())
         {
            Encoding LookupEncoding = EncodingInfo.GetEncoding();
            byte[] BOM = LookupEncoding.GetPreamble();
            bool eureka = true;
            if ((BOM.Length > 0) && (BOM.Length <= file.Length))
            {
               for (int bomIdx = 0; bomIdx <= BOM.Length - 1; bomIdx++)
               {
                  if (BOM[bomIdx] != file[bomIdx])
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
               return LookupEncoding;
            }
         }

         //return default encoding if nothing has been found
         return Encoding.Default;
      }

      /// <summary>
      /// Detects the CSV separator from a set of possible separators.
      /// </summary>
      /// <param name="CSVRows">CSV rows as set of strings</param>
      /// <param name="defSep">Default value to return if no other separators are detected</param>
      /// <returns>Separator, if detected, <paramref name="defSep">Default separator</paramref> if nothing else has been found</returns>
      private char DetectCSVSeparator(string[] CSVRows, IList<char> Separators, char defSep)
      {
         //copy the list of candidate separators
         //ControlChars.Tab
         if (Separators.Count == 0 || CSVRows.Length == 0)
         {
            return defSep;
         }

         List<char> Sep = (from chSep in Separators select chSep).Distinct().ToList();
         //read the first row and exclude separators from the array
         for (int sepIdx = Separators.Count - 1; sepIdx >= 0; sepIdx += -1)
         {
            if (!CSVRows[0].Contains(Sep[sepIdx]))
            {
               Sep.RemoveAt(sepIdx);
            }
         }

         //optimize, return separator if just one has been found
         if (Sep.Count == 1)
         {
            return Sep[0];
         }

         //initialize a counter for separators 
         //count the occurrences and compare them with the previous row
         int[,] sepCounts = new int[Sep.Count, CSVRows.Length];
         for (int i = 0; i <= CSVRows.Length - 1; i++)
         {
            char[] CharArray = CSVRows[i].ToCharArray();
            for (int j = 0; j <= Sep.Count - 1; j++)
            {
               bool skip = false;
               for (int iCharIndex = 0; iCharIndex <= CharArray.Length - 1; iCharIndex++)
               {
                  char current = CharArray[iCharIndex];
                  if (current == '"')
                  {
                     skip = !skip;
                  }
                  else if (current == Sep[j])
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
               for (int k = 0; k <= Sep.Count - 1; k++)
               {
                  if (sepCounts[k, i - 1] == sepCounts[k, i])
                  {
                     return Sep[k];
                  }
               }
            }
         }
         return defSep;
      }
   }
}
