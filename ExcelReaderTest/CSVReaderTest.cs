using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Data;
using System.Linq;

namespace ExcelReaderTest
{
   [TestClass]
   public class CSVReaderTest
   {
      private const string Sheet = ExcelReader.ExcelReader.DefaultExcelSheetName;

      [TestMethod]
      public void Test_CSVReader_GetsDataset()
      {
         ExcelReader.IDatasetReader E = new ExcelReader.CSVReader();
         var file = ExcelReaderTest.Properties.Resources.TestCSV;
         DataSet ds = E.GetDataSet(file);
         Assert.IsNotNull(ds);
      }

      [TestMethod]
      public void Test_CSVReader_GetsDatatable()
      {
         ExcelReader.IDatasetReader E = new ExcelReader.CSVReader();
         var file = ExcelReaderTest.Properties.Resources.TestCSV;
         DataSet ds = E.GetDataSet(file);
         Int32 expectedNoTables = 1;
         Assert.AreEqual<Int32>(expectedNoTables, ds.Tables.Count);
         Assert.IsTrue(ds.Tables.Contains(Sheet));
      }

      [TestMethod]
      public void Test_CSVReader_GetsDatatable_RowsCount()
      {
         ExcelReader.IDatasetReader E = new ExcelReader.CSVReader();
         var file = ExcelReaderTest.Properties.Resources.TestCSV;
         DataSet ds = E.GetDataSet(file);
         DataTable dt = ds.Tables[Sheet];
         Int32 expectedRowCount = 8;
         Assert.AreEqual<Int32>(expectedRowCount, dt.Rows.Count);
      }

      [TestMethod]
      public void Test_CSVReader_GetsDatatable_ColumnCount()
      {
         ExcelReader.IDatasetReader E = new ExcelReader.CSVReader();
         var file = ExcelReaderTest.Properties.Resources.TestCSV;
         DataSet ds = E.GetDataSet(file);
         DataTable dt = ds.Tables[Sheet];
         Int32 expectedColumnCount = 5;
         Assert.AreEqual<Int32>(expectedColumnCount, dt.Columns.Count);
      }

      [TestMethod]
      public void Test_CSVReader_GetsColumnNames()
      {
         ExcelReader.IDatasetReader E = new ExcelReader.CSVReader();
         var file = ExcelReaderTest.Properties.Resources.TestCSV;
         DataSet ds = E.GetDataSet(file);
         DataTable dt = ds.Tables[Sheet];
         string[] expectedColumnNames = new string[] { "A", "B", "C", "D", "E" };
         string[] actualColumnNames = (from DataColumn col in dt.Columns select col.ColumnName).ToArray();
         Assert.IsTrue(expectedColumnNames.SequenceEqual(actualColumnNames));
      }

      [TestMethod]
      public void Test_CSVReader_GetsData()
      {
         ExcelReader.IDatasetReader E = new ExcelReader.CSVReader();
         var file = ExcelReaderTest.Properties.Resources.TestCSV;
         DataSet ds = E.GetDataSet(file);
         DataTable dt = ds.Tables[0];
         Assert.AreEqual<string>("A1", dt.Rows[0].Field<string>(0));
         Assert.AreEqual<string>("A4", dt.Rows[3].Field<string>(0));
         Assert.AreEqual<string>("A8", dt.Rows[7].Field<string>(0));
         Assert.AreEqual<string>("C4", dt.Rows[3].Field<string>(2));
         Assert.AreEqual<string>("C5", dt.Rows[4].Field<string>(2));
         Assert.AreEqual<string>("C6", dt.Rows[5].Field<string>(2));
         Assert.AreEqual<string>("C7", dt.Rows[6].Field<string>(2));
         Assert.AreEqual<string>("D1", dt.Rows[0].Field<string>(3));
         Assert.AreEqual<string>("4", dt.Rows[3].Field<string>(3));
         Assert.AreEqual<string>("5", dt.Rows[4].Field<string>(3));
         Assert.AreEqual<string>("6", dt.Rows[5].Field<string>(3));
         Assert.AreEqual<string>("7", dt.Rows[6].Field<string>(3));
         Assert.AreEqual<string>("8", dt.Rows[7].Field<string>(3));
         Assert.AreEqual<string>("E1", dt.Rows[0].Field<string>(4));
         Assert.AreEqual<string>("E4", dt.Rows[1].Field<string>(4));
         Assert.AreEqual<string>("E3", dt.Rows[2].Field<string>(4));
      }

      [TestMethod]
      public void Test_CSVReader_GetsEmptyCells()
      {
         ExcelReader.IDatasetReader E = new ExcelReader.CSVReader();
         var file = ExcelReaderTest.Properties.Resources.TestCSV;
         DataSet ds = E.GetDataSet(file);
         DataTable dt = ds.Tables[0];
         Int32 expectedEmptyCellsCount = 24;
         int nullValuesCount = 0;
         for (int i = 0; i < 8; i++)
            for (int j = 0; j < 5; j++)
               if (dt.Rows[i][j].ToString().Length == 0)
                  nullValuesCount++;
         Assert.AreEqual<Int32>(expectedEmptyCellsCount, nullValuesCount);
      }
   }
}
