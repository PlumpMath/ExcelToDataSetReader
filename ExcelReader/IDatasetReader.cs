namespace ExcelReader
{
   using System.Data;

   /// <summary>
   /// Generic interface handling the transformation from generic files to datasets.
   /// </summary>
   public interface IDataSetReader
   {  
      /// <summary>
      /// Get a dataset out of a file.
      /// </summary>
      /// <param name="file">Generic file containing data that can be translated to a dataset</param>
      /// <returns>Translated dataset</returns>
      DataSet GetDataSet(byte[] file);
   }
}
