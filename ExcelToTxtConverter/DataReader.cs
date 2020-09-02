using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace ExcelToTxtConverter
{
    public class DataReader
    {
        public DataTable Execute(byte[] excelData)
        {
            try
            {
                var tempFilenamePath = Path.GetTempFileName();
                File.WriteAllBytes(tempFilenamePath, excelData);
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                using (var stream = File.Open(tempFilenamePath, FileMode.Open, FileAccess.Read))
                {
                    // Auto-detect format, supports:
                    //  - Binary Excel files (2.0-2003 format; *.xls)
                    //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsb)
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet();
                        DataTable dataTable = result.Tables[0];
                        reader.Close();

                        return dataTable;
                    }
                }

            }
            catch (Exception ex)
            {
                throw new Exception("Error executing reading", ex);
            }
        }
    }
}
