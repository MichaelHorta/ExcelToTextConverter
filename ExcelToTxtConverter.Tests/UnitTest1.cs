using System;
using System.IO;
using System.Reflection;
using System.Xml.Linq;
using Xunit;

namespace ExcelToTxtConverter.Tests
{
    public class UnitTest1
    {
        [Fact]
        public void Test1()
        {
            var definitionResource = RetrieveEmbeddedResourceAsStream("definition2.xml");
            var definition = XElement.Load(definitionResource);
            var converter = new Converter(definition);
            var excelData = RetrieveEmbeddedResourceAsByteArray("DIARIO_2019_01.xls");
            var result = converter.Execute(excelData);
            Assert.NotNull(result);
        }

        [Fact]
        public void Test2()
        {
            var definitionResource = RetrieveEmbeddedResourceAsStream("LM_Definition.xml");
            var definition = XElement.Load(definitionResource);
            var converter = new Converter(definition);
            var excelData = RetrieveEmbeddedResourceAsByteArray("LM_76166365-8_201701_0_1581301f-32f9-4119-8723-e6193bcd86cc_prod.xlsx");
            var result = converter.Execute(excelData);
            Assert.NotNull(result);
        }

        [Fact]
        public void Test3()
        {
            var definitionResource = RetrieveEmbeddedResourceAsStream("LD_Definition.xml");
            var definition = XElement.Load(definitionResource);
            var converter = new Converter(definition);
            var excelData = RetrieveEmbeddedResourceAsByteArray("LD_76245069-0_201901_21231231231231231234_prod.xlsx");
            var result = converter.Execute(excelData);
            Assert.NotNull(result);
        }

        private Stream RetrieveEmbeddedResourceAsStream(string resourceName)
        {
            var assembly = typeof(UnitTest1).GetTypeInfo().Assembly;
            return assembly.GetManifestResourceStream(string.Format("{0}.{1}", "ExcelToTxtConverter.Tests", resourceName));
        }

        private byte[] RetrieveEmbeddedResourceAsByteArray(string resourceName)
        {
            var stream = RetrieveEmbeddedResourceAsStream(resourceName);
            byte[] bufferRead = new byte[2048];
            int read = 0;
            using (MemoryStream ms = new MemoryStream())
            {
                while ((read = stream.Read(bufferRead, 0, bufferRead.Length)) > 0)
                {
                    ms.Write(bufferRead, 0, read);
                }

                return ms.ToArray();
            }
        }
    }
}
