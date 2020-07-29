using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelToTxtConverter
{
    public class TextOrderableWriterFactory
    {
        public static TextOrderableBase Build(bool orderable)
        {
            if (orderable)
            {
                return new TextOrderableWriter();
            }
            else
            {
                return new TextNonOrderableWriter();
            }
        }
    }
}
