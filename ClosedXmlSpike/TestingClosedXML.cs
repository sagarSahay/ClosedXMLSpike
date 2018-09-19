namespace ClosedXmlSpike
{
    using System;
    using System.Reflection;
    using ClosedXML.Excel;

    public class TestingClosedXML
    {
        public void LoadFile()
        {
            var path = AppContext.BaseDirectory.Substring(0, AppContext.BaseDirectory.IndexOf("bin"));

            var wb = new XLWorkbook(path + @"\Documents\spike.xlsx");

            var ws = wb.Worksheets.Worksheet("s1");

            foreach (var wsTable in ws.Tables)
            {
               
                Console.WriteLine("/////////////////////////////////////////////////////////////////////////////////////////////////////////");
                foreach (var xlTableRow in wsTable.DataRange.Rows())
                {
                    Console.WriteLine($"{xlTableRow.Field("KG").GetString()}\t" +
                                      $"{xlTableRow.Field("Zone 1").GetString()}->{xlTableRow.Field("Zone 1").Style.NumberFormat}\t" +
                                      $"{xlTableRow.Field("Zone 2").GetString()}\t" +
                                      $"{xlTableRow.Field("Zone 3").GetString()}\t" +
                                      $"{xlTableRow.Field("Zone 4").GetString()}\t" +
                                      $"{xlTableRow.Field("Zone 5").GetString()}\t" +
                                      $"{xlTableRow.Field("Zone 6").GetString()}\t" +
                                      $"{xlTableRow.Field("Zone 7").GetString()}\t" +
                                      $"{xlTableRow.Field("Zone 8").GetString()}\t" +
                                      $"{xlTableRow.Field("Zone 9").GetString()}\t" +
                                      $"{xlTableRow.Field("Zone 10").GetString()}");
                }
                Console.WriteLine("/////////////////////////////////////////////////////////////////////////////////////////////////////////");
            }


        }
    }
}
