using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnCorte.Clases
{
    class Globales
    {
        public static SAPbobsCOM.Company oCompany { get; set; }
        public static SAPbouiCOM.Application oApp { get; set; }
    }

    public class ColumnValueCount
    {
        public string ColumnName { get; set; }
        public string CellValue { get; set; }
        public double LargoValue { get; set; }
        public int Count { get; set; }
        public int ColumnId { get; set; }
    }
}
