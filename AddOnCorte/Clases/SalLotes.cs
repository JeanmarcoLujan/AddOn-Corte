using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnCorte.Clases
{
    class SalLotes
    {
        public BOM BOM { get; set; }
    }

    public class BOM
    {
        public BO BO { get; set; }
    }

    public class BO
    {
        public BTNT BTNT { get; set; }
        public IGE19 IGE19 { get; set; }
        public WTR19 WTR19 { get; set; }
        public DLN19 DLN19 { get; set; }
        public RPD19 RPD19 { get; set; }
        public RPC19 RPC19 { get; set; }
        public RDN19 RDN19 { get; set; }
        public RIN19 RIN19 { get; set; }
        public INV19 INV19 { get; set; }
        public PDN19 PDN19 { get; set; }
        public PCH19 PCH19 { get; set; }
        public IGN19 IGN19 { get; set; }
    }


    public class BTNT
    {
        public List<Row> row { get; set; }

    }

    public class IGE19
    {
        public List<Row> row { get; set; }
    }

    public class IGN19
    {
        public List<Row> row { get; set; }
    }

    public class DLN19
    {
        public List<Row> row { get; set; }
    }

    public class RDN19
    {
        public List<Row> row { get; set; }
    }

    public class PDN19
    {
        public List<Row> row { get; set; }
    }

    public class PCH19
    {
        public List<Row> row { get; set; }
    }

    public class RIN19
    {
        public List<Row> row { get; set; }
    }

    public class INV19
    {
        public List<Row> row { get; set; }
    }
    public class RPC19
    {
        public List<Row> row { get; set; }
    }
    public class RPD19
    {
        public List<Row> row { get; set; }
    }

    public class DOC
    {
        public Row row { get; set; }
    }


    public class WTR19
    {
        public List<Row> row { get; set; }
    }


    public class Row
    {
        public string ItemCode { get; set; }
        public int SysNumber { get; set; }
        public string DistNumber { get; set; }

        public string Quantity { get; set; }
        public string WhsCode { get; set; }
        public int DocType { get; set; }
        public int DocEntry { get; set; }
        public int DocLineNum { get; set; }
        public int BinAllocSe { get; set; }
        public int LineNum { get; set; }

        public int SnBType { get; set; }
        public int SnBMDAbs { get; set; }
        public int BinAbs { get; set; }
        public int ObjType { get; set; }
    }
}
