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

    public class Agenda
    {
        public string DocEntry { get; set; }
        public string Fecha { get; set; }
        public string Equipo { get; set; }
        public string Serie { get; set; }
    }


    public class Agendado
    {
        public string DocEntry { get; set; }
        public string Almacen { get; set; }
    }

    public class SolicitudDetalle
    {
        public string MGS_CL_LOTE { get; set; }
        public string MGS_CL_ANCM { get; set; }
        public string MGS_CL_MLAR { get; set; }
        public string MGS_CL_MNBO { get; set; }
        public string MGS_CL_MCAN { get; set; }
        public string MGS_CL_FECADM { get; set; }
        public string MGS_CL_FIFO { get; set; }
    }

    public class CorridasDetalle
    {
        public string MGS_CL_TITU { get; set; }
        public string MGS_CL_C1 { get; set; }
        public string MGS_CL_C2 { get; set; }
        public string MGS_CL_C3 { get; set; }
        public string MGS_CL_C4 { get; set; }
        public string MGS_CL_C5 { get; set; }
        public string MGS_CL_C6 { get; set; }
        public string MGS_CL_C7 { get; set; }
        public string MGS_CL_C8 { get; set; }
        public string MGS_CL_C9 { get; set; }
        public string MGS_CL_C10 { get; set; }
    }
    public class Solicitud
    {
        public string DocEntry { get; set; }
        public DateTime MGS_CL_FECHA { get; set; }
        public string MGS_CL_CLIE { get; set; }
        public string MGS_CL_CLID { get; set; }
        public string MGS_CL_ARTD { get; set; }
        public string MGS_CL_PRAR { get; set; }
        public string MGS_CL_OFVE { get; set; }
        public string MGS_CL_OFVI { get; set; }
        public string MGS_CL_EFCO { get; set; }
        public string MGS_CL_ARTC { get; set; }
        public string MGS_CL_MTANC { get; set; }
        public string MGS_CL_MTLAR { get; set; }
        public string MGS_CL_MTCANT { get; set; }
        public string MGS_CL_OFEV { get; set; }
        public DateTime MGS_CL_AFECOR { get; set; }
        public List<SolicitudDetalle> Detalle { get; set; }
        public List<CorridasDetalle> DetalleCorridas { get; set; }
    }
}
