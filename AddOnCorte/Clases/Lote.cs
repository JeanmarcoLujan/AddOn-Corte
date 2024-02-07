using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnCorte.Clases
{
    public class Detalle
    {
        public string MGS_CL_ARTICL { get; set; }
        public double MGS_CL_ANCHO { get; set; }
        public double MGS_CL_LARGO { get; set; }
        public double MGS_CL_CANTID { get; set; }
        public double MGS_CL_NUMBOB { get; set; }
        public string MGS_CL_LINNUM { get; set; }
        public string MGS_CL_LINNRE { get; set; }
        //public List<Ubicacion> Ubicaciones { get; set; }


    }
    class Lote
    {
        public string MGS_CL_ARTICL { get; set; }
        public double MGS_CL_ANCHO { get; set; }
        public double MGS_CL_LARGO { get; set; }
        public string MGS_CL_IDLOTE { get; set; }
        public double MGS_CL_CANTID { get; set; }
        public double MGS_CL_NUMBOB { get; set; }
        public string MGS_CL_LINNUM { get; set; }
        public string MGS_CL_LINNUMDET { get; set; }
        public string MGS_CL_CODWHS { get; set; }
    }

    public class Ubicacion
    {
        public string MGS_CL_UBICAC { get; set; }
        public string MGS_CL_ARTICL { get; set; }
        public double MGS_CL_ANCHO { get; set; }
        public double MGS_CL_LARGO { get; set; }
        public string MGS_CL_IDLOTE { get; set; }
        public double MGS_CL_CANTID { get; set; }
        public double MGS_CL_NUMBOB { get; set; }
        public string MGS_CL_LINNUM { get; set; }
        public string MGS_CL_LINNUMDET { get; set; }
    }
}
