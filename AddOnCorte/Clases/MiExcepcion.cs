using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnCorte.Clases
{
    class MiExcepcion: Exception
    {
        public MiExcepcion(string mensaje) : base(mensaje)
        {
        }

    }
}
