using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EnviarVentasDiariasGasolinerasYPoint.Modelos
{
    public class CuerpoMailSucursalPedidos
    {
        public string SUCURSAL { get; set; }
        public long NUM_SERVICIOS { get; set; }
        public decimal IMPORTE { get; set; }
    }
}
