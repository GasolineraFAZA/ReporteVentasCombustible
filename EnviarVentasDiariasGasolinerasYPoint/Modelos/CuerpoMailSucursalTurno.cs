using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EnviarVentasDiariasGasolinerasYPoint.Modelos
{
    public class CuerpoMailSucursalTurno
    {
        public string Sucursal { get; set; }
        public TimeSpan Turno1Caja1 { get; set; }
        public TimeSpan Turno1Caja2 { get; set; }
        public TimeSpan Turno2Caja1 { get; set; }
        public TimeSpan Turno2Caja2 { get; set; }
        public TimeSpan Turno3Caja1 { get; set; }
        public TimeSpan Turno3Caja2 { get; set; }
    }
}
