using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EnviarVentasDiariasGasolinerasYPoint.Modelos
{
    public class CuerpoMailSucursal
    {
        public string Sucursal { get; set; }
        public long Turno1Caja1 { get; set; }
        public long Turno1Caja2 { get; set; }
        public long Turno2Caja1 { get; set; }
        public long Turno2Caja2 { get; set; }
        public long Turno3Caja1 { get; set; }
        public long Turno3Caja2 { get; set; }
        public long TotalTrans { get; set; }
        public decimal TicketPromedio { get; set; }
        public decimal ImporteSucursal { get; set; }
    }
}
