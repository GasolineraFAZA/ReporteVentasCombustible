using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EnviarVentasDiariasGasolinerasYPoint.Modelos
{
    public class CuerpoMailEstacion
    {
        public string FechaEstacion { get; set; }
        public string Estacion { get; set; }
        public double UnidadesLitros { get; set; }
        public double LitrosContado { get; set; }
        public double LitrosCredito { get; set; }
        public double ImporteLitros { get; set; }
        public double UnidadesLitrosAcum { get; set; }
        public double LitrosContadoAcum { get; set; }
        public double LitrosCreditoAcum { get; set; }
        public double ImporteLitrosAcum { get; set; }
        public double UnidadesDinero { get; set; }
        public double ContadoDinero { get; set; }
        public double CreditoDinero { get; set; }
        public double ImporteDinero { get; set; }
        public double UnidadesDineroAcum { get; set; }
        public double ContadoDineroAcum { get; set; }      
        public double CreditoDineroAcum { get; set; }
        public double ImporteDineroAcum { get; set; }
    }
}
