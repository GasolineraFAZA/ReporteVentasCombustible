using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EnviarVentasDiariasGasolinerasYPoint.Clases
{
    public class EstacionesYVariablesHTML
    {
        public List<string> GenerarVariableHTML(string Abreviatura, Constantes.EstacOSucur estacOSucur)
        {
            return
                Enumerable.Range(
                                    1, 
                                    estacOSucur == Constantes.EstacOSucur.Estacion?
                                            Constantes.CantEstaciones :
                                            Constantes.CantSucursales
                                ).Select(r => $"{'{'.ToString()}{Abreviatura}{r}{'}'.ToString()}").ToList();
        }
    }
}
