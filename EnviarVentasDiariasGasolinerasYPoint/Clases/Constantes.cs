using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EnviarVentasDiariasGasolinerasYPoint.Clases
{
    public static class Constantes
    {
        public const int DiasAnteriores = -1;
        public const int CantSucursales = 7;
        public const int CantEstaciones = 20;

        public const bool PagoEnEfectivo = true;
        public const bool PagoEnCredito = true;
        public const bool PagoEnCredotYEfectivo = true;

        public enum EstacOSucur
        {
            Estacion,
            Sucursal
        }

        public static class EstacionesABuscar
        {
            public static string Mieleras = "mieleras";
            public static string SeisEnero = "seis";
            public static string SantaFe = "santa fe";
            public static string Bravo = "bravo";
            public static string Cuenca1 = "cuencame 1";

            public static string Cuenca2 = "cuencame 2";
            public static string SanJoaquin = "san joaquin";
            public static string Azulejos = "azulejos";
            public static string Urquizo = "urquizo";
            public static string SantaRita = "santa rita";

            public static string ParqueInd = "parque";
            public static string Filadelfia = "filadelfia";
            public static string Crit = "crit";
            public static string Triangulo = "triangulo";
            public static string Independencia = "indepen";
            public static string Abastos = "Abastos";

            public static string PuenteCent = "puente";

            public static readonly string[] Estaciones =
            {
                Mieleras,
                SeisEnero,
                SantaFe,
                Bravo,
                Cuenca1,

                Cuenca2,
                SanJoaquin,
                Azulejos,
                Urquizo,
                SantaRita,

                ParqueInd,
                Filadelfia,
                Crit,
                Triangulo,
                Independencia,

                PuenteCent,
                Abastos
            };
        }

        public static class EstacionesAbreviadas
        {
            public static string Mieleras = "MIE";
            public static string SeisEnero = "SEN";
            public static string SantaFe = "STF";
            public static string Bravo = "BRV";
            public static string Cuenca1 = "CU1";

            public static string Cuenca2 = "CU2";
            public static string SanJoaquin = "SJQ";
            public static string Azulejos = "AZU";
            public static string Urquizo = "URQ";
            public static string SantaRita = "STR";

            public static string ParqueInd = "PIN";
            public static string Filadelfia = "FIL";
            public static string Crit = "CRT";
            public static string Triangulo = "TRI";
            public static string Independencia = "IND";

            public static string PuenteCent = "PCE";
            public static string Abastos = "ABS";

            public static readonly string[] Estaciones =
            {
                Mieleras,
                SeisEnero,
                SantaFe,
                Bravo,
                Cuenca1,

                Cuenca2,
                SanJoaquin,
                Azulejos,
                Urquizo,
                SantaRita,

                ParqueInd,
                Filadelfia,
                Crit,
                Triangulo,
                Independencia,

                PuenteCent,
                Abastos
            };
        }

        public static class ColoresGraficas
        {
            public static Color DimGray = Color.DimGray;
            public static Color Fuchshia = Color.Fuchsia;
            public static Color SkyBlue = Color.SkyBlue;
            public static Color Blue = Color.Blue;
            public static Color Orange = Color.Orange;

            public static Color Red = Color.Red;
            public static Color Aqua = Color.Aqua;
            public static Color Indigo = Color.Indigo;
            public static Color Navy = Color.Navy;
            public static Color DarkCyan = Color.DarkCyan;

            public static Color Purple = Color.Purple;
            public static Color Brown = Color.Brown;
            public static Color Green = Color.Green;
            public static Color Gray = Color.Gray;
            public static Color LimeGreen = Color.LimeGreen;

            public static Color Maroon = Color.Maroon;
            public static Color Yellow = Color.Yellow;
            public static Color Turquoise = Color.Turquoise;

            public static readonly Color[] Colores =
            {
                DimGray,
                Fuchshia,
                SkyBlue,
                Blue,
                Orange,

                Red,
                Aqua,
                Indigo,
                Navy,
                DarkCyan,

                Purple,
                Brown,
                Green,
                Gray,
                LimeGreen,

                Maroon,
                Yellow,
                Turquoise
            };
        }

        public static class PathImagenes
        {
            public const int GraficaAnchoDeLinea = 5;

            public enum  PathsImgs
            {
                PathFAZA,           // 1
                PathCarrito,        // 2
                VtaProdEfec,        // 3
                VtaProdCred,        // 4
                VtaEstEfec,         // 5
                VtaEstEfec2,        // 6
                VtaEstCred,         // 7
                VtaEstCred2,        // 8
                VtaProdEfecYCred,   // 9
                VtaEstEfecYCred,    // 10
                VtaEstEfecYCred2,   // 11
            };

            private static string PathExe =
                System.AppDomain.CurrentDomain.BaseDirectory;

            private static string PathFAZA =                                //1
                PathExe + "Imagenes\\Grupo_Faza.jpg";

            private static string PathCarrito =                             //2
                PathExe + "Imagenes\\carrito.png";

            private static string GraficaVtaPorProductoEfectivo =           //3
                PathExe + "Imagenes\\GraficaVtaEfectivoPorProducto.jpg";

            private static string GraficaVtaPorProductoCredito =            //4
                PathExe + "Imagenes\\GraficaVtaCreditoPorProducto.jpg";

            private static string GraficaVtaPorEstacionEfectivo =           //5
                PathExe + "Imagenes\\GraficaVtaEfectivoPorEstacion.jpg";

            private static string GraficaVtaPorEstacionEfectivo2 =          //6
                PathExe + "Imagenes\\GraficaVtaEfectivoPorEstacion2.jpg";

            private static string GraficaVtaPorEstacionCredito =            //7
                PathExe + "Imagenes\\GraficaVtaCreditoPorEstacion.jpg";

            private static string GraficaVtaPorEstacionCredito2 =           //8
                PathExe + "Imagenes\\GraficaVtaCreditoPorEstacion2.jpg";

            private static string GraficaVtaPorProductoEfectivoYCred =      //9
                PathExe + "Imagenes\\GraficaVtaEfectivoYCredPorProducto.jpg";

            private static string GraficaVtaPorEstacionEfectivoYCred =      //10
                PathExe + "Imagenes\\GraficaVtaEfectivoYCredPorEstacion.jpg";

            private static string GraficaVtaPorEstacionEfectivoYCred2 =     //11
                PathExe + "Imagenes\\GraficaVtaEfectivoYCredPorEstacion2.jpg";

            public static readonly string[] ImagenesAIncluir =
            {
                PathFAZA,                               // 1
                PathCarrito,                            // 2
                GraficaVtaPorProductoEfectivo,          // 3
                GraficaVtaPorProductoCredito,           // 4
                GraficaVtaPorEstacionEfectivo,          // 5
                GraficaVtaPorEstacionEfectivo2,         // 6
                GraficaVtaPorEstacionCredito,           // 7
                GraficaVtaPorEstacionCredito2,          // 8
                GraficaVtaPorProductoEfectivoYCred,     // 9
                GraficaVtaPorEstacionEfectivoYCred,     // 10
                GraficaVtaPorEstacionEfectivoYCred2     // 11
            };
        }

        public static class CorreosTodaLaInformacion
        {
            public static string PathPlantillaCorreo =
                System.AppDomain.CurrentDomain.BaseDirectory + "\\Plantilla\\PlantillaTablaCorreo.html";

            public static string PathPlantillaRenglPedCNT =
                System.AppDomain.CurrentDomain.BaseDirectory + "\\Plantilla\\PlantillaRenglonPedidoCNT.html";

            public static string PathPlantillaRenglPedMTR =
                System.AppDomain.CurrentDomain.BaseDirectory + "\\Plantilla\\PlantillaRenglonPedidoMTR.html";

            public static string PathPlantillaRenglPedSTF =
                System.AppDomain.CurrentDomain.BaseDirectory + "\\Plantilla\\PlantillaRenglonPedidoSTF.html";

            public static readonly string[] CorreosConTodaLaInformacionTo = 
            {
                "55sdnzobWVoYyr3QJHZm4NfzERt+/llTYRJNwnNVv28=",       //Se debe mandar a los 5 primeros
                "CO1R7KrKv9+4YGcRhL5niOllu8/d7HS4Gijn0TER7as=",
                "+F9xYm1knWTsZ5X/Gm21uGpLdv9oOBiuPWpottrklEA=",
                "RT3T9D1QE9ZwCjd1gz2jezbdYjnP0kj7U3kVBjdtQLsOHkCU1YQSwQH0UbajapHZ",
                "WTR9lK7Hc3yUYtEgEum3WBCBlKVg8ZpC6OEx0gg+JK4=",         // Erik

                "uOzNxS19nr5pZ3/xle9LQo2AybS9C5xQAe8mY7gsPJc=",         // José
                "W7hPHv9jE7rDOI8Vnjh5z94J+BUcYcXYZjLhngy+WbQ=",         // Miguel
                "JbOnWg/pqkyrUPVp5uRahVc3UH0Dq+hy28T5/nUD208="          //reyes_campos@hotmail.com

            };

            public static readonly string[] CorreosConTodaLaInformacionBcc =
            {
                "uOzNxS19nr5pZ3/xle9LQo2AybS9C5xQAe8mY7gsPJc=",           //jose.campos@faza.com.mx
                //"W7hPHv9jE7rDOI8Vnjh5z94J+BUcYcXYZjLhngy+WbQ=",           //Miguel.flores@faza.com.mx
                "ZPOfMabfEM64FPG7mXfJ7BxO+5sRc5cjnhEHS13eVE8=",           //Iván
            };
        }

        public static class CorreosSoloEstaciones
        {
            public static string PathPlantillaCorreoGasolinera =
                System.AppDomain.CurrentDomain.BaseDirectory + "\\Plantilla\\PlantillaTablaCorreoGasolinera.html";

            public static readonly string[] CorreosConSoloEstacionesTo =
            {
                "Ru9w8BtCuA/xlOkOHcM5OgK7K9SiqGdaBW21r2NlMQI=",         //Se debe mandar a los 4 primeros
                "qXkdFmC1Br9Qy8s9JgZpmVCvkXveXZpDqa5p628AfUo=",
                "dlFGJp2g+gZr8kUhUrVvJo05JH02VDC+Nys48ucpULs=",
                "MSI8E0iLiWIFJ4Jz+AHtQjnFj3J23x73L7Gjl6o98v0=",

                "uOzNxS19nr5pZ3/xle9LQo2AybS9C5xQAe8mY7gsPJc=",          //jose.campos@faza.com.mx
                "W7hPHv9jE7rDOI8Vnjh5z94J+BUcYcXYZjLhngy+WbQ=",          //Miguel.flores@faza.com.mxm
                "ZPOfMabfEM64FPG7mXfJ7BxO+5sRc5cjnhEHS13eVE8=",          //Iván
            };

            public static readonly string[] CorreosConSoloEstacionesBcc =
            {
                "uOzNxS19nr5pZ3/xle9LQo2AybS9C5xQAe8mY7gsPJc=",          //jose.campos@faza.com.mx
                //"W7hPHv9jE7rDOI8Vnjh5z94J+BUcYcXYZjLhngy+WbQ=",           //Miguel.flores@faza.com.mx
                "ZPOfMabfEM64FPG7mXfJ7BxO+5sRc5cjnhEHS13eVE8=",           //Iván
            };
        }

        public static class CorreosSoloSucursales
        {
            public static string PathPlantillaCorreo =
                System.AppDomain.CurrentDomain.BaseDirectory + "\\Plantilla\\PlantillaTablaSoloSucursales.html";

            public static readonly string[] CorreosConSoloSucursalesTo =
            {
                //Se debe mandar solo a los 3 primeros
                "MSI8E0iLiWIFJ4Jz+AHtQjnFj3J23x73L7Gjl6o98v0=",                             //Correo jefe comercial1
                "6Rs6vIYtW349PIX5/46gxCqB4mQpHaM12Bqi1Gn6OYQ=",                             //Correo jefe comercial2
                "mie96bBhyMMQjQfD+pnfpXgGg5wrzOXCU917NurZbLM=",                           //Correo garza

                "uOzNxS19nr5pZ3/xle9LQo2AybS9C5xQAe8mY7gsPJc=",                             //jose.campos@faza.com.mx
                "W7hPHv9jE7rDOI8Vnjh5z94J+BUcYcXYZjLhngy+WbQ=",                             //Miguel.flores@faza.com.mx
                "ZPOfMabfEM64FPG7mXfJ7BxO+5sRc5cjnhEHS13eVE8=",                             //Iván
                "JbOnWg/pqkyrUPVp5uRahVc3UH0Dq+hy28T5/nUD208="                            //reyes_campos@hotmail.com
            };

            public static readonly string[] CorreosConSoloSucursalesBcc =
            {
                "uOzNxS19nr5pZ3/xle9LQo2AybS9C5xQAe8mY7gsPJc=",    //jose.campos@faza.com.mx
                //"W7hPHv9jE7rDOI8Vnjh5z94J+BUcYcXYZjLhngy+WbQ=",  //Miguel.flores@faza.com.mx
                "ZPOfMabfEM64FPG7mXfJ7BxO+5sRc5cjnhEHS13eVE8=",    //Iván
            };
        }

        public static class TamanioGrafica
        {
            private const int Ancho = 694;
            private const int Alto = 550;

            public enum PropTam { Ancho, Alto };

            public static readonly int[] Tamanio =
            {
                Ancho, Alto
            };
        }

        public static readonly Dictionary<string, string> CorregirEstaciones = new Dictionary<string, string>()
        {
            { "PuenteCent", "Puente Centenario" },
        };
    }
}
