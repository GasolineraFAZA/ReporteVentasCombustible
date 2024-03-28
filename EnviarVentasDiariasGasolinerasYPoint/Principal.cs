using Encriptacion;
using EnviarVentasDiariasGasolinerasYPoint.Clases;
using EnviarVentasDiariasGasolinerasYPoint.Modelos;
using FormsCostoVentas;
using FormsDespachos;
using FormsDespachos.Clases;
using FormsMicrosipVentasDia;
using FormsMicrosipVentasDia.Clases;
using FormsPrueba;
using FormsPrueba.Clases;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace EnviarVentasDiariasGasolinerasYPoint
{
    public partial class Principal : Form
    {
        bool ResultadoCargaMicrosip = false;
        bool ResultadoCargaAcumMicrosip = false;
        bool ResultadoCargaPedidosMicrosip = false;
        bool ResultadoCargaAcumPedidosMicrosip = false;

        string Fecha = string.Empty, FechaDia, FechaMensAcum;
        string FechaIniMes = string.Empty;

        string TextPlantilla = string.Empty;
        string TextPlantillaRengPedCNT = string.Empty;
        string TextPlantillaRengPedMTR = string.Empty;
        string TextPlantillaRengPedSTF = string.Empty;

        long TT1C1 = 0, TT1C2 = 0, TT2C1 = 0, TT2C2 = 0, TT3C1 = 0, TT3C2 = 0, TT = 0;
        long TT1C1A = 0, TT1C2A = 0, TT2C1A = 0, TT2C2A = 0, TT3C1A = 0, TT3C2A = 0, TTA = 0;
        long C1TKT = 0, C2TKT = 0, TOTTKT = 0;
        decimal IMPTKT = 0, TKPROMT = 0, TTP = 0, TTPA = 0;

        double TotalLLC1 = 0, TotalLLE1 = 0, TotalIE1 = 0, TotalIC1 = 0, TotalI1 = 0;
        double TotalLAC1 = 0, TotalLAE1 = 0, TotalIEA1 = 0, TotalICA1 = 0, TotalIA1 = 0;

        double TotalL1PI = 0, TotalU1PI = 0, TotalI1PI = 0;
        double TotalLA1PI = 0, TotalUA1PI = 0, TotalIA1PI = 0;

        double TOTALUL1 = 0, TOTALCL1 = 0, TOTALEL1 = 0, TOTALIL1 = 0;
        double TOTALULA1 = 0, TOTALCLA1 = 0, TOTALELA1 = 0, TOTALILA1 = 0;

        double TOTALUD1 = 0, TOTALCD1 = 0, TOTALED1 = 0, TOTALID1 = 0;
        double TOTALUDA1 = 0, TOTALCDA1 = 0, TOTALEDA1 = 0, TOTALIDA1 = 0;

        decimal TIS = 0, TISA = 0, TISP = 0, TISPA = 0;

        List<CuerpoMailEstacion> LstCuerpoMailEstacion = new List<CuerpoMailEstacion>();
        List<CuerpoMailSucursal> LstCuerpoMailSucursal = new List<CuerpoMailSucursal>();
        List<CuerpoMailSucursalPedidos> LstCuerpoMailSucursalPed = new List<CuerpoMailSucursalPedidos>();

        List<CuerpoMailEstacion> LstCuerpoMailEstacionAcum = new List<CuerpoMailEstacion>();
        List<CuerpoMailSucursal> LstCuerpoMailSucursalAcum = new List<CuerpoMailSucursal>();
        List<CuerpoMailSucursalPedidos> LstCuerpoMailSucursalPedAcum = new List<CuerpoMailSucursalPedidos>();

        //List<CuerpoMailTickets> 
        List<CuerpoMailSucursalTurno> LstCuerpoMailSucursalTurno = new List<CuerpoMailSucursalTurno>();

        enum TransaccionesPorTurnoYCajaReng1 { T11R1, T12R1, T21R1, T22R1, T31R1, T32R1 }
        enum TransaccionesPorTurnoYCajaReng2 { T11R2, T12R2, T21R2, T22R2, T31R2, T32R2 }
        enum TransaccionesPorTurnoYCajaReng3 { T11R3, T12R3, T21R3, T22R3, T31R3, T32R3 }
        enum TransaccionesPorTurnoYCajaReng4 { T11R4, T12R4, T21R4, T22R4, T31R4, T32R4 }
        enum TransaccionesPorTurnoYCajaReng5 { T11R5, T12R5, T21R5, T22R5, T31R5, T32R5 }
        enum TransaccionesPorTurnoYCajaReng6 { T11R6, T12R6, T21R6, T22R6, T31R6, T32R6 }
        enum TransaccionesPorTurnoYCajaReng7 { T11R7, T12R7, T21R7, T22R7, T31R7, T32R7 }

        enum ImportesSucursales { IS1, IS2, IS3, IS4, IS5, IS6, IS7 };

        enum TransaccionesTotales { T1, T2, T3, T4, T5, T6, T7 };
        enum TransaccionesTurnoYCajaTotales { TT1C1, TT1C2, TT2C1, TT2C2, TT3C1, TT3C2 };
        enum TransaccionesPromedio { TP1, TP2, TP3, TP4, TP5, TP6, TP7 };

        enum TransaccionesPorTurnoYCajaRengAcum1 { T11A1, T12A1, T21A1, T22A1, T31A1, T32A1 }
        enum TransaccionesPorTurnoYCajaRengAcum2 { T11A2, T12A2, T21A2, T22A2, T31A2, T32A2 }
        enum TransaccionesPorTurnoYCajaRengAcum3 { T11A3, T12A3, T21A3, T22A3, T31A3, T32A3 }
        enum TransaccionesPorTurnoYCajaRengAcum4 { T11A4, T12A4, T21A4, T22A4, T31A4, T32A4 }
        enum TransaccionesPorTurnoYCajaRengAcum5 { T11A5, T12A5, T21A5, T22A5, T31A5, T32A5 }
        enum TransaccionesPorTurnoYCajaRengAcum6 { T11A6, T12A6, T21A6, T22A6, T31A6, T32A6 }
        enum TransaccionesPorTurnoYCajaRengAcum7 { T11A7, T12A7, T21A7, T22A7, T31A7, T32A7 }

        enum TurnosCajaHoraR1 { T1C1R1H, T1C2R1H, T2C1R1H, T2C2R1H, T3C1R1H, T3C2R1H }
        enum TurnosCajaHoraR2 { T1C1R2H, T1C2R2H, T2C1R2H, T2C2R2H, T3C1R2H, T3C2R2H }
        enum TurnosCajaHoraR3 { T1C1R3H, T1C2R3H, T2C1R3H, T2C2R3H, T3C1R3H, T3C2R3H }
        enum TurnosCajaHoraR4 { T1C1R4H, T1C2R4H, T2C1R4H, T2C2R4H, T3C1R4H, T3C2R4H }
        enum TurnosCajaHoraR5 { T1C1R5H, T1C2R5H, T2C1R5H, T2C2R5H, T3C1R5H, T3C2R5H }
        enum TurnosCajaHoraR6 { T1C1R6H, T1C2R6H, T2C1R6H, T2C2R6H, T3C1R6H, T3C2R6H }
        enum TurnosCajaHoraR7 { T1C1R7H, T1C2R7H, T2C1R7H, T2C2R7H, T3C1R7H, T3C2R7H }

        enum NoTicketsReng1 { C1TK1, C2TK1, TOTTK1, IMPTK1, TKPROM1 }
        enum NoTicketsReng2 { C1TK2, C2TK2, TOTTK2, IMPTK2, TKPROM2 }
        enum NoTicketsReng3 { C1TK3, C2TK3, TOTTK3, IMPTK3, TKPROM3 }
        enum NoTicketsReng4 { C1TK4, C2TK4, TOTTK4, IMPTK4, TKPROM4 }
        enum NoTicketsReng5 { C1TK5, C2TK5, TOTTK5, IMPTK5, TKPROM5 }
        enum NoTicketsReng6 { C1TK6, C2TK6, TOTTK6, IMPTK6, TKPROM6 }
        enum NoTicketsReng7 { C1TK7, C2TK7, TOTTK7, IMPTK7, TKPROM7 }
        enum NoTicketsRengTotales { C1TKT, C2TKT, TOTTKT, IMPTKT, TKPROMT }

        enum ImportesSucursalesAcumuladas { ISA1, ISA2, ISA3, ISA4, ISA5, ISA6, ISA7 };
        enum TransaccionesTotalesAcumuladas { TA1, TA2, TA3, TA4, TA5, TA6, TA7 };
        enum TransaccionesPromedioAcumuladas { TPA1, TPA2, TPA3, TPA4, TPA5, TPA6, TPA7 };
        enum TransaccionesTurnoYCajaTotalesAcumuladas { TT11A, TT12A, TT21A, TT22A, TT31A, TT32A };

        enum Productos
        {
            Magna = 1,
            Premium = 2,
            Diesel = 27
        };

        enum TipoPago {
            Efectivo = 49,
            Debito = 52,
            Contado = 53,
            Credito1 = 50,
            Credito2 = 51,
            Transf_Banc = 56
        };

        DateTime DiaAntes = DateTime.Now.AddDays(Constantes.DiasAnteriores);
        //DateTime DiaAntes = DateTime.Now.AddDays(-(DateTime.Now.Day + DateTime.DaysInMonth(2021, 9) + DateTime.DaysInMonth(2021, 8) + DateTime.DaysInMonth(2021, 7) + DateTime.DaysInMonth(2021 ,6) + DateTime.DaysInMonth(2021, 5) + DateTime.DaysInMonth(2021, 4) + DateTime.DaysInMonth(2021, 3) + DateTime.DaysInMonth(2021, 2)));
        DateTime InicioMes = DateTime.MinValue;

        Func<TimeSpan, string> DarFormatoDeTiempo = r => r == TimeSpan.MinValue ? "-" : r.ToString().Split(':')[0] + ":" +
            r.ToString().Split(':')[1] + ":" + r.ToString().Split(':')[2].Substring(0, 2);

        Func<string, string> PrimMayDemMinus = r => r.Length < 2 ? r : r.First().ToString().ToUpper() +
            r.Substring(1, r.Length - 1).ToLower();

        Func<string, string> DivLetrDeNum = r =>
           r.Insert(r.Select((s, i) => char.IsDigit(s) ? i : 2000).Min(), " ");

        Func<string, string> ElimCeroIni = r =>
            r.Length < 2 ? r : r.First() == '0' ? r.Substring(1, r.Length - 1) : r;

        Func<string, bool> EsGasolina = r =>    r.ToUpper().Contains("MAGNA") ||
                                                r.ToUpper().Contains("DIESEL") ||
                                                r.ToUpper().Contains("PREMIUM") ||
                                                r.ToUpper().Contains("REGULAR") ||
                                                r.ToUpper().Contains("SUPREME") ||
                                                r.ToUpper().Contains("DIESEL-");

        public Principal()
        {
            InitializeComponent();
        }

        private void AsignarPermisosEscrituraArchsImagenes()
        {
            int IndVtaProdEfec = (int)Constantes.PathImagenes.PathsImgs.VtaProdEfec;
            AddFileSecurity(Constantes.PathImagenes.ImagenesAIncluir[IndVtaProdEfec], @"DESKTOP-P8HBRIE\desarrollo",
                FileSystemRights.CreateFiles, AccessControlType.Allow);

            int IndVtaProdCred = (int)Constantes.PathImagenes.PathsImgs.VtaProdCred;
            AddFileSecurity(Constantes.PathImagenes.ImagenesAIncluir[IndVtaProdCred], @"DESKTOP-P8HBRIE\desarrollo",
                FileSystemRights.CreateFiles, AccessControlType.Allow);

            int IndVtaEstEfec = (int)Constantes.PathImagenes.PathsImgs.VtaEstEfec;
            AddFileSecurity(Constantes.PathImagenes.ImagenesAIncluir[IndVtaEstEfec], @"DESKTOP-P8HBRIE1\desarrollo",
                FileSystemRights.CreateFiles, AccessControlType.Allow);

            int IndVtaEstCred = (int)Constantes.PathImagenes.PathsImgs.VtaEstEfec;
            AddFileSecurity(Constantes.PathImagenes.ImagenesAIncluir[IndVtaEstCred], @"DESKTOP-P8HBRIE\desarrollo",
                FileSystemRights.CreateFiles, AccessControlType.Allow);

            int IndVtaProdEfecYCred = (int)Constantes.PathImagenes.PathsImgs.VtaProdEfecYCred;
            AddFileSecurity(Constantes.PathImagenes.ImagenesAIncluir[IndVtaProdEfecYCred], @"DESKTOP-P8HBRIE\desarrollo",
                FileSystemRights.CreateFiles, AccessControlType.Allow);

            int IndVtaEstEfecYCred = (int)Constantes.PathImagenes.PathsImgs.VtaEstEfecYCred;
            AddFileSecurity(Constantes.PathImagenes.ImagenesAIncluir[IndVtaEstEfecYCred], @"DESKTOP-P8HBRIE\desarrollo",
                FileSystemRights.CreateFiles, AccessControlType.Allow);
        }

        // Removes an ACL entry on the specified file for the specified account.
        // Adds an ACL entry on the specified file for the specified account.
        public void AddFileSecurity(string fileName, string account,
            FileSystemRights rights, AccessControlType controlType)
        {

            // Get a FileSecurity object that represents the
            // current security settings.
            FileSecurity fSecurity = File.GetAccessControl(fileName);

            var sid = new SecurityIdentifier(WellKnownSidType.AuthenticatedUserSid, null);

            // Add the FileSystemAccessRule to the security settings.
            fSecurity.AddAccessRule(new FileSystemAccessRule(sid,
                rights, controlType));

            // Set the new access settings.
            File.SetAccessControl(fileName, fSecurity);
        }

        private void Principal_Load(object sender, EventArgs e)
        {
            IniciarFechas();

            CargarInfoControlGas(true, ref LstCuerpoMailEstacion);
            ResultadoCargaMicrosip = CargarInfoMicrosip(true, ref LstCuerpoMailSucursal);
            ResultadoCargaPedidosMicrosip = CargarIfoMicrosipPedidos(true, ref LstCuerpoMailSucursalPed);

            CargarInfoControlGas(false, ref LstCuerpoMailEstacionAcum);
            ResultadoCargaAcumMicrosip = CargarInfoMicrosip(false, ref LstCuerpoMailSucursalAcum);
            ResultadoCargaAcumPedidosMicrosip = CargarIfoMicrosipPedidos(false, ref LstCuerpoMailSucursalPedAcum);

            CargarInfoMicrosipHoraInicioTurno();
            //AsignarPermisosEscrituraArchsImagenes();

            CargarInfoCG(Constantes.PagoEnEfectivo, !Constantes.PagoEnCredito, !Constantes.PagoEnCredotYEfectivo);
            CargarInfoCG(!Constantes.PagoEnEfectivo, Constantes.PagoEnCredito, !Constantes.PagoEnCredotYEfectivo);
            CargarInfoCG(!Constantes.PagoEnEfectivo, !Constantes.PagoEnCredito, Constantes.PagoEnCredotYEfectivo);

            EnviarSMTP();
        }

        private bool CargarIfoMicrosipPedidos(bool DiarioOAcumulado,
            ref List<CuerpoMailSucursalPedidos> LstCuerpoMailSucursalPed)
        {
            MicrosipVentasDia microsipVentaDia = new MicrosipVentasDia();

            Func<bool, string> InicioMesODiaAnterior = r => r ?
                Fecha : InicioMes.ToString("dddd, dd MMMM yyyy");

            Func<string, DateTime> ConvAFecha = r => DateTime.Parse(r);

            List<FormsMicrosipVentasDia.Clases.DatosAlmacenes> LstSucursales =
                new List<FormsMicrosipVentasDia.Clases.DatosAlmacenes>();

            if (!CargarSucursales(ref LstSucursales))
                return false;

            LstCuerpoMailSucursalPed.Clear();

            foreach (FormsMicrosipVentasDia.Clases.DatosAlmacenes almacen in LstSucursales)
            {
                List<DatosVtasDiaV2> LstVtasDia = null;

                bool Error = false;

                DateTime FechaIni = ConvAFecha(InicioMesODiaAnterior(DiarioOAcumulado));
                DateTime FechaFin = ConvAFecha(Fecha);

                try
                {
                    switch (almacen.NOMBRE)
                    {
                        case "POINT MIELERAS":
                            LstVtasDia = microsipVentaDia.
                                ObtVtasDiaV2MIE(FechaIni, FechaFin);
                            break;

                        case "POINT SANTA FE":
                            LstVtasDia = microsipVentaDia.
                                ObtVtasDiaV2STF(FechaIni, FechaFin);
                            break;

                        case "POINT MONTE REAL":
                            LstVtasDia = microsipVentaDia.
                                ObtVtasDiaV2MTR(FechaIni, FechaFin);
                            break;

                        case "POINT CENTENARIO":
                            LstVtasDia = microsipVentaDia.
                                ObtVtasDiaV2CNT(FechaIni, FechaFin);
                            break;

                        case "POINT LA AMISTAD":
                            LstVtasDia = microsipVentaDia.
                                ObtVtasDiaV2AMD(FechaIni, FechaFin);
                            break;

                        case "POINT LAS  ETNIAS":
                            LstVtasDia = microsipVentaDia.
                                ObtVtasDiaV2ETN(FechaIni, FechaFin);
                            break;

                        case "POINT SAN JOSE DE VIÑEDO":
                            //LstVtasDia = microsipVentaDia.
                            //    ObtVtasDiaV2SJV(FechaIni, FechaFin);

                            LstVtasDia = new List<DatosVtasDiaV2>();
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Error = true;
                    throw new Exception(ex.InnerException.Message);
                }

                foreach (var elem1 in Error ? new List<DatosVtasDiaV2>() : LstVtasDia)
                {
                    LstCuerpoMailSucursalPed.Add(new CuerpoMailSucursalPedidos
                    {
                        SUCURSAL = almacen.NOMBRE,
                        NUM_SERVICIOS = elem1.NUM_SERVICIOS,
                        IMPORTE = elem1.IMPORTE
                    });
                }
            }

            return true;
        }

        private Tuple<DateTime, DateTime> EntregarMesIniYMesFin(int MesesADescontar, int DiasADescontar)
        {
            int MesesDescon = (-1) * MesesADescontar;
            int DiasDescon = (-1) * DiasADescontar;

            DateTime FechaMesIni =
                new DateTime(
                                DateTime.Now.AddMonths(MesesDescon).AddDays(DiasDescon).Year,
                                DateTime.Now.AddMonths(MesesDescon).AddDays(DiasDescon).Month,
                                1
                            );

            DateTime FechaMesFin =
                new DateTime(
                                DateTime.Now.AddMonths(MesesDescon).AddDays(DiasDescon).Year,
                                DateTime.Now.AddMonths(MesesDescon).AddDays(DiasDescon).Month,
                                DateTime.DaysInMonth(
                                                        DateTime.Now.AddMonths(MesesDescon).AddDays(DiasDescon).Year,
                                                        DateTime.Now.AddMonths(MesesDescon).AddDays(DiasDescon).Month
                                                    )
                            );

            return new Tuple<DateTime, DateTime>(FechaMesIni, FechaMesFin);
        }

        private void CuerpoProductoPorMes(Tuple<DateTime, DateTime>[] CalcFechaIniYFinArr,
            List<List<DatosDespachosFormPago>> LstDespachosEnEfectivoArr, ref Chart chartVtaProdLineaBarra,
            string PathImagen, SeriesChartType chartType)
        {
            chartVtaProdLineaBarra.Series.Clear();
            chartVtaProdLineaBarra.Size = new Size(
                Constantes.TamanioGrafica.Tamanio[(int)Constantes.TamanioGrafica.PropTam.Ancho],
                Constantes.TamanioGrafica.Tamanio[(int)Constantes.TamanioGrafica.PropTam.Alto]
            );

            Series seriesMagna = chartVtaProdLineaBarra.Series.Add("Magna");
            Series seriesPremium = chartVtaProdLineaBarra.Series.Add("Premium");
            Series seriesDiesel = chartVtaProdLineaBarra.Series.Add("Diesel");

            seriesMagna.BorderWidth = Constantes.PathImagenes.GraficaAnchoDeLinea;
            seriesPremium.BorderWidth = Constantes.PathImagenes.GraficaAnchoDeLinea;
            seriesDiesel.BorderWidth = Constantes.PathImagenes.GraficaAnchoDeLinea;

            seriesMagna.Color = Color.Green;
            seriesPremium.Color = Color.Red;
            seriesDiesel.Color = Color.Black;

            seriesMagna.ChartType = chartType;
            seriesPremium.ChartType = chartType;
            seriesDiesel.ChartType = chartType;

            chartVtaProdLineaBarra.ChartAreas["ChartArea1"].Position.Auto = false;
            chartVtaProdLineaBarra.ChartAreas["ChartArea1"].Position.X = 0;
            chartVtaProdLineaBarra.ChartAreas["ChartArea1"].Position.Y = 20;
            chartVtaProdLineaBarra.ChartAreas["ChartArea1"].Position.Width = 90;
            chartVtaProdLineaBarra.ChartAreas["ChartArea1"].Position.Height = 100;
            chartVtaProdLineaBarra.ChartAreas["ChartArea1"].AxisY.LabelStyle.Format = "###,##0.00 Lts";

            for (int ii = 0; ii < 12; ii++)
            {
                string Mes = new DateTime(DateTime.Now.Year, CalcFechaIniYFinArr[ii].Item1.Month, 1).ToString("MMMM",
                    CultureInfo.CreateSpecificCulture("es")) + " /n " + CalcFechaIniYFinArr[ii].Item1.Year;

                seriesMagna.Points.AddXY(Mes,
                    LstDespachosEnEfectivoArr[ii].Where(r => 
                                                            r.Producto.Trim().ToLower().Contains("magna")  ||
                                                            r.Producto.Trim().ToLower().Contains("regular")
                                                       ).
                    Sum(r => r.Litros));

                seriesPremium.Points.AddXY(Mes,
                    LstDespachosEnEfectivoArr[ii].Where(r => 
                                                            r.Producto.Trim().ToLower().Contains("premium") ||
                                                            r.Producto.Trim().ToLower().Contains("supreme")
                                                       ).
                    Sum(r => r.Litros));

                seriesDiesel.Points.AddXY(Mes,
                    LstDespachosEnEfectivoArr[ii].Where(r => 
                                                            r.Producto.Trim().ToLower().Contains("diesel") ||
                                                            r.Producto.Trim().ToLower().Contains("diesel-")
                                                       ).
                    Sum(r => r.Litros));

                var lst =
                    LstDespachosEnEfectivoArr[ii].Where(r =>
                                                            r.Producto.Trim().ToLower().Contains("regular")
                                                       ).ToList();

                if (lst.Count() > 0)
                {
                }
            }

            chartVtaProdLineaBarra.Legends[0].LegendStyle = LegendStyle.Column;
            chartVtaProdLineaBarra.Legends[0].Position = new ElementPosition(90, 40, 10, 40);

            chartVtaProdLineaBarra.SaveImage(PathImagen, ChartImageFormat.Jpeg);
        }

        private void CuerpoProductoPorMes(Tuple<DateTime, DateTime>[] CalcFechaIniYFinArr,
            List<List<DespachosContadoCreditoDebitoEfectivoMontoYLitros>> LstDespachosEnEfectivoArr, ref Chart chartVtaProdLineaBarra,
            string PathImagen, SeriesChartType chartType)
        {
            chartVtaProdLineaBarra.Series.Clear();
            chartVtaProdLineaBarra.Size = new Size(
                Constantes.TamanioGrafica.Tamanio[(int)Constantes.TamanioGrafica.PropTam.Ancho],
                Constantes.TamanioGrafica.Tamanio[(int)Constantes.TamanioGrafica.PropTam.Alto]
            );

            Series seriesMagna = chartVtaProdLineaBarra.Series.Add("Magna");
            Series seriesPremium = chartVtaProdLineaBarra.Series.Add("Premium");
            Series seriesDiesel = chartVtaProdLineaBarra.Series.Add("Diesel");

            seriesMagna.BorderWidth = Constantes.PathImagenes.GraficaAnchoDeLinea;
            seriesPremium.BorderWidth = Constantes.PathImagenes.GraficaAnchoDeLinea;
            seriesDiesel.BorderWidth = Constantes.PathImagenes.GraficaAnchoDeLinea;

            seriesMagna.Color = Color.Green;
            seriesPremium.Color = Color.Red;
            seriesDiesel.Color = Color.Black;

            seriesMagna.ChartType = chartType;
            seriesPremium.ChartType = chartType;
            seriesDiesel.ChartType = chartType;

            chartVtaProdLineaBarra.ChartAreas["ChartArea1"].Position.Auto = false;
            chartVtaProdLineaBarra.ChartAreas["ChartArea1"].Position.X = 0;
            chartVtaProdLineaBarra.ChartAreas["ChartArea1"].Position.Y = 20;
            chartVtaProdLineaBarra.ChartAreas["ChartArea1"].Position.Width = 90;
            chartVtaProdLineaBarra.ChartAreas["ChartArea1"].Position.Height = 100;
            chartVtaProdLineaBarra.ChartAreas["ChartArea1"].AxisY.LabelStyle.Format = "###,##0.00 Lts";
            chartVtaProdLineaBarra.ChartAreas["ChartArea1"].AxisX.MajorGrid.Interval = 1;
            chartVtaProdLineaBarra.ChartAreas["ChartArea1"].AxisX.LabelStyle.Interval = 1;

            for (int ii = 0; ii < 12; ii++)
            {
                string Mes = new DateTime(DateTime.Now.Year, CalcFechaIniYFinArr[ii].Item1.Month, 1).ToString("MMM",
                    CultureInfo.CreateSpecificCulture("es")).Replace(".", "") + " \n " + CalcFechaIniYFinArr[ii].Item1.Year; 

                seriesMagna.Points.AddXY(Mes,
                    LstDespachosEnEfectivoArr[ii].Where(r =>
                                                            r.PRODUCTO.Trim().ToLower().Contains("magna") ||
                                                            r.PRODUCTO.Trim().ToLower().Contains("regular")
                                                       ).
                    Sum(r => r.LITROS));

                seriesPremium.Points.AddXY(Mes,
                    LstDespachosEnEfectivoArr[ii].Where(r =>
                                                            r.PRODUCTO.Trim().ToLower().Contains("premium") ||
                                                            r.PRODUCTO.Trim().ToLower().Contains("supreme")
                                                       ).
                    Sum(r => r.LITROS));

                seriesDiesel.Points.AddXY(Mes,
                    LstDespachosEnEfectivoArr[ii].Where(r =>
                                                            r.PRODUCTO.Trim().ToLower().Contains("diesel") ||
                                                            r.PRODUCTO.Trim().ToLower().Contains("diesel-")
                                                       ).
                    Sum(r => r.LITROS));

                var lst =
                    LstDespachosEnEfectivoArr[ii].Where(r =>
                                                            r.PRODUCTO.Trim().ToLower().Contains("regular")
                                                       ).ToList();

                if (lst.Count() > 0)
                {
                }
            }

            chartVtaProdLineaBarra.Legends[0].LegendStyle = LegendStyle.Column;
            chartVtaProdLineaBarra.Legends[0].Position = new ElementPosition(90, 40, 10, 40);

            chartVtaProdLineaBarra.SaveImage(PathImagen, ChartImageFormat.Jpeg);
        }

        private void GraficasProductoEfectivoPorMes(Tuple<DateTime, DateTime>[] CalcFechaIniYFinArr,
            List<List<DatosDespachosFormPago>> LstDespachosEnEfectivo)
        {
            int IndVtaProdEfec = (int)Constantes.PathImagenes.PathsImgs.VtaProdEfec;

            CuerpoProductoPorMes(CalcFechaIniYFinArr, LstDespachosEnEfectivo,
                ref chartVtasProdEfec,
                Constantes.PathImagenes.ImagenesAIncluir[IndVtaProdEfec],
                SeriesChartType.Column);
        }

        private void GraficasProductoEfectivoPorMes(Tuple<DateTime, DateTime>[] CalcFechaIniYFinArr,
            List<List<DespachosContadoCreditoDebitoEfectivoMontoYLitros>> LstDespachosEnEfectivo)
        {
            int IndVtaProdEfec = (int)Constantes.PathImagenes.PathsImgs.VtaProdEfec;

            CuerpoProductoPorMes(CalcFechaIniYFinArr, LstDespachosEnEfectivo,
                ref chartVtasProdEfec,
                Constantes.PathImagenes.ImagenesAIncluir[IndVtaProdEfec],
                SeriesChartType.Column);
        }

        private void CuerpoEstacionPorMes(Tuple<DateTime, DateTime>[] CalcFechaIniYFinArr,
            List<List<DatosDespachosFormPago>> LstDespachosEnEfectivoArr, ref Chart chartVtaEstacLineaBarra,
            string PathImagen, SeriesChartType chartType)
        {
            chartVtaEstacLineaBarra.Series.Clear();
            chartVtaEstacLineaBarra.Size = new Size(
                Constantes.TamanioGrafica.Tamanio[(int)Constantes.TamanioGrafica.PropTam.Ancho],
                Constantes.TamanioGrafica.Tamanio[(int)Constantes.TamanioGrafica.PropTam.Alto]
            );

            Series seriesMieleras = chartVtaEstacLineaBarra.Series.Add("MI");
            Series seriesSeisEnero = chartVtaEstacLineaBarra.Series.Add("SE");
            Series seriesStaFe = chartVtaEstacLineaBarra.Series.Add("SF");
            Series seriesBravo = chartVtaEstacLineaBarra.Series.Add("BR");
            Series seriesCuenca1 = chartVtaEstacLineaBarra.Series.Add("C1");

            Series seriesCuenca2 = chartVtaEstacLineaBarra.Series.Add("C2");
            Series seriesSanJoaq = chartVtaEstacLineaBarra.Series.Add("SJ");
            Series seriesAzulejos = chartVtaEstacLineaBarra.Series.Add("AZ");
            Series seriesUrquizo = chartVtaEstacLineaBarra.Series.Add("UR");
            Series seriesStaRita = chartVtaEstacLineaBarra.Series.Add("SR");

            Series seriesParqueInd = chartVtaEstacLineaBarra.Series.Add("PI");
            Series series5deMayo = chartVtaEstacLineaBarra.Series.Add("FI");
            Series seriesCrit = chartVtaEstacLineaBarra.Series.Add("CR");
            Series seriesTriangulo = chartVtaEstacLineaBarra.Series.Add("TR");
            Series seriesIndependencia = chartVtaEstacLineaBarra.Series.Add("IN");

            Series seriesPuenteCentenario = chartVtaEstacLineaBarra.Series.Add("PC");

            seriesMieleras.BorderWidth = Constantes.PathImagenes.GraficaAnchoDeLinea;
            seriesSeisEnero.BorderWidth = Constantes.PathImagenes.GraficaAnchoDeLinea;
            seriesStaFe.BorderWidth = Constantes.PathImagenes.GraficaAnchoDeLinea;
            seriesBravo.BorderWidth = Constantes.PathImagenes.GraficaAnchoDeLinea;
            seriesCuenca1.BorderWidth = Constantes.PathImagenes.GraficaAnchoDeLinea;

            seriesCuenca2.BorderWidth = Constantes.PathImagenes.GraficaAnchoDeLinea;
            seriesSanJoaq.BorderWidth = Constantes.PathImagenes.GraficaAnchoDeLinea;
            seriesAzulejos.BorderWidth = Constantes.PathImagenes.GraficaAnchoDeLinea;
            seriesUrquizo.BorderWidth = Constantes.PathImagenes.GraficaAnchoDeLinea;
            seriesStaRita.BorderWidth = Constantes.PathImagenes.GraficaAnchoDeLinea;

            seriesParqueInd.BorderWidth = Constantes.PathImagenes.GraficaAnchoDeLinea;
            series5deMayo.BorderWidth = Constantes.PathImagenes.GraficaAnchoDeLinea;
            seriesCrit.BorderWidth = Constantes.PathImagenes.GraficaAnchoDeLinea;
            seriesTriangulo.BorderWidth = Constantes.PathImagenes.GraficaAnchoDeLinea;
            seriesIndependencia.BorderWidth = Constantes.PathImagenes.GraficaAnchoDeLinea;

            seriesPuenteCentenario.BorderWidth = Constantes.PathImagenes.GraficaAnchoDeLinea;

            seriesBravo.Color = Color.LightGray;
            seriesAzulejos.Color = Color.Fuchsia;
            seriesCuenca1.Color = Color.LightSkyBlue;
            seriesCuenca2.Color = Color.Blue;
            seriesSanJoaq.Color = Color.Orange;
            seriesUrquizo.Color = Color.Red;
            seriesStaRita.Color = Color.LightGray;
            series5deMayo.Color = Color.LightPink;

            seriesCuenca2.BorderDashStyle = ChartDashStyle.Dash;
            seriesSanJoaq.BorderDashStyle = ChartDashStyle.Dash;
            seriesUrquizo.BorderDashStyle = ChartDashStyle.Dash;
            seriesStaRita.BorderDashStyle = ChartDashStyle.Dash;

            seriesMieleras.ChartType = chartType;
            seriesSeisEnero.ChartType = chartType;
            seriesStaFe.ChartType = chartType;
            seriesBravo.ChartType = chartType;
            seriesCuenca1.ChartType = chartType;

            seriesCuenca2.ChartType = chartType;
            seriesSanJoaq.ChartType = chartType;
            seriesAzulejos.ChartType = chartType;
            seriesUrquizo.ChartType = chartType;
            seriesStaRita.ChartType = chartType;

            seriesParqueInd.ChartType = chartType;
            series5deMayo.ChartType = chartType;
            seriesCrit.ChartType = chartType;
            seriesTriangulo.ChartType = chartType;
            seriesIndependencia.ChartType = chartType;

            seriesPuenteCentenario.ChartType = chartType;

            chartVtaEstacLineaBarra.ChartAreas["ChartArea1"].Position.Auto = false;
            chartVtaEstacLineaBarra.ChartAreas["ChartArea1"].Position.X = 0;
            chartVtaEstacLineaBarra.ChartAreas["ChartArea1"].Position.Y = 20;
            chartVtaEstacLineaBarra.ChartAreas["ChartArea1"].Position.Width = 90;
            chartVtaEstacLineaBarra.ChartAreas["ChartArea1"].Position.Height = 100;
            chartVtaEstacLineaBarra.ChartAreas["ChartArea1"].AxisY.LabelStyle.Format = "###,##0.00 Lts";

            double Result = 0;
            for (int ii = 0; ii < 12; ii++)
            {
                string Mes = new DateTime(DateTime.Now.Year, CalcFechaIniYFinArr[ii].Item1.Month, 1).ToString("MMMM",
                    CultureInfo.CreateSpecificCulture("es"));

                seriesMieleras.Points.AddXY(Mes, Result =
                    LstDespachosEnEfectivoArr[ii].Where(r => r.Estacion.Trim().ToLower().Contains("mieleras")).
                   Sum(r => r.Litros));

                seriesSeisEnero.Points.AddXY(Mes, Result =
                    LstDespachosEnEfectivoArr[ii].Where(r => r.Estacion.Trim().ToLower().Contains("seis")).
                    Sum(r => r.Litros));

                seriesStaFe.Points.AddXY(Mes, Result =
                    LstDespachosEnEfectivoArr[ii].Where(r => r.Estacion.Trim().ToLower().Contains("santa fe")).
                    Sum(r => r.Litros));

                seriesBravo.Points.AddXY(Mes, Result =
                    LstDespachosEnEfectivoArr[ii].Where(r => r.Estacion.Trim().ToLower().Contains("bravo")).
                    Sum(r => r.Litros));

                seriesCuenca1.Points.AddXY(Mes, Result =
                    LstDespachosEnEfectivoArr[ii].Where(r => r.Estacion.Trim().ToLower().Contains("cuencame 1")).
                    Sum(r => r.Litros));

                seriesCuenca2.Points.AddXY(Mes, Result =
                    LstDespachosEnEfectivoArr[ii].Where(r => r.Estacion.Trim().ToLower().Contains("cuencame 2")).
                    Sum(r => r.Litros));

                seriesSanJoaq.Points.AddXY(Mes, Result =
                    LstDespachosEnEfectivoArr[ii].Where(r => r.Estacion.Trim().ToLower().Contains("san joaquin")).
                    Sum(r => r.Litros));

                seriesAzulejos.Points.AddXY(Mes, Result =
                    LstDespachosEnEfectivoArr[ii].Where(r => r.Estacion.Trim().ToLower().Contains("azulejos")).
                    Sum(r => r.Litros));

                seriesUrquizo.Points.AddXY(Mes, Result =
                    LstDespachosEnEfectivoArr[ii].Where(r => r.Estacion.Trim().ToLower().Contains("urquizo")).
                    Sum(r => r.Litros));

                seriesStaRita.Points.AddXY(Mes, Result =
                    LstDespachosEnEfectivoArr[ii].Where(r => r.Estacion.Trim().ToLower().Contains("santa rita")).
                    Sum(r => r.Litros));

                series5deMayo.Points.AddXY(Mes, Result =
                    LstDespachosEnEfectivoArr[ii].Where(r => r.Estacion.Trim().ToLower().Contains("parque")).
                    Sum(r => r.Litros));

                series5deMayo.Points.AddXY(Mes, Result =
                    LstDespachosEnEfectivoArr[ii].Where(r => r.Estacion.Trim().ToLower().Contains("filadelfia")).
                    Sum(r => r.Litros));

                series5deMayo.Points.AddXY(Mes, Result =
                    LstDespachosEnEfectivoArr[ii].Where(r => r.Estacion.Trim().ToLower().Contains("crit")).
                    Sum(r => r.Litros));

                series5deMayo.Points.AddXY(Mes, Result =
                    LstDespachosEnEfectivoArr[ii].Where(r => r.Estacion.Trim().ToLower().Contains("triangulo")).
                    Sum(r => r.Litros));

                series5deMayo.Points.AddXY(Mes, Result =
                    LstDespachosEnEfectivoArr[ii].Where(r => r.Estacion.Trim().ToLower().Contains("indepen")).
                    Sum(r => r.Litros));

                series5deMayo.Points.AddXY(Mes, Result =
                    LstDespachosEnEfectivoArr[ii].Where(r => r.Estacion.Trim().ToLower().Contains("puente cente")).
                    Sum(r => r.Litros));
            }

            chartVtaEstacLineaBarra.Legends[0].LegendStyle = LegendStyle.Column;
            chartVtaEstacLineaBarra.Legends[0].Position = new ElementPosition(90, 20, 10, 60);
            chartVtaEstacLineaBarra.Legends[0].AutoFitMinFontSize = 9;

            chartVtaEstacLineaBarra.SaveImage(PathImagen, ChartImageFormat.Jpeg);
        }

        private Series[] CrearSeriesGraficas(ref Chart chartVtaEstacLineaBarra,
            SeriesChartType chartType, bool PrimeraMitad)
        {
            int ii = 0;
            Chart chart = chartVtaEstacLineaBarra;
            Func<bool, int> ValInicio = r => r ? 0 : Constantes.EstacionesAbreviadas.Estaciones.Count() / 2;

            chartVtaEstacLineaBarra.Series.Clear();
            chartVtaEstacLineaBarra.Size = new Size(
                Constantes.TamanioGrafica.Tamanio[(int)Constantes.TamanioGrafica.PropTam.Ancho],
                Constantes.TamanioGrafica.Tamanio[(int)Constantes.TamanioGrafica.PropTam.Alto]
            );

            ii = 0;
            List<Series> series = new Series[Constantes.EstacionesAbreviadas.Estaciones.Count()].ToList().Select(r =>
            {
                {
                    r = chart.Series.Add(Constantes.EstacionesAbreviadas.Estaciones[ii]);
                    r.BorderWidth = Constantes.PathImagenes.GraficaAnchoDeLinea;
                    r.Color = Constantes.ColoresGraficas.Colores[ii];
                    r.BorderDashStyle = ChartDashStyle.Dash;
                    r.ChartType = chartType;

                    ii++;

                    return r;
                }
            }).ToList();

            return series.ToArray();
        }

        private void CuerpoEstacionPorMes(Tuple<DateTime, DateTime>[] CalcFechaIniYFinArr,
            List<List<DespachosContadoCreditoDebitoEfectivoMontoYLitros>> LstDespachosEnEfectivoArr, ref Chart chartVtaEstacLineaBarra,
            string PathImagen, SeriesChartType chartType, bool PrimerMitad)
        {
            Func<bool, int> ValInicio = r => r ? 0 : LstDespachosEnEfectivoArr.Count() / 2;
            Func<bool, int> FinFor = r => r ? LstDespachosEnEfectivoArr.Count() / 2 : LstDespachosEnEfectivoArr.Count();

            chartVtaEstacLineaBarra.Series.Clear();
            chartVtaEstacLineaBarra.Size = new Size(
                Constantes.TamanioGrafica.Tamanio[(int)Constantes.TamanioGrafica.PropTam.Ancho],
                Constantes.TamanioGrafica.Tamanio[(int)Constantes.TamanioGrafica.PropTam.Alto]
            );

            Series[] series = CrearSeriesGraficas(ref chartVtaEstacLineaBarra, chartType, PrimerMitad);

            chartVtaEstacLineaBarra.ChartAreas["ChartArea1"].Position.Auto = false;
            chartVtaEstacLineaBarra.ChartAreas["ChartArea1"].Position.X = 0;
            chartVtaEstacLineaBarra.ChartAreas["ChartArea1"].Position.Y = 20;
            chartVtaEstacLineaBarra.ChartAreas["ChartArea1"].Position.Width = 90;
            chartVtaEstacLineaBarra.ChartAreas["ChartArea1"].Position.Height = 100;
            chartVtaEstacLineaBarra.ChartAreas["ChartArea1"].AxisY.LabelStyle.Format = "###,##0.00 Lts";

            double Result = 0;
            for (int ii = ValInicio(PrimerMitad); ii < FinFor(PrimerMitad); ii++)
            {
                int jj = 0;

                string Mes = new DateTime(DateTime.Now.Year, CalcFechaIniYFinArr[ii].Item1.Month, 1).ToString("MMM",
                    CultureInfo.CreateSpecificCulture("es")).Replace(".", "") + " \n " + CalcFechaIniYFinArr[ii].Item1.Year;

                for (jj = 0; jj < Constantes.EstacionesAbreviadas.Estaciones.Count(); jj++)
                {
                    string Estacion = Constantes.EstacionesABuscar.Estaciones[jj];

                    var lst1 = LstDespachosEnEfectivoArr[ii].Where(r => r.GASOLINERA.Trim().ToLower().Contains("indep")).ToList();
                    var lst2 = LstDespachosEnEfectivoArr[ii].Where(r => r.GASOLINERA.Trim().ToLower().Contains("puente")).ToList();

                    Result = LstDespachosEnEfectivoArr[ii].Where(r => r.GASOLINERA.Trim().ToLower().
                                Contains(Estacion)).Sum(r => r.LITROS);

                    series[jj].Points.AddXY(Mes, Result);
                }
            }

            chartVtaEstacLineaBarra.Legends[0].LegendStyle = LegendStyle.Column;
            chartVtaEstacLineaBarra.Legends[0].Position = new ElementPosition(90, 20, 10, 60);
            chartVtaEstacLineaBarra.Legends[0].AutoFitMinFontSize = 9;
            chartVtaEstacLineaBarra.ChartAreas["ChartArea1"].AxisY.Maximum = series.Select(r => r.Points.Select(s => s.YValues.Max()).Max()).Max();
            var Val = series.Select(r => r.Points.Select(s => s.YValues.Max()).Max()).Max();

            chartVtaEstacLineaBarra.SaveImage(PathImagen, ChartImageFormat.Jpeg);
        }

        private void GraficasEstacionEfectivoPorMes(Tuple<DateTime, DateTime>[] CalcFechaIniYFinArr,
            List<List<DatosDespachosFormPago>> LstDespachosEnEfectivoOYCredArr)
        {
            int IndVtaEstEfec = (int)Constantes.PathImagenes.PathsImgs.VtaEstEfec;

            CuerpoEstacionPorMes(CalcFechaIniYFinArr, LstDespachosEnEfectivoOYCredArr,
                ref chartVtasEstacionEfec,
                Constantes.PathImagenes.ImagenesAIncluir[IndVtaEstEfec],
                SeriesChartType.Column);
        }

        private void GraficasEstacionEfectivoPorMes(Tuple<DateTime, DateTime>[] CalcFechaIniYFinArr,
            List<List<DespachosContadoCreditoDebitoEfectivoMontoYLitros>> LstDespachosEnEfectivoOYCredArr,
                bool PrimeraMitad)
        {
            int IndVtaEstEfec = PrimeraMitad? 
                                    (int)Constantes.PathImagenes.PathsImgs.VtaEstEfec:
                                    (int)Constantes.PathImagenes.PathsImgs.VtaEstEfec2;

            CuerpoEstacionPorMes(CalcFechaIniYFinArr, LstDespachosEnEfectivoOYCredArr,
                ref chartVtasEstacionEfec,
                Constantes.PathImagenes.ImagenesAIncluir[IndVtaEstEfec],
                SeriesChartType.Column, PrimeraMitad);
        }

        private void GraficasEstacionCredPorMes(Tuple<DateTime, DateTime>[] CalcFechaIniYFinArr,
            List<List<DatosDespachosFormPago>> LstDespachosEnCredArr)
        {
            int IndVtaEstCred = (int)Constantes.PathImagenes.PathsImgs.VtaEstCred;

            CuerpoEstacionPorMes(CalcFechaIniYFinArr, LstDespachosEnCredArr,
                ref chartVtasEstacionCred,
                Constantes.PathImagenes.ImagenesAIncluir[IndVtaEstCred],
                SeriesChartType.Column);
        }

        private void GraficasEstacionCredPorMes(Tuple<DateTime, DateTime>[] CalcFechaIniYFinArr,
            List<List<DespachosContadoCreditoDebitoEfectivoMontoYLitros>> LstDespachosEnCredArr, bool PrimeraMitad)
        {
            int IndVtaEstCred = 0;

            IndVtaEstCred = PrimeraMitad? 
                                (int)Constantes.PathImagenes.PathsImgs.VtaEstCred: 
                                (int)Constantes.PathImagenes.PathsImgs.VtaEstCred2;

            CuerpoEstacionPorMes(CalcFechaIniYFinArr, LstDespachosEnCredArr,
                ref chartVtasEstacionCred,
                Constantes.PathImagenes.ImagenesAIncluir[IndVtaEstCred],
                SeriesChartType.Column, PrimeraMitad);
        }

        private void GraficasProductoCredPorMes(Tuple<DateTime, DateTime>[] CalcFechaIniYFinArr,
            List<List<DatosDespachosFormPago>> LstDespachosCredArr)
        {
            int IndVtaProdCred = (int)Constantes.PathImagenes.PathsImgs.VtaProdCred;

            CuerpoProductoPorMes(CalcFechaIniYFinArr, LstDespachosCredArr,
                ref chartVtasProdCred,
                Constantes.PathImagenes.ImagenesAIncluir[IndVtaProdCred],
                SeriesChartType.Column);
        }

        private void GraficasProductoCredPorMes(Tuple<DateTime, DateTime>[] CalcFechaIniYFinArr,
            List<List<DespachosContadoCreditoDebitoEfectivoMontoYLitros>> LstDespachosCredArr)
        {
            int IndVtaProdCred = (int)Constantes.PathImagenes.PathsImgs.VtaProdCred;

            CuerpoProductoPorMes(CalcFechaIniYFinArr, LstDespachosCredArr,
                ref chartVtasProdCred,
                Constantes.PathImagenes.ImagenesAIncluir[IndVtaProdCred],
                SeriesChartType.Column);
        }

        private string DarProductos()
        {
            string resultado =
                    ((int)Productos.Magna).ToString() + "," +
                    ((int)Productos.Premium).ToString() + "," +
                    ((int)Productos.Diesel).ToString();

            return resultado;
        }

        private string DarPagoEfectivo()
        {
            string resultado =
                    ((int)TipoPago.Efectivo).ToString() + "," +
                    ((int)TipoPago.Debito).ToString() + "," +
                    ((int)TipoPago.Contado).ToString();

            return resultado;
        }

        private string DarPagoCredito()
        {
            string resultado =
                    ((int)TipoPago.Credito1).ToString() + "," +
                    ((int)TipoPago.Credito2).ToString();

            return resultado;
        }

        private string DarPagoEfectYCred()
        {
            string resultado =
                    ((int)TipoPago.Efectivo).ToString() + "," +
                    ((int)TipoPago.Debito).ToString() + "," +
                    ((int)TipoPago.Contado).ToString() + "," +
                    ((int)TipoPago.Transf_Banc).ToString() + "," +
                    ((int)TipoPago.Credito1).ToString() + "," +
                    ((int)TipoPago.Credito2).ToString();

            return resultado;
        }

        private List<DatosDespachosFormPago> ContadoConvTipoDespachoADespachoSegFormaPago(DateTime FechaI, DateTime FechaF)
        {
            Despachos despacho = new Despachos();
            CostoVentas costoVentas = new CostoVentas();

            List<FormsCostoVentas.Clases.DatosGasolinera> LstGasolineras = costoVentas.ObtDatosGasolineras(true);
            List<DatosDespachosFormPago> LstAgregDespachosFormPago = new List<DatosDespachosFormPago>();

            foreach (var elem1 in LstGasolineras)
            {
                LstAgregDespachosFormPago.AddRange(
                    despacho.ObtDespachosContado(FechaI, FechaF, elem1.Codigo).Select(r => new DatosDespachosFormPago
                    {
                        Estacion = r.GASOLINERA,
                        NumeroTurno = r.NROTRN,
                        Fecha = (int)r.FCHTRN.ToOADate() - 1,
                        Hora = r.HRATRN,
                        NroBom = 0,
                        Producto = r.PRODUCTO,
                        Litros = r.CAN,
                        Importe = r.MTO,
                        NombreCliente = "",
                    }).ToList());

                LstAgregDespachosFormPago.AddRange(
                    despacho.ObtDespachosEfectivo(FechaI, FechaF, elem1.Codigo).Select(r => new DatosDespachosFormPago
                    {
                        Estacion = r.GASOLINERA,
                        NumeroTurno = r.NROTRN,
                        Fecha = (int)r.FCHTRN.ToOADate() - 1,
                        Hora = r.HRATRN,
                        NroBom = 0,
                        Producto = r.PRODUCTO,
                        Litros = r.CAN,
                        Importe = r.MTO,
                        NombreCliente = "",
                    }).ToList());

                LstAgregDespachosFormPago.AddRange(
                    despacho.ObtDespachosDebito(FechaI, FechaF, elem1.Codigo).Select(r => new DatosDespachosFormPago
                    {
                        Estacion = r.GASOLINERA,
                        NumeroTurno = r.NROTRN,
                        Fecha = (int)r.FCHTRN.ToOADate() - 1,
                        Hora = r.HRATRN,
                        NroBom = 0,
                        Producto = r.PRODUCTO,
                        Litros = r.CAN,
                        Importe = r.MTO,
                        NombreCliente = "",
                    }).ToList());
            }

            return LstAgregDespachosFormPago;
        }

        private List<DatosDespachosFormPago> CreditoConvTipoDespachoADespachoSegFormaPago(DateTime FechaI, DateTime FechaF)
        {
            Despachos despacho = new Despachos();
            CostoVentas costoVentas = new CostoVentas();

            List<FormsCostoVentas.Clases.DatosGasolinera> LstGasolineras = costoVentas.ObtDatosGasolineras(true);
            List<DatosDespachosFormPago> LstAgregDespachosFormPago = new List<DatosDespachosFormPago>();

            foreach (var elem1 in LstGasolineras)
            {
                LstAgregDespachosFormPago.AddRange(
                    despacho.ObtDespachosCredito(FechaI, FechaF, elem1.Codigo).Select(r => new DatosDespachosFormPago
                    {
                        Estacion = r.GASOLINERA,
                        NumeroTurno = r.NROTRN,
                        Fecha = (int)r.FCHTRN.ToOADate() - 1,
                        Hora = r.HRATRN,
                        NroBom = 0,
                        Producto = r.PRODUCTO,
                        Litros = r.CAN,
                        Importe = r.MTO,
                        NombreCliente = "",
                    }).ToList());
            }

            return LstAgregDespachosFormPago;
        }

        private List<DatosDespachosFormPago> TodoConvTipoDespachoADespachoSegFormaPago(DateTime FechaI, DateTime FechaF)
        {
            Despachos despacho = new Despachos();
            CostoVentas costoVentas = new CostoVentas();

            List<FormsCostoVentas.Clases.DatosGasolinera> LstGasolineras = costoVentas.ObtDatosGasolineras(true);
            List<DatosDespachosFormPago> LstAgregDespachosFormPago = new List<DatosDespachosFormPago>();

            foreach (var elem1 in LstGasolineras)
            {
                LstAgregDespachosFormPago.AddRange(
                    despacho.ObtVerDespachadorSegNumDespacho(FechaI, FechaF).Select(r => new DatosDespachosFormPago
                    {
                        Estacion = r.GASOLINERA,
                        NumeroTurno = r.DESPACHO,
                        Fecha = (int)r.FECHA_TRANSACCION.ToOADate() - 1,
                        Hora = 0,
                        NroBom = 0,
                        Producto = r.PRODUCTO,
                        Litros = r.LITROS,
                        Importe = r.IMPORTE,
                        NombreCliente = "",
                    }).ToList());

            }

            return LstAgregDespachosFormPago;
        }
       
        private void CargarInfoCG(bool VtasEfectivo, bool VtasCredito, bool VtasEfectYCred)
        {
            Despachos despacho = new Despachos();

            Tuple<DateTime, DateTime>[] CalcFechaIniYFinArr = new Tuple<DateTime, DateTime>[12];
            List<DatosDespachosFormPago> LstAgregDespachosFormPago = null;
            List<DespachosContadoCreditoDebitoEfectivoMontoYLitros> LstDespachosContadoMontoYLitros = null;
            List<List<DespachosContadoCreditoDebitoEfectivoMontoYLitros>> LstDespachosContadoMontoYLitrosArr = 
                new List<List<DespachosContadoCreditoDebitoEfectivoMontoYLitros>>();
            List<List<DatosDespachosFormPago>> LstDespachosSegunPagoArr =
                new List<List<DatosDespachosFormPago>>();

            int ii = 0;
            Random rand = new Random();

            for (; ii < 12; ii++)
            {
                bool[] bEfectivo = new bool[3];
                bool bCredito = false;
                bool[] bEfectivoYCredito = new bool[4];

                CalcFechaIniYFinArr[ii] = EntregarMesIniYMesFin(12 - ii, 1);
                LstAgregDespachosFormPago = new List<DatosDespachosFormPago>();
                LstDespachosContadoMontoYLitros = new List<DespachosContadoCreditoDebitoEfectivoMontoYLitros>();

                bool Salir = true;

                do
                {
                    try
                    {
                        Salir = true;

                        //LstAgregDespachosFormPago.AddRange(
                        //despacho.ObtDespachosFormPago(
                        //    CalcFechaIniYFinArr[ii].Item1,
                        //    CalcFechaIniYFinArr[ii].Item2,
                        //    DarProductos(),

                        //    VtasEfectivo ?
                        //        DarPagoEfectivo() :
                        //        VtasCredito ?
                        //            DarPagoCredito() :
                        //            DarPagoEfectYCred()
                        //).ToList());

                        if (VtasEfectivo)
                        {
                            //LstAgregDespachosFormPago.
                            //    AddRange(
                            //                ContadoConvTipoDespachoADespachoSegFormaPago
                            //                (
                            //                    CalcFechaIniYFinArr[ii].Item1, CalcFechaIniYFinArr[ii].Item2
                            //                )
                            //            );

                            if (!bEfectivo[0])
                            {
                                LstDespachosContadoMontoYLitros.AddRange(despacho.ObtDespachoContadoSoloMontoYLitros(1, CalcFechaIniYFinArr[ii].Item1, CalcFechaIniYFinArr[ii].Item2, true));
                                bEfectivo[0] = true;
                            }

                            if (!bEfectivo[1])
                            {
                                LstDespachosContadoMontoYLitros.AddRange(despacho.ObtDespachoEfectivoMontoYLitros(1, CalcFechaIniYFinArr[ii].Item1, CalcFechaIniYFinArr[ii].Item2, true));
                                bEfectivo[1] = true;
                            }

                            if (!bEfectivo[2])
                            {
                                LstDespachosContadoMontoYLitros.AddRange(despacho.ObtDespachoDebitoMontoYLitros(1, CalcFechaIniYFinArr[ii].Item1, CalcFechaIniYFinArr[ii].Item2, true));
                                bEfectivo[2] = true;
                            }
                        }

                        if (VtasCredito)
                        {
                            //LstAgregDespachosFormPago.
                            //    AddRange(
                            //                CreditoConvTipoDespachoADespachoSegFormaPago
                            //                (
                            //                    CalcFechaIniYFinArr[ii].Item1, CalcFechaIniYFinArr[ii].Item2
                            //                )
                            //            );

                            if (!bCredito)
                            {
                                LstDespachosContadoMontoYLitros.AddRange(despacho.ObtDespachoCreditoMontoYLitro(1, CalcFechaIniYFinArr[ii].Item1, CalcFechaIniYFinArr[ii].Item2, true));
                                bCredito = true;
                            }
                        }
 
                        if (VtasEfectYCred)
                        {
                            //LstAgregDespachosFormPago.
                            //    AddRange(
                            //                TodoConvTipoDespachoADespachoSegFormaPago
                            //                (
                            //                    CalcFechaIniYFinArr[ii].Item1, CalcFechaIniYFinArr[ii].Item2
                            //                )
                            //            );

                            if (!bEfectivoYCredito[0])
                            {
                                LstDespachosContadoMontoYLitros.AddRange(despacho.ObtDespachoContadoSoloMontoYLitros(1, CalcFechaIniYFinArr[ii].Item1, CalcFechaIniYFinArr[ii].Item2, true));
                                bEfectivoYCredito[0] = true;
                            }

                            if (!bEfectivoYCredito[1])
                            {
                                LstDespachosContadoMontoYLitros.AddRange(despacho.ObtDespachoEfectivoMontoYLitros(1, CalcFechaIniYFinArr[ii].Item1, CalcFechaIniYFinArr[ii].Item2, true));
                                bEfectivoYCredito[1] = true;
                            }

                            if (!bEfectivoYCredito[2])
                            {
                                LstDespachosContadoMontoYLitros.AddRange(despacho.ObtDespachoDebitoMontoYLitros(1, CalcFechaIniYFinArr[ii].Item1, CalcFechaIniYFinArr[ii].Item2, true));
                                bEfectivoYCredito[2] = true;
                            }

                            if (!bEfectivoYCredito[3])
                            {
                                LstDespachosContadoMontoYLitros.AddRange(despacho.ObtDespachoCreditoMontoYLitro(1, CalcFechaIniYFinArr[ii].Item1, CalcFechaIniYFinArr[ii].Item2, true));
                                bEfectivoYCredito[3] = true;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Salir = false;
                    }
                }
                while (!Salir);

                //LstDespachosSegunPagoArr.Add(LstAgregDespachosFormPago);
                LstDespachosContadoMontoYLitrosArr.Add(LstDespachosContadoMontoYLitros);
                Application.DoEvents();
            }

            if (VtasEfectivo)
            {
                //GraficasProductoEfectivoPorMes(CalcFechaIniYFinArr, LstDespachosSegunPagoArr);
                //GraficasEstacionEfectivoPorMes(CalcFechaIniYFinArr, LstDespachosSegunPagoArr);

                GraficasProductoEfectivoPorMes(CalcFechaIniYFinArr, LstDespachosContadoMontoYLitrosArr);
                GraficasEstacionEfectivoPorMes(CalcFechaIniYFinArr, LstDespachosContadoMontoYLitrosArr, true);
                GraficasEstacionEfectivoPorMes(CalcFechaIniYFinArr, LstDespachosContadoMontoYLitrosArr, false);
            }

            if (VtasCredito)
            {
                //GraficasProductoCredPorMes(CalcFechaIniYFinArr, LstDespachosSegunPagoArr);
                //GraficasEstacionCredPorMes(CalcFechaIniYFinArr, LstDespachosSegunPagoArr);

                GraficasProductoCredPorMes(CalcFechaIniYFinArr, LstDespachosContadoMontoYLitrosArr);
                GraficasEstacionCredPorMes(CalcFechaIniYFinArr, LstDespachosContadoMontoYLitrosArr, true);
                GraficasEstacionCredPorMes(CalcFechaIniYFinArr, LstDespachosContadoMontoYLitrosArr, false);
            }

            if (VtasEfectYCred)
            {
                //GraficasProductoEfectivoYCredPorMes(CalcFechaIniYFinArr, LstDespachosSegunPagoArr);
                //GraficasEstacionEfectivoYCredPorMes(CalcFechaIniYFinArr, LstDespachosSegunPagoArr);

                GraficasProductoEfectivoYCredPorMes(CalcFechaIniYFinArr, LstDespachosContadoMontoYLitrosArr);
                GraficasEstacionEfectivoYCredPorMes(CalcFechaIniYFinArr, LstDespachosContadoMontoYLitrosArr, true);
                GraficasEstacionEfectivoYCredPorMes(CalcFechaIniYFinArr, LstDespachosContadoMontoYLitrosArr, false);
            }
        }

        private void CargarInfoControlGasVtasEfectivo()
        {
            Despachos despacho = new Despachos();

            Tuple<DateTime, DateTime>[] CalcFechaIniYFinArr = new Tuple<DateTime, DateTime>[12];
            List<List<DatosDespachosFormPago>> LstDespachosEnEfectivoArr =
                new List<List<DatosDespachosFormPago>>();

            int ii = 0;
            Random rand = new Random();

            for (; ii < 12; ii++)
            {
                CalcFechaIniYFinArr[ii] = EntregarMesIniYMesFin(12 - ii, 1);

                List<DatosDespachosFormPago> LstAgregDespachosFormPago =
                    new List<DatosDespachosFormPago>();

                bool Salir = true;

                do
                {
                    try
                    {
                        Salir = true;

                        LstAgregDespachosFormPago.AddRange(
                        despacho.ObtDespachosFormPago(
                            CalcFechaIniYFinArr[ii].Item1,
                            CalcFechaIniYFinArr[ii].Item2,
                            "1,2,27",           // 1 = "Magna", 2 = "Premium", 27 = Diesel
                            "49,52,53,56"       // 49 = Efectivo, 52 = Débito, 53 = Contado
                        ).ToList());            // 56 = Transferencia Bancaria
                    }
                    catch
                    {
                        Salir = false;
                    }
                }
                while (!Salir);

                LstDespachosEnEfectivoArr.Add(LstAgregDespachosFormPago);

                Application.DoEvents();
            }

            GraficasProductoEfectivoPorMes(CalcFechaIniYFinArr, LstDespachosEnEfectivoArr);
            GraficasEstacionEfectivoPorMes(CalcFechaIniYFinArr, LstDespachosEnEfectivoArr);
        }

        private void CargarInfoControlGasVtasCredito()
        {
            Despachos despacho = new Despachos();

            Tuple<DateTime, DateTime>[] CalcFechaIniYFinArr = new Tuple<DateTime, DateTime>[12];
            List<List<DatosDespachosFormPago>> LstDespachosCreditoArr = new List<List<DatosDespachosFormPago>>();

            int ii = 0;
            Random rand = new Random();

            for (; ii < 12; ii++)
            {
                CalcFechaIniYFinArr[ii] = EntregarMesIniYMesFin(12 - ii, 1);

                List<DatosDespachosFormPago> LstAgregDespachosFormPago =
                    new List<DatosDespachosFormPago>();

                bool Salir = true;

                do
                {
                    try
                    {
                        Salir = true;

                        LstAgregDespachosFormPago.AddRange(
                        despacho.ObtDespachosFormPago(
                            CalcFechaIniYFinArr[ii].Item1,
                            CalcFechaIniYFinArr[ii].Item2,
                            "1,2,27",           // 1 = "Magna", 2 = "Premium", 27 = Diesel
                            "50,51"             // 50 = Crédito, 51 = Crédito
                        ).ToList());
                    }
                    catch
                    {
                        Salir = false;
                    }
                }
                while (!Salir);

                LstDespachosCreditoArr.Add(LstAgregDespachosFormPago);

                Application.DoEvents();
            }

            GraficasProductoCredPorMes(CalcFechaIniYFinArr, LstDespachosCreditoArr);
            GraficasEstacionCredPorMes(CalcFechaIniYFinArr, LstDespachosCreditoArr);
        }

        private void CargarInfoControlGasVtasEfectYCredito()
        {
            Despachos despacho = new Despachos();

            Tuple<DateTime, DateTime>[] CalcFechaIniYFinArr = new Tuple<DateTime, DateTime>[12];
            List<List<DatosDespachosFormPago>> LstDespachosCreditoArr = new List<List<DatosDespachosFormPago>>();

            int ii = 0;
            Random rand = new Random();

            for (; ii < 12; ii++)
            {
                CalcFechaIniYFinArr[ii] = EntregarMesIniYMesFin(12 - ii, 1);

                List<DatosDespachosFormPago> LstAgregDespachosFormPago =
                    new List<DatosDespachosFormPago>();

                bool Salir = true;

                do
                {
                    Salir = true;

                    try
                    {
                        LstAgregDespachosFormPago.AddRange(
                        despacho.ObtDespachosFormPago(
                            CalcFechaIniYFinArr[ii].Item1,
                            CalcFechaIniYFinArr[ii].Item2,
                            "1,2,27",           // 1 = "Magna", 2 = "Premium", 27 = Diesel
                            "49,52,53,56,50,51"     // 49 = Efectivo, 52 = Débito,  53 = Contado
                        ).ToList());                // 56 = Transferencia Bancaria, 50 = Crédito, 51 = Crédito
                    }
                    catch
                    {
                        Salir = false;
                    }
                }
                while (!Salir);

                LstDespachosCreditoArr.Add(LstAgregDespachosFormPago);

                Application.DoEvents();
            }

            GraficasProductoEfectivoYCredPorMes(CalcFechaIniYFinArr, LstDespachosCreditoArr);
            GraficasEstacionEfectivoYCredPorMes(CalcFechaIniYFinArr, LstDespachosCreditoArr);
        }

        private void GraficasProductoEfectivoYCredPorMes(Tuple<DateTime, DateTime>[] CalcFechaIniYFinArr,
            List<List<DatosDespachosFormPago>> LstDespachosCredArr)
        {
            int IndVtaProdEfecYCred = (int)Constantes.PathImagenes.PathsImgs.VtaProdEfecYCred;

            CuerpoProductoPorMes(CalcFechaIniYFinArr, LstDespachosCredArr,
                ref chartVtasProdEfecYCred,
                Constantes.PathImagenes.ImagenesAIncluir[IndVtaProdEfecYCred],
                SeriesChartType.Column);
        }

        private void GraficasProductoEfectivoYCredPorMes(Tuple<DateTime, DateTime>[] CalcFechaIniYFinArr,
            List<List<DespachosContadoCreditoDebitoEfectivoMontoYLitros>> LstDespachosCredArr)
        {
            int IndVtaProdEfecYCred = (int)Constantes.PathImagenes.PathsImgs.VtaProdEfecYCred;

            CuerpoProductoPorMes(CalcFechaIniYFinArr, LstDespachosCredArr,
                ref chartVtasProdEfecYCred,
                Constantes.PathImagenes.ImagenesAIncluir[IndVtaProdEfecYCred],
                SeriesChartType.Column);
        }

        private void GraficasEstacionEfectivoYCredPorMes(Tuple<DateTime, DateTime>[] CalcFechaIniYFinArr,
            List<List<DatosDespachosFormPago>> LstDespachosCredArr)
        {
            int VtaEstEfecYCred = (int)Constantes.PathImagenes.PathsImgs.VtaEstEfecYCred;

            CuerpoEstacionPorMes(CalcFechaIniYFinArr, LstDespachosCredArr,
                ref chartVtasEstacEfecYCred,
                Constantes.PathImagenes.ImagenesAIncluir[VtaEstEfecYCred],
                SeriesChartType.Column);
        }

        private void GraficasEstacionEfectivoYCredPorMes(Tuple<DateTime, DateTime>[] CalcFechaIniYFinArr,
            List<List<DespachosContadoCreditoDebitoEfectivoMontoYLitros>> LstDespachosCredArr, 
                bool PrimeraMitad)
        {
            int VtaEstEfecYCred = PrimeraMitad? 
                                    (int)Constantes.PathImagenes.PathsImgs.VtaEstEfecYCred:
                                    (int)Constantes.PathImagenes.PathsImgs.VtaEstEfecYCred2;

            CuerpoEstacionPorMes(CalcFechaIniYFinArr, LstDespachosCredArr,
                ref chartVtasEstacEfecYCred,
                Constantes.PathImagenes.ImagenesAIncluir[VtaEstEfecYCred],
                SeriesChartType.Column, PrimeraMitad);
        }

        private void IniciarFechas()
        {
            Fecha = DateTime.Now.AddDays(Constantes.DiasAnteriores).ToShortDateString();

            FechaDia = DateTime.Now.AddDays(Constantes.DiasAnteriores).ToString("dddd, dd aa MMMM bbb yyyy").
                Replace("aa", "de").Replace("bbb", "del");

            FechaMensAcum = DateTime.Now.AddDays(Constantes.DiasAnteriores).ToString("MMMM aaa yyyy").
                Replace("aaa", "del");

            InicioMes = new DateTime(DiaAntes.Year, DiaAntes.Month, 1);
        }

        private List<FormsCostoVentas.Clases.DatosVentaDelDia> ObtenerDatosGasolineras(bool DiarioOAcumulado,
            FormsCostoVentas.Clases.DatosGasolinera elemGasolineras)
        {
            CostoVentas costoVentas = new CostoVentas();

            Func<bool, int> InicioMesODiaAnterior = r => r ?
                (int)(DiaAntes.ToOADate() - 1) : (int)(InicioMes.ToOADate() - 1);

            List<FormsCostoVentas.Clases.DatosVentaDelDia> LstVentaDelDia = null;

            switch (elemGasolineras.Nombre)
            {
                case "Azulejos PM":
                    return
                        LstVentaDelDia =
                            costoVentas.ObtDatosVentasDelDiaAzulejos(
                                    InicioMesODiaAnterior(DiarioOAcumulado),
                                    (int)(DateTime.Parse(Fecha).ToOADate() - 1),
                                    elemGasolineras.Codigo
                            );

                case "Bravo PM":
                    return
                        LstVentaDelDia =
                            costoVentas.ObtDatosVentasDelDiaBravo(
                                    InicioMesODiaAnterior(DiarioOAcumulado),
                                    (int)(DateTime.Parse(Fecha).ToOADate() - 1),
                                    elemGasolineras.Codigo
                            );

                case "Crit":
                    return
                        LstVentaDelDia =
                            costoVentas.ObtDatosVentasDelDiaCRIT(
                                    InicioMesODiaAnterior(DiarioOAcumulado),
                                    (int)(DateTime.Parse(Fecha).ToOADate() - 1),
                                    elemGasolineras.Codigo
                            );

                case "Cuencame 1":
                    return
                        LstVentaDelDia=
                            costoVentas.ObtDatosVentasDelDiaCuenca1(
                                    InicioMesODiaAnterior(DiarioOAcumulado),
                                    (int)(DateTime.Parse(Fecha).ToOADate() - 1),
                                    elemGasolineras.Codigo
                            );

                case "Cuencame 2":
                    return
                        LstVentaDelDia =
                            costoVentas.ObtDatosVentasDelDiaCuenca2(
                                    InicioMesODiaAnterior(DiarioOAcumulado),
                                    (int)(DateTime.Parse(Fecha).ToOADate() - 1),
                                    elemGasolineras.Codigo
                            );

                case "GPI Parque Ind":
                    return
                        LstVentaDelDia =
                            costoVentas.ObtDatosVentasDelDiaParqueInd(
                                    InicioMesODiaAnterior(DiarioOAcumulado),
                                    (int)(DateTime.Parse(Fecha).ToOADate() - 1),
                                    elemGasolineras.Codigo
                            );

                case "Mieleras PM":
                    return
                        LstVentaDelDia =
                            costoVentas.ObtDatosVentasDelDiaMieleras(
                                    InicioMesODiaAnterior(DiarioOAcumulado),
                                    (int)(DateTime.Parse(Fecha).ToOADate() - 1),
                                    elemGasolineras.Codigo
                            );

                case "Santa Fe":
                    return
                        LstVentaDelDia =
                            costoVentas.ObtDatosVentasDelDiaSantaFe(
                                    InicioMesODiaAnterior(DiarioOAcumulado),
                                    (int)(DateTime.Parse(Fecha).ToOADate() - 1),
                                    elemGasolineras.Codigo
                            );

                case "San Joaquin":
                    return
                        LstVentaDelDia =
                            costoVentas.ObtDatosVentasDelDiaSanJoaquin(
                                    InicioMesODiaAnterior(DiarioOAcumulado),
                                    (int)(DateTime.Parse(Fecha).ToOADate() - 1),
                                    elemGasolineras.Codigo
                            );

                case "Santa Rita":
                    return
                        LstVentaDelDia =
                            costoVentas.ObtDatosVentasDelDiaSantaRita(
                                    InicioMesODiaAnterior(DiarioOAcumulado),
                                    (int)(DateTime.Parse(Fecha).ToOADate() - 1),
                                    elemGasolineras.Codigo
                            );

                case "Seis de Enero":
                    return
                        LstVentaDelDia =
                            costoVentas.ObtDatosVentasDelDia6DeEnero(
                                    InicioMesODiaAnterior(DiarioOAcumulado),
                                    (int)(DateTime.Parse(Fecha).ToOADate() - 1),
                                    elemGasolineras.Codigo
                            );

                case "SNM Filadelfia":
                    return
                        LstVentaDelDia =
                            costoVentas.ObtDatosVentasDelDiaFiladelfia(
                                    InicioMesODiaAnterior(DiarioOAcumulado),
                                    (int)(DateTime.Parse(Fecha).ToOADate() - 1),
                                    elemGasolineras.Codigo
                            );

                case "Triangulo":
                    return
                        LstVentaDelDia =
                            costoVentas.ObtDatosVentasDelDiaTriangulo(
                                    InicioMesODiaAnterior(DiarioOAcumulado),
                                    (int)(DateTime.Parse(Fecha).ToOADate() - 1),
                                    elemGasolineras.Codigo
                            );

                case "Urquizo":
                    return
                        LstVentaDelDia =
                            costoVentas.ObtDatosVentasDelDiaUrquizo(
                                    InicioMesODiaAnterior(DiarioOAcumulado),
                                    (int)(DateTime.Parse(Fecha).ToOADate() - 1),
                                    elemGasolineras.Codigo
                            );
            }

            return new List<FormsCostoVentas.Clases.DatosVentaDelDia>();
        }

        private void CargarInfoControlGas(bool DiarioOAcumulado, ref List<CuerpoMailEstacion> LstCuerpoMailEstacion)
        {
            CostoVentas costoVentas = new CostoVentas();
            Despachos despachos = new Despachos();

            Func<bool, int> InicioMesODiaAnterior = r => r ?
                (int)(DiaAntes.ToOADate() - 1) : (int)(InicioMes.ToOADate() - 1);

            Func<bool, DateTime> InicioMesODiaAnteriorF = r => r? DiaAntes : InicioMes;

            List<FormsCostoVentas.Clases.DatosGasolinera> LstGasolineras = 
                costoVentas.ObtDatosGasolineras(false).OrderBy(r => r.Nombre).ToList();

            bool Entre = false;

            string Estacion = string.Empty;
            double Litros = 0, MontoLitros = 0;
            double UnidadesLitros = 0, UnidadesMonto = 0;
            double ImporteContado = 0, ImporteCredito = 0;
            double LitrosContado = 0, LitrosCredito = 0;

            LstCuerpoMailEstacion.Clear();

            foreach (FormsCostoVentas.Clases.DatosGasolinera elemGasolineras in LstGasolineras)
            {
                //Esta lista se utiliza para obtener las ventas de las estaciones desde el corporativo
                List<FormsCostoVentas.Clases.DatosVentaDelDia> LstVentaDelDia = costoVentas.ObtDatosVentasDelDia(
                            InicioMesODiaAnterior(DiarioOAcumulado),
                            (int)(DateTime.Parse(Fecha).ToOADate() - 1),
                            elemGasolineras.Codigo
                        );

                DateTime FechaI = InicioMesODiaAnteriorF(DiarioOAcumulado);

                List<DatosDespachoTipoPagoV2> LstDespachosCredito =
                   despachos.ObtDespachosCredito(FechaI, DateTime.Parse(Fecha), elemGasolineras.Codigo);

                //Esta lista se utiliza para obtener las ventas de las estaciones desde cada BD de las estaciones
                //List<FormsCostoVentas.Clases.DatosVentaDelDia> LstVentaDelDia =
                //    new List<FormsCostoVentas.Clases.DatosVentaDelDia>(ObtenerDatosGasolineras(DiarioOAcumulado, elemGasolineras));

                Estacion = elemGasolineras.Nombre;

                foreach (FormsCostoVentas.Clases.DatosVentaDelDia elemVtasDia in LstVentaDelDia)
                {
                    Entre = true;

                    switch (elemVtasDia.Unidad)
                    {
                        case "H87":
                            UnidadesLitros += elemVtasDia.Cantidad;
                            UnidadesMonto += elemVtasDia.Monto;
                            break;

                        case "LTR":
                            Litros += elemVtasDia.Cantidad;
                            MontoLitros += elemVtasDia.Monto;
                            break;
                    }
                }

                ImporteCredito = LstDespachosCredito.Select(r => r.MTO).Sum();
                ImporteContado = Math.Abs(ImporteCredito - (UnidadesMonto + MontoLitros));

                LitrosCredito = LstDespachosCredito.Select(r => r.CAN).Sum();
                LitrosContado = Math.Abs(LitrosCredito - (Litros + UnidadesLitros));

                LstCuerpoMailEstacion.Add(new CuerpoMailEstacion
                {
                    Estacion = Estacion,

                    UnidadesLitros = Entre? UnidadesLitros : 0,
                    LitrosCredito = Entre? LitrosCredito : 0,
                    LitrosContado = Entre? LitrosContado : 0,
                    ImporteLitros = Entre? (Litros + UnidadesLitros) : 0,

                    UnidadesDinero = Entre? UnidadesMonto : 0,
                    CreditoDinero = Entre? ImporteCredito : 0,
                    ContadoDinero = Entre? ImporteContado : 0,
                    ImporteDinero = Entre? (UnidadesMonto + MontoLitros) : 0,

                    FechaEstacion = DateTime.Now.ToShortDateString(),
                });

                if (Entre)
                {
                    if (DiarioOAcumulado)
                    {
                        TOTALUL1 += UnidadesLitros;
                        TOTALCL1 += LitrosCredito;
                        TOTALEL1 += LitrosContado;
                        TOTALIL1 += Litros + UnidadesLitros;

                        TOTALUD1 += UnidadesMonto;
                        TOTALCD1 += ImporteCredito;
                        TOTALED1 += ImporteContado;
                        TOTALID1 += UnidadesMonto + MontoLitros;
                    }
                    else
                    {
                        TOTALULA1 += UnidadesLitros;
                        TOTALCLA1 += LitrosCredito;
                        TOTALELA1 += LitrosContado;
                        TOTALILA1 += Litros + UnidadesLitros;

                        TOTALUDA1 += UnidadesMonto;
                        TOTALCDA1 += ImporteCredito;
                        TOTALEDA1 += ImporteContado;
                        TOTALIDA1 += UnidadesMonto + MontoLitros;
                    }
                }

                UnidadesLitros = LitrosCredito = LitrosContado = 0;
                Litros = UnidadesLitros = UnidadesMonto = ImporteCredito = 0;
                ImporteContado = UnidadesMonto = MontoLitros = UnidadesLitros = 0;

                LitrosCredito = LitrosContado = Litros = UnidadesLitros = 0;
                UnidadesMonto = ImporteCredito = ImporteContado = 0;
                UnidadesMonto = MontoLitros = 0;

                Entre = false;
            }
        }

        private bool CargarSucursales(ref List<FormsMicrosipVentasDia.Clases.DatosAlmacenes> LstSucursales)
        {
            MicrosipVentasDia microsipVentaDia = new MicrosipVentasDia();
            List<DatosAlmacenes> LstAlmacenes = new List<DatosAlmacenes>();

            for (int ii = 0; ii < Constantes.CantSucursales; ii++)
            {
                try
                {
                    switch (ii)
                    {
                        case 0:
                            LstAlmacenes = microsipVentaDia.ObtMicrosipAlmacenCNT();
                            break;

                        case 1:
                            LstAlmacenes = microsipVentaDia.ObtMicrosipAlmacenAMD();
                            break;

                        case 2:
                            LstAlmacenes = microsipVentaDia.ObtMicrosipAlmacenETN();
                            break;

                        case 3:
                            LstAlmacenes = microsipVentaDia.ObtMicrosipAlmacenMIE();
                            break;

                        case 4:
                            LstAlmacenes = microsipVentaDia.ObtMicrosipAlmacenMTR();
                            break;

                        case 5:
                            LstAlmacenes = microsipVentaDia.ObtMicrosipAlmacenSTF();
                            break;

                        case 6:
                            LstAlmacenes = new List<DatosAlmacenes>() { new DatosAlmacenes {
                                ALMACEN_ID = 19,
                                NOMBRE = "POINT SAN JOSE DE VIÑEDO"
                            } };
                            break;
                    }
                }
                catch (Exception ex)
                {
                    //throw new Exception(ex.InnerException.Message);
                    continue;
                }

                LstSucursales.AddRange(LstAlmacenes);
            }

            if (LstSucursales.Count() == 0)
                return false;
            else
                return true;
        }

        private bool CargarInfoMicrosip(bool DiarioOAcumulado,
            ref List<CuerpoMailSucursal> LstCuerpoMailSucursal)
        {
            MicrosipVentasDia microsipVentaDia = new MicrosipVentasDia();

            Func<bool, string> InicioMesODiaAnterior = r => r ?
                Fecha : InicioMes.ToString("dddd, dd MMMM yyyy");

            List<FormsMicrosipVentasDia.Clases.DatosAlmacenes> LstSucursales =
                new List<FormsMicrosipVentasDia.Clases.DatosAlmacenes>();

            if (!CargarSucursales(ref LstSucursales))
                return false;
            
            LstCuerpoMailSucursal.Clear();

            foreach (FormsMicrosipVentasDia.Clases.DatosAlmacenes almacen in LstSucursales)
            {
                List<FormsMicrosipVentasDia.Clases.DatosVtaDia> LstVtasDia = null;

                try
                {
                    switch (almacen.NOMBRE)
                    {
                        case "POINT MIELERAS":
                            LstVtasDia = microsipVentaDia.
                                ObtFbMIEVtasDia(InicioMesODiaAnterior(DiarioOAcumulado), Fecha, "S", true,
                                    almacen.ALMACEN_ID);
                            break;

                        case "POINT SANTA FE":
                            LstVtasDia = microsipVentaDia.
                                ObtFbSTFVtasDia(InicioMesODiaAnterior(DiarioOAcumulado), Fecha, "S", true,
                                    almacen.ALMACEN_ID);
                            break;

                        case "POINT MONTE REAL":
                            LstVtasDia = microsipVentaDia.
                                ObtFbMTRVtasDia(InicioMesODiaAnterior(DiarioOAcumulado), Fecha, "S", true,
                                    almacen.ALMACEN_ID);
                            break;

                        case "POINT CENTENARIO":
                            LstVtasDia = microsipVentaDia.
                                ObtFbCNTVtasDia(InicioMesODiaAnterior(DiarioOAcumulado), Fecha, "S", true,
                                    almacen.ALMACEN_ID);
                            break;

                        case "POINT LA AMISTAD":
                            LstVtasDia = microsipVentaDia.
                                ObtFbAMDVtasDia(InicioMesODiaAnterior(DiarioOAcumulado), Fecha, "S", true,
                                    almacen.ALMACEN_ID);
                            break;

                        case "POINT LAS  ETNIAS":
                            LstVtasDia = microsipVentaDia.
                                ObtFbETNVtasDia(InicioMesODiaAnterior(DiarioOAcumulado), Fecha, "S", true,
                                    almacen.ALMACEN_ID);
                            break;

                        case "POINT SAN JOSE DE VIÑEDO":
                            //LstVtasDia = microsipVentaDia.
                            //    ObtVtasDiaV2SJV(FechaIni, FechaFin);

                            LstVtasDia = new List<DatosVtaDia>();
                            break;
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }

                long Turno1Caja1 = 0, Turno1Caja2 = 0, Turno2Caja1 = 0,
                    Turno2Caja2 = 0, Turno3Caja1 = 0, Turno3Caja2 = 0, TotalTrans = 0;

                decimal ImporteSucursal = 0, TicketPromedio = 0;
                decimal SubTotTurnos = 0;

                Turno1Caja1 = LstVtasDia.Where(r =>
                                                r.TURNO.ToLower().Contains("turno1") &&
                                                r.NOMBRE.ToLower().Contains("caja") &&
                                                r.NOMBRE.ToLower().Contains("1")
                                              ).FirstOrDefault()?.TRANSACCION ?? 0;

                Turno1Caja2 = LstVtasDia.Where(r =>
                                                r.TURNO.ToLower().Contains("turno1") &&
                                                r.NOMBRE.ToLower().Contains("caja") &&
                                                r.NOMBRE.ToLower().Contains("2")
                                              ).FirstOrDefault()?.TRANSACCION ?? 0;

                Turno2Caja1 = LstVtasDia.Where(r =>
                                                r.TURNO.ToLower().Contains("turno2") &&
                                                r.NOMBRE.ToLower().Contains("caja") &&
                                                r.NOMBRE.ToLower().Contains("1")
                                              ).FirstOrDefault()?.TRANSACCION ?? 0;

                Turno2Caja2 = LstVtasDia.Where(r =>
                                                r.TURNO.ToLower().Contains("turno2") &&
                                                r.NOMBRE.ToLower().Contains("caja") &&
                                                r.NOMBRE.ToLower().Contains("2")
                                              ).FirstOrDefault()?.TRANSACCION ?? 0;

                Turno3Caja1 = LstVtasDia.Where(r =>
                                                r.TURNO.ToLower().Contains("turno3") &&
                                                r.NOMBRE.ToLower().Contains("caja") &&
                                                r.NOMBRE.ToLower().Contains("1")
                                              ).FirstOrDefault()?.TRANSACCION ?? 0;

                Turno3Caja2 = LstVtasDia.Where(r =>
                                                r.TURNO.ToLower().Contains("turno3") &&
                                                r.NOMBRE.ToLower().Contains("caja") &&
                                                r.NOMBRE.ToLower().Contains("2")
                                              ).FirstOrDefault()?.TRANSACCION ?? 0;

                ImporteSucursal = ImporteSucursal = LstVtasDia.Where(r => r.TURNO.ToLower().Contains("total")).
                                        FirstOrDefault()?.IMPORTE ?? 0;

                SubTotTurnos = Turno1Caja1 + Turno1Caja2 +
                                Turno2Caja1 + Turno2Caja2 +
                                Turno3Caja1 + Turno3Caja2;

                TicketPromedio =
                                    (
                                        (decimal)ImporteSucursal /
                                            (decimal)(
                                                        SubTotTurnos == 0 ? 1 : SubTotTurnos
                                                        )
                                    );

                LstCuerpoMailSucursal.Add(
                    new CuerpoMailSucursal
                    {
                        Sucursal = almacen.NOMBRE,

                        Turno1Caja1 = Turno1Caja1,
                        Turno1Caja2 = Turno1Caja2,
                        Turno2Caja1 = Turno2Caja1,
                        Turno2Caja2 = Turno2Caja2,
                        Turno3Caja1 = Turno3Caja1,
                        Turno3Caja2 = Turno3Caja2,

                        TotalTrans = Turno1Caja1 + Turno1Caja2 + Turno2Caja1 + Turno2Caja2 + Turno3Caja1 + Turno3Caja2,

                        ImporteSucursal = ImporteSucursal,
                        TicketPromedio = TicketPromedio
                    }
                );

                if (DiarioOAcumulado)
                {
                    TT1C1 += Turno1Caja1; TT1C2 += Turno1Caja2;
                    TT2C1 += Turno2Caja1; TT2C2 += Turno2Caja2;
                    TT3C1 += Turno3Caja1; TT3C2 += Turno3Caja2;

                    TT += TotalTrans;
                    TTP += TicketPromedio;
                    TIS += ImporteSucursal;

                    TISP = TIS;
                }
                else
                {
                    TT1C1A += Turno1Caja1; TT1C2A += Turno1Caja2;
                    TT2C1A += Turno2Caja1; TT2C2A += Turno2Caja2;
                    TT3C1A += Turno3Caja1; TT2C2A += Turno3Caja2;

                    TTA += TotalTrans;
                    TTPA += TicketPromedio;
                    TISA += ImporteSucursal;

                    TISPA = TISA;
                }
            }

            if (DiarioOAcumulado)
                TISP = (decimal)TISP / (decimal)LstSucursales.Count();
            else
                TISPA = (decimal)TISPA / (decimal)LstSucursales.Count();

            return true;
        }

        private bool CargarInfoMicrosipHoraInicioTurno()
        {
            MicrosipVentasDia microsipVentaDia = new MicrosipVentasDia();

            Func<bool, string> InicioMesODiaAnterior = r => r ?
                Fecha : InicioMes.ToString("dddd, dd MMMM yyyy");

            List<FormsMicrosipVentasDia.Clases.DatosAlmacenes> LstSucursales =
                new List<FormsMicrosipVentasDia.Clases.DatosAlmacenes>();

            try
            {

                LstSucursales.AddRange(microsipVentaDia.ObtMicrosipAlmacenCNT());
                LstSucursales.AddRange(microsipVentaDia.ObtMicrosipAlmacenMIE());
                LstSucursales.AddRange(microsipVentaDia.ObtMicrosipAlmacenMTR());
                LstSucursales.AddRange(microsipVentaDia.ObtMicrosipAlmacenSTF());
                LstSucursales.AddRange(microsipVentaDia.ObtMicrosipAlmacenAMD());
                LstSucursales.AddRange(microsipVentaDia.ObtMicrosipAlmacenETN());
            }
            catch
            {
                if (LstSucursales.Count() == 0)
                    return false;
            }

            LstCuerpoMailSucursalTurno.Clear();

            foreach (FormsMicrosipVentasDia.Clases.DatosAlmacenes almacen in LstSucursales)
            {
                List<FormsMicrosipVentasDia.Clases.DatosPrimerHoraTurno> LstTurnos = null;

                switch (almacen.NOMBRE)
                {
                    case "POINT MIELERAS":
                        LstTurnos = microsipVentaDia.
                            ObtFbMIEPrimerHoraTurno(Fecha, Fecha);
                        break;

                    case "POINT SANTA FE":
                        LstTurnos = microsipVentaDia.
                            ObtFbSTFPrimerHoraTurno(Fecha, Fecha);
                        break;

                    case "POINT MONTE REAL":
                        LstTurnos = microsipVentaDia.
                            ObtFbMTRPrimerHoraTurno(Fecha, Fecha);
                        break;

                    case "POINT CENTENARIO":
                        LstTurnos = microsipVentaDia.
                            ObtFbCNTPrimerHoraTurno(Fecha, Fecha);
                        break;

                    case "POINT LA AMISTAD":
                        LstTurnos = microsipVentaDia.
                            ObtFbAMDPrimerHoraTurno(Fecha, Fecha);
                        break;

                    case "POINT LAS  ETNIAS":
                        LstTurnos = microsipVentaDia.
                            ObtFbETNPrimerHoraTurno(Fecha, Fecha);
                        break;
                }

                TimeSpan Turno1Caja1, Turno1Caja2, Turno2Caja1, Turno2Caja2, Turno3Caja1, Turno3Caja2;

                Turno1Caja1 = LstTurnos.Where(r =>
                                r.TURNO.ToLower().Contains("turno1") &&
                                r.NOMBRE.ToLower().Contains("caja") &&
                                r.NOMBRE.ToLower().Contains("1")
                              )?.FirstOrDefault()?.HORA ?? TimeSpan.MinValue;

                Turno1Caja2 = LstTurnos.Where(r =>
                                r.TURNO.ToLower().Contains("turno1") &&
                                r.NOMBRE.ToLower().Contains("caja") &&
                                r.NOMBRE.ToLower().Contains("2")
                             )?.FirstOrDefault()?.HORA ?? TimeSpan.MinValue;

                Turno2Caja1 = LstTurnos.Where(r =>
                                r.TURNO.ToLower().Contains("turno2") &&
                                r.NOMBRE.ToLower().Contains("caja") &&
                                r.NOMBRE.ToLower().Contains("1")
                                )?.FirstOrDefault()?.HORA ?? TimeSpan.MinValue;

                Turno2Caja2 = LstTurnos.Where(r =>
                                r.TURNO.ToLower().Contains("turno2") &&
                                r.NOMBRE.ToLower().Contains("caja") &&
                                r.NOMBRE.ToLower().Contains("2")
                                )?.FirstOrDefault()?.HORA ?? TimeSpan.MinValue;

                Turno3Caja1 = LstTurnos.Where(r =>
                                r.TURNO.ToLower().Contains("turno3") &&
                                r.NOMBRE.ToLower().Contains("caja") &&
                                r.NOMBRE.ToLower().Contains("1")
                                )?.FirstOrDefault()?.HORA ?? TimeSpan.MinValue;

                Turno3Caja2 = LstTurnos.Where(r =>
                                r.TURNO.ToLower().Contains("turno3") &&
                                r.NOMBRE.ToLower().Contains("caja") &&
                                r.NOMBRE.ToLower().Contains("2")
                                )?.FirstOrDefault()?.HORA ?? TimeSpan.MinValue;

                LstCuerpoMailSucursalTurno.Add(new CuerpoMailSucursalTurno
                {
                    Sucursal = almacen.NOMBRE,
                    Turno1Caja1 = Turno1Caja1,
                    Turno1Caja2 = Turno1Caja2,
                    Turno2Caja1 = Turno2Caja1,
                    Turno2Caja2 = Turno2Caja2,
                    Turno3Caja1 = Turno3Caja1,
                    Turno3Caja2 = Turno3Caja2
                });
            }

            return true;
        }

        private void ExcepcionNombreEstacion(ref List<CuerpoMailEstacion> LstCuerpoMailEstacion, List<string> Estaciones)
        {
            //Del nombre la primera letra mayúscula, las demás minúsculas
            LstCuerpoMailEstacion.Select(r => 
                Estaciones.Contains(r.Estacion)? 
                    r.Estacion = 
                        string.Concat(
                            r.Estacion.Select(
                                                (s, i) => i == 0? 
                                                            s.ToString().ToUpper() : 
                                                            s.ToString().ToLower()
                                             ).ToList()
                                     ) : 
                        ""
                    ).ToList();
        }

        private void ReemplazarVariablesEnCuerpoMailEstacion()
        {
            ExcepcionNombreEstacion(ref LstCuerpoMailEstacion, new List<string> { "CRIT" });

            List<string> Estaciones = new EstacionesYVariablesHTML().GenerarVariableHTML("E", Constantes.EstacOSucur.Estacion);
            List<string> UnidadesLitros = new EstacionesYVariablesHTML().GenerarVariableHTML("UL", Constantes.EstacOSucur.Estacion);
            List<string> LitrosCredito = new EstacionesYVariablesHTML().GenerarVariableHTML("CL", Constantes.EstacOSucur.Estacion);
            List<string> LitrosContado = new EstacionesYVariablesHTML().GenerarVariableHTML("EL", Constantes.EstacOSucur.Estacion); 
            List<string> ImporteLitros = new EstacionesYVariablesHTML().GenerarVariableHTML("IL", Constantes.EstacOSucur.Estacion);

            int ii = 0;
            foreach (CuerpoMailEstacion cuerpo in LstCuerpoMailEstacion)
            {
                string Estacion = string.Empty;

                try
                {
                    Estacion = Constantes.CorregirEstaciones[cuerpo.Estacion];
                }
                catch
                {
                    Estacion = cuerpo.Estacion;
                }

                TextPlantilla = TextPlantilla.Replace(Estaciones[ii], Estacion);
                TextPlantilla = TextPlantilla.Replace(UnidadesLitros[ii], cuerpo.UnidadesLitros.ToString("N2"));
                TextPlantilla = TextPlantilla.Replace(LitrosCredito[ii], cuerpo.LitrosCredito.ToString("N2"));
                TextPlantilla = TextPlantilla.Replace(LitrosContado[ii], cuerpo.LitrosContado.ToString("N2"));
                TextPlantilla = TextPlantilla.Replace(ImporteLitros[ii++], cuerpo.ImporteLitros.ToString("N2"));
            }

            if (ii < Constantes.CantEstaciones)
            {
                for (ii = 0; ii < Constantes.CantEstaciones; ii++)
                {
                    TextPlantilla = TextPlantilla.Replace(Estaciones[ii], "");
                    TextPlantilla = TextPlantilla.Replace(UnidadesLitros[ii], "");
                    TextPlantilla = TextPlantilla.Replace(LitrosCredito[ii], "");
                    TextPlantilla = TextPlantilla.Replace(LitrosContado[ii], "");
                    TextPlantilla = TextPlantilla.Replace(ImporteLitros[ii++], "");
                }
            }

            TextPlantilla = TextPlantilla.Replace("{Fecha}", FechaDia);
            TextPlantilla = TextPlantilla.Replace("{TOTALUL1}", TOTALUL1.ToString("###,##0.00"));
            TextPlantilla = TextPlantilla.Replace("{TOTALCL1}", TOTALCL1.ToString("###,##0.00"));
            TextPlantilla = TextPlantilla.Replace("{TOTALEL1}", TOTALEL1.ToString("###,##0.00"));
            TextPlantilla = TextPlantilla.Replace("{TOTALIL1}", TOTALIL1.ToString("###,##0.00"));
        }

        private void ReemplazarVariablesAcumuladasEnCuerpoMailEstacion()
        {
            List<string> Estaciones = new EstacionesYVariablesHTML().GenerarVariableHTML("E", Constantes.EstacOSucur.Estacion);
            List<string> UnidadesLitros = new EstacionesYVariablesHTML().GenerarVariableHTML("ULA", Constantes.EstacOSucur.Estacion);
            List<string> LitrosCredito = new EstacionesYVariablesHTML().GenerarVariableHTML("CLA", Constantes.EstacOSucur.Estacion);
            List<string> LitrosContado = new EstacionesYVariablesHTML().GenerarVariableHTML("ELA", Constantes.EstacOSucur.Estacion);
            List<string> ImporteLitros = new EstacionesYVariablesHTML().GenerarVariableHTML("ILA", Constantes.EstacOSucur.Estacion);

            int ii = 0;
            foreach (CuerpoMailEstacion cuerpo in LstCuerpoMailEstacionAcum)
            {
                string Estacion = string.Empty;

                try
                {
                    Estacion = Constantes.CorregirEstaciones[cuerpo.Estacion];
                }
                catch
                {
                    Estacion = cuerpo.Estacion;
                }

                TextPlantilla = TextPlantilla.Replace(Estaciones[ii], Estacion);
                TextPlantilla = TextPlantilla.Replace(UnidadesLitros[ii], cuerpo.UnidadesLitros.ToString("N2"));
                TextPlantilla = TextPlantilla.Replace(LitrosCredito[ii], cuerpo.LitrosCredito.ToString("N2"));
                TextPlantilla = TextPlantilla.Replace(LitrosContado[ii], cuerpo.LitrosContado.ToString("N2"));
                TextPlantilla = TextPlantilla.Replace(ImporteLitros[ii++], cuerpo.ImporteLitros.ToString("N2"));
            }

            if (ii < Constantes.CantEstaciones)
            {
                for (; ii < Constantes.CantEstaciones; ii++)
                {
                    TextPlantilla = TextPlantilla.Replace(Estaciones[ii], "");
                    TextPlantilla = TextPlantilla.Replace(UnidadesLitros[ii], "");
                    TextPlantilla = TextPlantilla.Replace(LitrosCredito[ii], "");
                    TextPlantilla = TextPlantilla.Replace(LitrosContado[ii], "");
                    TextPlantilla = TextPlantilla.Replace(ImporteLitros[ii++], "");
                }
            }

            TextPlantilla = TextPlantilla.Replace("{Fecha}", FechaDia);
            TextPlantilla = TextPlantilla.Replace("{TOTALULA1}", TOTALULA1.ToString("N2"));
            TextPlantilla = TextPlantilla.Replace("{TOTALCLA1}", TOTALCLA1.ToString("N2"));
            TextPlantilla = TextPlantilla.Replace("{TOTALELA1}", TOTALELA1.ToString("N2"));
            TextPlantilla = TextPlantilla.Replace("{TOTALILA1}", TOTALILA1.ToString("N2"));
        }

        private void ReemplazarVariablesEnCuerpoMailEstacionDinero()
        {
            ExcepcionNombreEstacion(ref LstCuerpoMailEstacion, new List<string> { "CRIT" });

            List<string> Estaciones = new EstacionesYVariablesHTML().GenerarVariableHTML("E", Constantes.EstacOSucur.Estacion);
            List<string> UnidadesContado = new EstacionesYVariablesHTML().GenerarVariableHTML("UD", Constantes.EstacOSucur.Estacion);
            List<string> ImporteCredito = new EstacionesYVariablesHTML().GenerarVariableHTML("CD", Constantes.EstacOSucur.Estacion);
            List<string> ImporteContado = new EstacionesYVariablesHTML().GenerarVariableHTML("ED", Constantes.EstacOSucur.Estacion);
            List<string> ImporteTotal = new EstacionesYVariablesHTML().GenerarVariableHTML("ID", Constantes.EstacOSucur.Estacion);

            int ii = 0;
            foreach (CuerpoMailEstacion cuerpo in LstCuerpoMailEstacion)
            {
                TextPlantilla = TextPlantilla.Replace(Estaciones[ii], cuerpo.Estacion);
                TextPlantilla = TextPlantilla.Replace(UnidadesContado[ii], cuerpo.UnidadesDinero.ToString("C2"));
                TextPlantilla = TextPlantilla.Replace(ImporteCredito[ii], cuerpo.CreditoDinero.ToString("C2"));
                TextPlantilla = TextPlantilla.Replace(ImporteContado[ii], cuerpo.ContadoDinero.ToString("C2"));
                TextPlantilla = TextPlantilla.Replace(ImporteTotal[ii++], cuerpo.ImporteDinero.ToString("C2"));
            }

            if (ii < Constantes.CantEstaciones)
            {
                for (ii = 0; ii < Constantes.CantEstaciones; ii++)
                {
                    TextPlantilla = TextPlantilla.Replace(Estaciones[ii], "");
                    TextPlantilla = TextPlantilla.Replace(UnidadesContado[ii], "");
                    TextPlantilla = TextPlantilla.Replace(ImporteCredito[ii], "");
                    TextPlantilla = TextPlantilla.Replace(ImporteContado[ii], "");
                    TextPlantilla = TextPlantilla.Replace(ImporteTotal[ii++], "");
                }
            }

            TextPlantilla = TextPlantilla.Replace("{Fecha}", FechaDia);
            TextPlantilla = TextPlantilla.Replace("{TOTALUD1}", TOTALUD1.ToString("C2"));
            TextPlantilla = TextPlantilla.Replace("{TOTALCD1}", TOTALCD1.ToString("C2"));
            TextPlantilla = TextPlantilla.Replace("{TOTALED1}", TOTALED1.ToString("C2"));
            TextPlantilla = TextPlantilla.Replace("{TOTALID1}", TOTALID1.ToString("C2"));
        }

        private void ReemplazarVariablesAcumuladasEnCuerpoMailEstacionDinero()
        {
            List<string> Estaciones = new EstacionesYVariablesHTML().GenerarVariableHTML("E", Constantes.EstacOSucur.Estacion);
            List<string> UnidadesContado = new EstacionesYVariablesHTML().GenerarVariableHTML("UDA", Constantes.EstacOSucur.Estacion);
            List<string> ImportesCredito = new EstacionesYVariablesHTML().GenerarVariableHTML("CDA", Constantes.EstacOSucur.Estacion);
            List<string> ImportesContado = new EstacionesYVariablesHTML().GenerarVariableHTML("EDA", Constantes.EstacOSucur.Estacion);
            List<string> ImporteTotal = new EstacionesYVariablesHTML().GenerarVariableHTML("IDA", Constantes.EstacOSucur.Estacion);

            int ii = 0;
            foreach (CuerpoMailEstacion cuerpo in LstCuerpoMailEstacionAcum)
            {
                TextPlantilla = TextPlantilla.Replace(Estaciones[ii], cuerpo.Estacion);
                TextPlantilla = TextPlantilla.Replace(UnidadesContado[ii], cuerpo.UnidadesDinero.ToString("C2"));
                TextPlantilla = TextPlantilla.Replace(ImportesCredito[ii], cuerpo.CreditoDinero.ToString("C2"));
                TextPlantilla = TextPlantilla.Replace(ImportesContado[ii], cuerpo.ContadoDinero.ToString("C2"));
                TextPlantilla = TextPlantilla.Replace(ImporteTotal[ii++], cuerpo.ImporteDinero.ToString("C2"));
            }

            if (ii < Constantes.CantEstaciones)
            {
                for (ii = 0; ii < Constantes.CantEstaciones; ii++)
                {
                    TextPlantilla = TextPlantilla.Replace(Estaciones[ii], "");
                    TextPlantilla = TextPlantilla.Replace(UnidadesContado[ii], "");
                    TextPlantilla = TextPlantilla.Replace(ImportesCredito[ii], "");
                    TextPlantilla = TextPlantilla.Replace(ImportesContado[ii], "");
                    TextPlantilla = TextPlantilla.Replace(ImporteTotal[ii++], "");
                }
            }

            TextPlantilla = TextPlantilla.Replace("{Fecha}", FechaDia);
            TextPlantilla = TextPlantilla.Replace("{TOTALUDA1}", TOTALUDA1.ToString("C2"));
            TextPlantilla = TextPlantilla.Replace("{TOTALCDA1}", TOTALCDA1.ToString("C2"));
            TextPlantilla = TextPlantilla.Replace("{TOTALEDA1}", TOTALEDA1.ToString("C2"));
            TextPlantilla = TextPlantilla.Replace("{TOTALIDA1}", TOTALIDA1.ToString("C2"));
        }

        private string EntregarResultTurnoCaja(int ii, CuerpoMailSucursal cuerpo)
        {
            switch (ii)
            {
                case 0:
                    return
                        cuerpo.Turno1Caja1 == -1? "" : cuerpo.Turno1Caja1.ToString("###,##0");
                case 1:
                    return
                        cuerpo.Turno1Caja2 == -1 ? "" : cuerpo.Turno1Caja2.ToString("###,##0");
                case 2:
                    return
                        cuerpo.Turno2Caja1 == -1 ? "" : cuerpo.Turno2Caja1.ToString("###,##0");
                case 3:
                    return
                        cuerpo.Turno2Caja2 == -1 ? "" : cuerpo.Turno2Caja2.ToString("###,##0");
                case 4:
                    return
                        cuerpo.Turno3Caja1 == -1 ? "" : cuerpo.Turno3Caja1.ToString("###,##0");
                case 5:
                    return
                        cuerpo.Turno3Caja2 == -1 ? "" : cuerpo.Turno3Caja2.ToString("###,##0");
            }

            return "";
        }

        private void ReemplazarRenglonesSucursales(ref string Plantilla, List<string> Renglon,
            CuerpoMailSucursal cuerpo, ref List<long> TurnoCaja)
        {
            int jj = 0;

            foreach (string elem in Renglon)
            {
                string valor = "";

                TextPlantilla = TextPlantilla.Replace(elem,
                    EntregarResultTurnoCaja(jj, cuerpo));

                TurnoCaja[jj] += 
                    (valor = EntregarResultTurnoCaja(jj++, cuerpo)) == ""? 0 : 
                        long.Parse(valor, NumberStyles.AllowThousands);
            }
        }

        private string DarFormatoTicketImporte(int Ind, string Dato)
        {
            switch (Ind)
            {
                case 0:
                    return long.Parse(Dato).ToString("###,##0");
                case 1:
                    return long.Parse(Dato).ToString("###,##0");
                case 2:
                    return long.Parse(Dato).ToString("###,##0");
                case 3:
                    return decimal.Parse(Dato).ToString("$#,###,##0.00");
                case 4:
                    return decimal.Parse(Dato).ToString("$#,###,##0.00");
            }

            return "";
        }

        private List<string> DarTransaccReng(int ii)
        {
            List<string> TransaccionesReng1 = Enum.GetNames(typeof(TransaccionesPorTurnoYCajaReng1)).ToList()
                                            .Select(r => "{" + r + "}").ToList();

            List<string> TransaccionesReng2 = Enum.GetNames(typeof(TransaccionesPorTurnoYCajaReng2)).ToList()
                                            .Select(r => "{" + r + "}").ToList();
            List<string> TransaccionesReng3 = Enum.GetNames(typeof(TransaccionesPorTurnoYCajaReng3)).ToList()
                                            .Select(r => "{" + r + "}").ToList();
            List<string> TransaccionesReng4 = Enum.GetNames(typeof(TransaccionesPorTurnoYCajaReng4)).ToList()
                                            .Select(r => "{" + r + "}").ToList();

            List<string> TransaccionesReng5 = Enum.GetNames(typeof(TransaccionesPorTurnoYCajaReng5)).ToList()
                                            .Select(r => "{" + r + "}").ToList();

            List<string> TransaccionesReng6 = Enum.GetNames(typeof(TransaccionesPorTurnoYCajaReng6)).ToList()
                                            .Select(r => "{" + r + "}").ToList();

            List<string> TransaccionesReng7 = Enum.GetNames(typeof(TransaccionesPorTurnoYCajaReng7)).ToList()
                                .Select(r => "{" + r + "}").ToList();

            switch (ii)
            {
                case 0:
                    return TransaccionesReng1;
                case 1:
                    return TransaccionesReng2;
                case 2:
                    return TransaccionesReng3;
                case 3:
                    return TransaccionesReng4;
                case 4:
                    return TransaccionesReng5;
                case 5:
                    return TransaccionesReng6;
                case 6:
                    return TransaccionesReng7;
            }

            return new List<string>();
        }

        private List<string> DarTransTicketsEImporteReng(int ii)
        {
            List<string> TransaccionesReng1 = Enum.GetNames(typeof(NoTicketsReng1)).ToList()
                                            .Select(r => "{" + r + "}").ToList();

            List<string> TransaccionesReng2 = Enum.GetNames(typeof(NoTicketsReng2)).ToList()
                                            .Select(r => "{" + r + "}").ToList();
            List<string> TransaccionesReng3 = Enum.GetNames(typeof(NoTicketsReng3)).ToList()
                                            .Select(r => "{" + r + "}").ToList();
            List<string> TransaccionesReng4 = Enum.GetNames(typeof(NoTicketsReng4)).ToList()
                                            .Select(r => "{" + r + "}").ToList();

            List<string> TransaccionesReng5 = Enum.GetNames(typeof(NoTicketsReng5)).ToList()
                                            .Select(r => "{" + r + "}").ToList();

            List<string> TransaccionesReng6 = Enum.GetNames(typeof(NoTicketsReng6)).ToList()
                                            .Select(r => "{" + r + "}").ToList();

            List<string> TransaccionesReng7 = Enum.GetNames(typeof(NoTicketsReng6)).ToList()
                                            .Select(r => "{" + r + "}").ToList();

            switch (ii)
            {
                case 0:
                    return TransaccionesReng1;
                case 1:
                    return TransaccionesReng2;
                case 2:
                    return TransaccionesReng3;
                case 3:
                    return TransaccionesReng4;
                case 4:
                    return TransaccionesReng5;
                case 5:
                    return TransaccionesReng6;
                case 6:
                    return TransaccionesReng6;
            }

            return new List<string>();
        }

        private List<string> DarTransaccRengAcum(int ii)
        {
            List<string> TransaccionesReng1 = Enum.GetNames(typeof(TransaccionesPorTurnoYCajaRengAcum1)).ToList()
                                            .Select(r => "{" + r + "}").ToList();

            List<string> TransaccionesReng2 = Enum.GetNames(typeof(TransaccionesPorTurnoYCajaRengAcum2)).ToList()
                                            .Select(r => "{" + r + "}").ToList();
            List<string> TransaccionesReng3 = Enum.GetNames(typeof(TransaccionesPorTurnoYCajaRengAcum3)).ToList()
                                            .Select(r => "{" + r + "}").ToList();
            List<string> TransaccionesReng4 = Enum.GetNames(typeof(TransaccionesPorTurnoYCajaRengAcum4)).ToList()
                                            .Select(r => "{" + r + "}").ToList();
            List<string> TransaccionesReng5 = Enum.GetNames(typeof(TransaccionesPorTurnoYCajaRengAcum5)).ToList()
                                            .Select(r => "{" + r + "}").ToList();

            List<string> TransaccionesReng6 = Enum.GetNames(typeof(TransaccionesPorTurnoYCajaRengAcum6)).ToList()
                                            .Select(r => "{" + r + "}").ToList();

            List<string> TransaccionesReng7 = Enum.GetNames(typeof(TransaccionesPorTurnoYCajaRengAcum7)).ToList()
                                            .Select(r => "{" + r + "}").ToList();

            switch (ii)
            {
                case 0:
                    return TransaccionesReng1;
                case 1:
                    return TransaccionesReng2;
                case 2:
                    return TransaccionesReng3;
                case 3:
                    return TransaccionesReng4;
                case 4:
                    return TransaccionesReng5;
                case 5:
                    return TransaccionesReng6;
                case 6:
                    return TransaccionesReng7;
            }

            return new List<string>();
        }

        private void CargarDatosPedidosiNoHayConexServ()
        {
            LstCuerpoMailSucursalPed = new List<CuerpoMailSucursalPedidos>() {
                new CuerpoMailSucursalPedidos(){
                    SUCURSAL = "No hay conexión" ,
                    NUM_SERVICIOS = 0,
                    IMPORTE = 0
                },
                new CuerpoMailSucursalPedidos(){
                    SUCURSAL = "No hay conexión" ,
                    NUM_SERVICIOS = 0,
                    IMPORTE = 0
                },
                new CuerpoMailSucursalPedidos(){
                    SUCURSAL = "No hay conexión" ,
                    NUM_SERVICIOS = 0,
                    IMPORTE = 0
                },
                new CuerpoMailSucursalPedidos(){
                    SUCURSAL = "No hay conexión" ,
                    NUM_SERVICIOS = 0,
                    IMPORTE = 0
                },
                new CuerpoMailSucursalPedidos(){
                    SUCURSAL = "No hay conexión" ,
                    NUM_SERVICIOS = 0,
                    IMPORTE = 0
                },
            };

            LstCuerpoMailSucursalPedAcum = new List<CuerpoMailSucursalPedidos>() {
                new CuerpoMailSucursalPedidos(){
                    SUCURSAL = "No hay conexión" ,
                    NUM_SERVICIOS = 0,                     
                    IMPORTE = 0
                },
                new CuerpoMailSucursalPedidos(){
                    SUCURSAL = "No hay conexión" ,
                    NUM_SERVICIOS = 0,
                    IMPORTE = 0
                },
                new CuerpoMailSucursalPedidos(){
                    SUCURSAL = "No hay conexión" ,
                    NUM_SERVICIOS = 0,
                    IMPORTE = 0
                },
                new CuerpoMailSucursalPedidos(){
                    SUCURSAL = "No hay conexión" ,
                    NUM_SERVICIOS = 0,
                    IMPORTE = 0
                },
                new CuerpoMailSucursalPedidos(){
                    SUCURSAL = "No hay conexión" ,
                    NUM_SERVICIOS = 0,
                    IMPORTE = 0
                },
            };
        }

        private void MostrarDatosSucursalesPed(bool Acum)
        {
            int ii = 1;
            int Servicios = 0;
            decimal Importe = 0;

            if (Acum? !ResultadoCargaAcumPedidosMicrosip : !ResultadoCargaPedidosMicrosip)
                CargarDatosPedidosiNoHayConexServ();

            foreach (CuerpoMailSucursalPedidos cuerpo in
                (Acum? LstCuerpoMailSucursalPedAcum : LstCuerpoMailSucursalPed).ToList())
            {
                TextPlantilla = TextPlantilla.Replace("{SPD" + ii + "}", 
                    cuerpo.SUCURSAL.Replace("POINT ", ""));

                TextPlantilla = TextPlantilla.Replace("{SERVPD" + (Acum? "A" : "") +  ii + "}", 
                    cuerpo.NUM_SERVICIOS.ToString());

                TextPlantilla = TextPlantilla.Replace("{IMPPD" + (Acum? "A" : "") + ii++ + "}",
                    cuerpo.IMPORTE.ToString("C2"));

                Servicios += (int)cuerpo.NUM_SERVICIOS;
                Importe += cuerpo.IMPORTE;
            }

            for(ii = 1; ii <= Constantes.CantSucursales; ii++)
            {
                TextPlantilla = TextPlantilla.Replace("{SPD" + ii + "}", "");
                TextPlantilla = TextPlantilla.Replace("{SERVPD" + (Acum ? "A" : "") + ii + "}", "");
                TextPlantilla = TextPlantilla.Replace("{IMPPD" + (Acum ? "A" : "") + ii + "}", "");
            }

            TextPlantilla = TextPlantilla.
                        Replace(Acum? "{TSERVPDA}" : "{TSERVPD}", Servicios.ToString()).
                        Replace(Acum? "{TIMPPDA}" : "{TIMPPD}", Importe.ToString("C2"));
        }

        private void SustInfoCuandoNoHayConexConAlgunServidor()
        {
            LstCuerpoMailSucursal = new List<CuerpoMailSucursal>() {
                new CuerpoMailSucursal() {
                    Sucursal = "Sin Conexión", 
                    Turno1Caja1 = -1,
                    Turno1Caja2 = -1,
                    Turno2Caja1 = -1,
                    Turno2Caja2 = -1,
                    Turno3Caja1 = -1,
                    Turno3Caja2 = -1,
                    TicketPromedio = 0,
                    ImporteSucursal = 0
                },
                new CuerpoMailSucursal() {
                    Sucursal = "Sin Conexión",
                    Turno1Caja1 = -1,
                    Turno1Caja2 = -1,
                    Turno2Caja1 = -1,
                    Turno2Caja2 = -1,
                    Turno3Caja1 = -1,
                    Turno3Caja2 = -1,
                    TicketPromedio = 0,
                    ImporteSucursal = 0
                },
                new CuerpoMailSucursal() {
                    Sucursal = "Sin Conexión",
                    Turno1Caja1 = -1,
                    Turno1Caja2 = -1,
                    Turno2Caja1 = -1,
                    Turno2Caja2 = -1,
                    Turno3Caja1 = -1,
                    Turno3Caja2 = -1,
                    TicketPromedio = 0,
                    ImporteSucursal = 0
                },
                new CuerpoMailSucursal() {
                    Sucursal = "Sin Conexión",
                    Turno1Caja1 = -1,
                    Turno1Caja2 = -1,
                    Turno2Caja1 = -1,
                    Turno2Caja2 = -1,
                    Turno3Caja1 = -1,
                    Turno3Caja2 = -1,
                    TicketPromedio = 0,
                    ImporteSucursal = 0
                },
                new CuerpoMailSucursal() {
                    Sucursal = "Sin Conexión",
                    Turno1Caja1 = -1,
                    Turno1Caja2 = -1,
                    Turno2Caja1 = -1,
                    Turno2Caja2 = -1,
                    Turno3Caja1 = -1,
                    Turno3Caja2 = -1,
                    TicketPromedio = 0,
                    ImporteSucursal = 0
                }
            };
        }

        private void SustInfoAcumCuandoNoHayConexConAlgunServidor()
        {
            LstCuerpoMailSucursalAcum = new List<CuerpoMailSucursal>() {
                new CuerpoMailSucursal() {
                    Sucursal = "Sin Conexión",
                    Turno1Caja1 = -1,
                    Turno1Caja2 = -1,
                    Turno2Caja1 = -1,
                    Turno2Caja2 = -1,
                    Turno3Caja1 = -1,
                    Turno3Caja2 = -1,
                    TicketPromedio = 0,
                    ImporteSucursal = 0
                },
                new CuerpoMailSucursal() {
                    Sucursal = "Sin Conexión",
                    Turno1Caja1 = -1,
                    Turno1Caja2 = -1,
                    Turno2Caja1 = -1,
                    Turno2Caja2 = -1,
                    Turno3Caja1 = -1,
                    Turno3Caja2 = -1,
                    TicketPromedio = 0,
                    ImporteSucursal = 0
                },
                new CuerpoMailSucursal() {
                    Sucursal = "Sin Conexión",
                    Turno1Caja1 = -1,
                    Turno1Caja2 = -1,
                    Turno2Caja1 = -1,
                    Turno2Caja2 = -1,
                    Turno3Caja1 = -1,
                    Turno3Caja2 = -1,
                    TicketPromedio = 0,
                    ImporteSucursal = 0
                },
                new CuerpoMailSucursal() {
                    Sucursal = "Sin Conexión",
                    Turno1Caja1 = -1,
                    Turno1Caja2 = -1,
                    Turno2Caja1 = -1,
                    Turno2Caja2 = -1,
                    Turno3Caja1 = -1,
                    Turno3Caja2 = -1,
                    TicketPromedio = 0,
                    ImporteSucursal = 0
                },
                new CuerpoMailSucursal() {
                    Sucursal = "Sin Conexión",
                    Turno1Caja1 = -1,
                    Turno1Caja2 = -1,
                    Turno2Caja1 = -1,
                    Turno2Caja2 = -1,
                    Turno3Caja1 = -1,
                    Turno3Caja2 = -1,
                    TicketPromedio = 0,
                    ImporteSucursal = 0
                }
            };
        }

        private void ReemplazarVariablesEnCuerpoMailSucursal()
        {
            decimal SumaTicketPromeio = 0;
            List<string> Sucursales = new EstacionesYVariablesHTML().GenerarVariableHTML("S", Constantes.EstacOSucur.Sucursal);

            List<string> Totales = 
                Enum.GetNames(typeof(TransaccionesTotales)).ToList().Select(r => "{" + r + "}").ToList();

            List<string> ImportesSucursales = 
                Enum.GetNames(typeof(ImportesSucursales)).ToList().Select(r => "{" + r + "}").ToList();

            List<string> TotalTransaccionesTurnoCaja = Enum.GetNames(typeof(TransaccionesTurnoYCajaTotales)).ToList().
                        Select(r => "{" + r + "}").ToList();

            List<string> TotalTransaccionesPromedio = Enum.GetNames(typeof(TransaccionesPromedio)).ToList().
                        Select(r => "{" + r + "}").ToList();

            int ii = 0, jj = 0;
            long TotalTrans = 0;
            List<long> LstTurnoCaja = new List<long>() { 0, 0, 0, 0, 0, 0 };
            List<decimal> LstTotTicketsEImporte = new List<decimal>() { 0, 0, 0, 0, 0 };
            decimal TotalImporteSucursal = 0, 
                TotalTicketProm = 0, TotalVtaPromPorTienda = 0;

            if (!ResultadoCargaMicrosip)
                SustInfoCuandoNoHayConexConAlgunServidor();

            foreach (CuerpoMailSucursal cuerpo in LstCuerpoMailSucursal)
            {
                TextPlantilla = TextPlantilla.Replace(Sucursales[ii], cuerpo.Sucursal.Replace("POINT ", ""));

                ReemplazarRenglonesSucursales(ref TextPlantilla,
                    DarTransaccReng(ii), cuerpo, ref LstTurnoCaja);

                TextPlantilla = TextPlantilla.
                    Replace(Totales[ii], cuerpo.TotalTrans.ToString("###,##0"));
                TextPlantilla = TextPlantilla.
                    Replace(ImportesSucursales[ii], cuerpo.ImporteSucursal.ToString("C2"));
                TextPlantilla = TextPlantilla.
                    Replace(TotalTransaccionesPromedio[ii++], cuerpo.TicketPromedio.
                        ToString("C2"));

                TotalTrans += cuerpo.TotalTrans;
                SumaTicketPromeio += cuerpo.TicketPromedio;
                TotalImporteSucursal += cuerpo.ImporteSucursal;                
            }

            TotalVtaPromPorTienda = ii == 0? (decimal)0 : (decimal) TotalImporteSucursal / (decimal)ii;
            TotalTicketProm = ii == 0? (decimal)0 : (decimal)SumaTicketPromeio / (decimal)ii;

            MostrarDatosSucursalesPed(false);

            for (ii = 0; ii < Constantes.CantSucursales; ii++)
            {
                TextPlantilla = TextPlantilla.Replace(Sucursales[ii], "");

                ReemplazarRenglonesSucursales(ref TextPlantilla,
                    DarTransaccReng(ii), new CuerpoMailSucursal {
                        Turno1Caja1 = -1, Turno2Caja1 = -1, Turno1Caja2 = -1, Turno2Caja2 = -1,
                        Turno3Caja1 = -1, Turno3Caja2 = -1
                    }, ref LstTurnoCaja);

                TextPlantilla = TextPlantilla.Replace(Totales[ii], "");
                TextPlantilla = TextPlantilla.Replace(ImportesSucursales[ii], "");
                TextPlantilla = TextPlantilla.Replace(TotalTransaccionesPromedio[ii], "");
            }

            for (jj = 0; jj < TotalTransaccionesTurnoCaja.Count; jj++)
                TextPlantilla = TextPlantilla.Replace(TotalTransaccionesTurnoCaja[jj], 
                    LstTurnoCaja[jj].ToString());

            TextPlantilla = TextPlantilla.Replace("{Fecha}", FechaDia);
            TextPlantilla = TextPlantilla.Replace("{TT}", TotalTrans.ToString("###,##0"));
            TextPlantilla = TextPlantilla.Replace("{TTP}", SumaTicketPromeio.ToString("C2"));
            TextPlantilla = TextPlantilla.Replace("{TIS}", TotalImporteSucursal.ToString("C2"));
            TextPlantilla = TextPlantilla.Replace("{TISP}", TotalVtaPromPorTienda.ToString("C2"));
        }

        private void ReemplazarVariablesAcumuladasEnCuerpoMailSucursal()
        {
            List<string> Sucursales = new EstacionesYVariablesHTML().GenerarVariableHTML("S", Constantes.EstacOSucur.Sucursal);

            List<string> Totales = Enum.GetNames(typeof(TransaccionesTotalesAcumuladas)).ToList().Select(r => "{" + r + "}").ToList();

            List<string> ImportesSucursales =
                Enum.GetNames(typeof(ImportesSucursalesAcumuladas)).ToList().Select(r => "{" + r + "}").ToList();

            List<string> TotalTransaccionesTurnoCaja = Enum.GetNames(typeof(TransaccionesTurnoYCajaTotalesAcumuladas)).ToList().
                        Select(r => "{" + r + "}").ToList();

            List<string> TotalTransaccionesPromedio = Enum.GetNames(typeof(TransaccionesPromedioAcumuladas)).ToList().
                        Select(r => "{" + r + "}").ToList();

            int ii = 0, jj = 0;
            long TotalTrans = 0;
            decimal SumaTicketProm = 0;

            List<long> LstTurnoCaja = new List<long>() { 0, 0, 0, 0, 0, 0 };
            decimal TotalImporteSucursal = 0, TotalTransaccionesProm = 0,
                TotalVtaPromPorTienda = 0;

            if (!ResultadoCargaAcumMicrosip)
                SustInfoAcumCuandoNoHayConexConAlgunServidor();

            foreach (CuerpoMailSucursal cuerpo in LstCuerpoMailSucursalAcum)
            {
                ReemplazarRenglonesSucursales(ref TextPlantilla,
                    DarTransaccRengAcum(ii), cuerpo, ref LstTurnoCaja);

                TextPlantilla = TextPlantilla.Replace(Totales[ii], 
                                cuerpo.TotalTrans.ToString("###,##0"));
                TextPlantilla = TextPlantilla.Replace(ImportesSucursales[ii], 
                                cuerpo.ImporteSucursal.ToString("C2"));
                TextPlantilla = TextPlantilla.
                                Replace(TotalTransaccionesPromedio[ii++], cuerpo.TicketPromedio.
                                    ToString("C2"));

                TotalTrans += cuerpo.TotalTrans;
                SumaTicketProm += cuerpo.TicketPromedio;
                TotalImporteSucursal += cuerpo.ImporteSucursal;
            }

            TotalVtaPromPorTienda = ii == 0? 0 : (decimal)TotalImporteSucursal / (decimal)ii;
            TotalTransaccionesProm = ii == 0? 0 : (decimal)SumaTicketProm / (decimal)ii;

            MostrarDatosSucursalesPed(true);

            for (jj = 0; jj < TotalTransaccionesTurnoCaja.Count; jj++)
                TextPlantilla = TextPlantilla.Replace(TotalTransaccionesTurnoCaja[jj],
                    LstTurnoCaja[jj].ToString("###,##0"));

            ii = 0;
            for (; ii < Constantes.CantSucursales; ii++)
            {
                TextPlantilla = TextPlantilla.Replace(Sucursales[ii], "");

                ReemplazarRenglonesSucursales(ref TextPlantilla,
                    DarTransaccRengAcum(ii), new CuerpoMailSucursal
                    {
                        Turno1Caja1 = -1, Turno2Caja1 = -1, Turno1Caja2 = -1, Turno2Caja2 = -1,
                        Turno3Caja1 = -1, Turno3Caja2 = -1
                    }, ref LstTurnoCaja);

                TextPlantilla = TextPlantilla.Replace(Totales[ii], "");
                TextPlantilla = TextPlantilla.Replace(ImportesSucursales[ii], "");
                TextPlantilla = TextPlantilla.Replace(TotalTransaccionesPromedio[ii], "");
            }

            TextPlantilla = TextPlantilla.Replace("{FechaA}", FechaMensAcum);
            TextPlantilla = TextPlantilla.Replace("{TTA}", TotalTrans.ToString("###,##0"));
            TextPlantilla = TextPlantilla.Replace("{TTPA}", SumaTicketProm.ToString("C2"));
            TextPlantilla = TextPlantilla.Replace("{TISA}", TotalImporteSucursal.ToString("C2"));
            TextPlantilla = TextPlantilla.Replace("{TISPA}", TotalVtaPromPorTienda.ToString("C2"));
        }

        private TimeSpan EntregarTurnoPorCaja(int ii, CuerpoMailSucursalTurno cuerpo)
        {
            TimeSpan respuesta = TimeSpan.MinValue;

            switch (ii)
            {
                case 0:
                    respuesta = cuerpo.Turno1Caja1;
                    break;
                case 1:
                    respuesta = cuerpo.Turno1Caja2;
                    break;
                case 2:
                    respuesta = cuerpo.Turno2Caja1;
                    break;
                case 3:
                    respuesta = cuerpo.Turno2Caja2;
                    break;
                case 4:
                    respuesta = cuerpo.Turno3Caja1;
                    break;
                case 5:
                    respuesta = cuerpo.Turno3Caja2;
                    break;
            }

            return respuesta;
        }

        private void ReemplazarRenglonesTurnoCajaHora(ref string Plantilla, List<string> Renglon,
            CuerpoMailSucursalTurno cuerpo)
        {
            int jj = 0;

            foreach (string elem in Renglon)
            {
                TextPlantilla = TextPlantilla.Replace(elem, cuerpo.Turno1Caja1 == new TimeSpan(100, 0, 0)? "" :
                    DarFormatoDeTiempo(EntregarTurnoPorCaja(jj++, cuerpo)));
            }
        }

        private void ReemplazarVariablesTurnoCajHoraEnCuerpoMailSucursalTurno()
        {
            List<List<string>> TurnosCajaHoraReng = new List<List<string>>();

            TurnosCajaHoraReng.Add(Enum.GetNames(typeof(TurnosCajaHoraR1)).ToList()
                                            .Select(r => "{" + r + "}").ToList());

            TurnosCajaHoraReng.Add(Enum.GetNames(typeof(TurnosCajaHoraR2)).ToList()
                                            .Select(r => "{" + r + "}").ToList());

            TurnosCajaHoraReng.Add(Enum.GetNames(typeof(TurnosCajaHoraR3)).ToList()
                                            .Select(r => "{" + r + "}").ToList());

            TurnosCajaHoraReng.Add(Enum.GetNames(typeof(TurnosCajaHoraR4)).ToList()
                                            .Select(r => "{" + r + "}").ToList());

            TurnosCajaHoraReng.Add(Enum.GetNames(typeof(TurnosCajaHoraR5)).ToList()
                                            .Select(r => "{" + r + "}").ToList());

            TurnosCajaHoraReng.Add(Enum.GetNames(typeof(TurnosCajaHoraR6)).ToList()
                                            .Select(r => "{" + r + "}").ToList());

            TurnosCajaHoraReng.Add(Enum.GetNames(typeof(TurnosCajaHoraR7)).ToList()
                                            .Select(r => "{" + r + "}").ToList());

            int ii = 0;
            foreach (CuerpoMailSucursalTurno elem in LstCuerpoMailSucursalTurno)
            {
                TextPlantilla = TextPlantilla.Replace("{Sucursal}", elem.Sucursal);
                ReemplazarRenglonesTurnoCajaHora(ref TextPlantilla, TurnosCajaHoraReng[ii++], elem);
            }   

            for (; ii < Constantes.CantSucursales; ii++)
            {
                TextPlantilla = TextPlantilla.Replace("{Sucursal}", "");
                ReemplazarRenglonesTurnoCajaHora(ref TextPlantilla, TurnosCajaHoraReng[ii], 
                    new CuerpoMailSucursalTurno { Turno1Caja1 = new TimeSpan(100, 0, 0) });
            }
        }

        private SmtpClient CuerpoCorreo(int Puerto)
        {
            SmtpClient smtp = null;

            try
            {
                smtp = new SmtpClient();
                //smtp.Port = 587;
                smtp.Port = Puerto;
                smtp.Host = new Encriptar().Decrypt("dEm6EKl0OL8e5I9MCf4rgQ=="); //IONOS
                //smtp.Host = new Encriptar().Decrypt("e3d+RfSBsCXpjDqkBopQhJMy9HzNZQUtTJDKmXQQtKY="); //hotmail
                smtp.EnableSsl = true;
                smtp.UseDefaultCredentials = false;
                smtp.Credentials = new NetworkCredential(
                    //new Encriptar().Decrypt("C/TiT1WeZH61hx/QI9hVFb428iE8pVYqgAPxY+9IkMA="),
                    //new Encriptar().Decrypt("xXf4kwXblk0Igu97NR7Cxg=="));

                //new Encriptar().Decrypt("JbOnWg/pqkyrUPVp5uRahVc3UH0Dq+hy28T5/nUD208="), //REYES CAMPOS
                //    new Encriptar().Decrypt("uUSRdF0cnqciZD91VUTR9A=="));

                new Encriptar().Decrypt("uOzNxS19nr5pZ3/xle9LQo2AybS9C5xQAe8mY7gsPJc="), //jose.campos@faza.com.mx
                    //new Encriptar().Decrypt("8+w5UkC+0F0NqAT/mCd6BVmvSpMVG8MI504+ofAv2aQ="));
                    new Encriptar().Decrypt("S6dhLAr1K2TcmhPfd41aRg=="));
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(ex.InnerException.Message);
            }

            return smtp;
        }

        private MailMessage Mensaje(string htmlString, string [] To, string [] Bcc)
        {
            MailMessage message = new MailMessage();

            message.From = new MailAddress("sistema.automatico@faza.com.mx");
            //message.From = new MailAddress("reyes_campos@hotmail.com");

            foreach (string elem in To)
                message.To.Add(new MailAddress(new Encriptar().Decrypt(elem)));

            foreach(string elem in Bcc)
                message.Bcc.Add(new MailAddress(new Encriptar().Decrypt(elem)));

            message.IsBodyHtml = true;
            message.Subject = "Ventas " + DateTime.Now.AddDays(Constantes.DiasAnteriores)
                .ToString("dddd, dd MMMM yyyy").Replace(".", "");

            LinkedResource[] recursosImgs = 
                new LinkedResource[Constantes.PathImagenes.ImagenesAIncluir.Count()];

            AlternateView av1 = AlternateView.
                CreateAlternateViewFromString(htmlString, null, "text/html");

            for (int ii = 0; ii < Constantes.PathImagenes.ImagenesAIncluir.Count(); ii++)
            {
                recursosImgs[ii] = 
                    new LinkedResource(Constantes.PathImagenes.ImagenesAIncluir[ii]);
                recursosImgs[ii].ContentId = "Imagen" + (ii + 1);
                av1.LinkedResources.Add(recursosImgs[ii]);
            }

            message.AlternateViews.Add(av1);
            message.IsBodyHtml = true;

            return message;
        }

        private void FuncionesComunesEnEnviarSMTP(string PathPlantilla, string PathPlantillaRengPedCNT,
            string PathPlantillaRengPedMTR, string PathPlantillaRengPedSTF)
        {
            TextPlantilla = System.IO.File.ReadAllText(PathPlantilla);
            TextPlantillaRengPedCNT = System.IO.File.ReadAllText(PathPlantillaRengPedCNT);
            TextPlantillaRengPedMTR = System.IO.File.ReadAllText(PathPlantillaRengPedMTR);
            TextPlantillaRengPedSTF = System.IO.File.ReadAllText(PathPlantillaRengPedSTF);

            ReemplazarVariablesEnCuerpoMailEstacion();
            ReemplazarVariablesEnCuerpoMailEstacionDinero();
            ReemplazarVariablesAcumuladasEnCuerpoMailEstacion();
            ReemplazarVariablesAcumuladasEnCuerpoMailEstacionDinero();

            ReemplazarVariablesEnCuerpoMailSucursal();
            ReemplazarVariablesAcumuladasEnCuerpoMailSucursal();

            ReemplazarVariablesTurnoCajHoraEnCuerpoMailSucursalTurno();
        }

        private void EnviarSMTP()
        {
            //SmtpClient smtp = CuerpoCorreo(587);
            SmtpClient smtp = CuerpoCorreo(587);

            if (Constantes.CorreosTodaLaInformacion.CorreosConTodaLaInformacionTo.Count() > 0)
            {
                FuncionesComunesEnEnviarSMTP(Constantes.CorreosTodaLaInformacion.PathPlantillaCorreo,
                    Constantes.CorreosTodaLaInformacion.PathPlantillaRenglPedCNT,
                    Constantes.CorreosTodaLaInformacion.PathPlantillaRenglPedMTR,
                    Constantes.CorreosTodaLaInformacion.PathPlantillaRenglPedSTF);

                smtp.Send(Mensaje(TextPlantilla,
                    Constantes.CorreosTodaLaInformacion.CorreosConTodaLaInformacionTo,
                    Constantes.CorreosTodaLaInformacion.CorreosConTodaLaInformacionBcc));
            }

            smtp = CuerpoCorreo(587);

            if (Constantes.CorreosSoloSucursales.CorreosConSoloSucursalesTo.Count() > 0)
            {
                FuncionesComunesEnEnviarSMTP(Constantes.CorreosSoloSucursales.PathPlantillaCorreo,
                    Constantes.CorreosTodaLaInformacion.PathPlantillaRenglPedCNT,
                    Constantes.CorreosTodaLaInformacion.PathPlantillaRenglPedMTR,
                    Constantes.CorreosTodaLaInformacion.PathPlantillaRenglPedSTF);

                smtp.Send(Mensaje(TextPlantilla,
                    Constantes.CorreosSoloSucursales.CorreosConSoloSucursalesTo,
                    Constantes.CorreosSoloSucursales.CorreosConSoloSucursalesBcc));
            }

            smtp = CuerpoCorreo(587);

            if (Constantes.CorreosSoloEstaciones.CorreosConSoloEstacionesTo.Count() > 0)
            {
                FuncionesComunesEnEnviarSMTP(Constantes.CorreosSoloEstaciones.PathPlantillaCorreoGasolinera,
                    Constantes.CorreosTodaLaInformacion.PathPlantillaRenglPedCNT,
                    Constantes.CorreosTodaLaInformacion.PathPlantillaRenglPedMTR,
                    Constantes.CorreosTodaLaInformacion.PathPlantillaRenglPedSTF);

                smtp.Send(Mensaje(TextPlantilla,
                    Constantes.CorreosSoloEstaciones.CorreosConSoloEstacionesTo,
                    Constantes.CorreosSoloEstaciones.CorreosConSoloEstacionesBcc));
            }

            Application.Exit();
        }
    }
}
