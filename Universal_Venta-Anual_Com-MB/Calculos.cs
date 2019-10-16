using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Universal_Venta_Anual_Com_MB
{
    class Calculos
    {
        public static double SumaVentasMeses(SqlDataReader readerS)
        {
            return Convert.ToDouble(readerS["Enero"].ToString()) +
                                 Convert.ToDouble(readerS["Febrero"].ToString()) +
                                 Convert.ToDouble(readerS["Marzo"].ToString()) +
                                 Convert.ToDouble(readerS["Abril"].ToString()) +
                                 Convert.ToDouble(readerS["Mayo"].ToString()) +
                                 Convert.ToDouble(readerS["Junio"].ToString()) +
                                 Convert.ToDouble(readerS["Julio"].ToString()) +
                                 Convert.ToDouble(readerS["Agosto"].ToString()) +
                                 Convert.ToDouble(readerS["Septiembre"].ToString()) +
                                 Convert.ToDouble(readerS["Octubre"].ToString()) +
                                 Convert.ToDouble(readerS["Noviembre"].ToString()) +
                                 Convert.ToDouble(readerS["Diciembre"].ToString());
        }
        public static void SumasVentas(ref double sumaVentas,ref int mesesVentas, SqlDataReader readerS)
        {
            if (Convert.ToDouble(readerS["Enero"].ToString()) > 0)
            {
                sumaVentas = sumaVentas + Convert.ToDouble(readerS["Enero"].ToString());
                mesesVentas++;

            }
            if (Convert.ToDouble(readerS["Febrero"].ToString()) > 0)
            {
                sumaVentas = sumaVentas + Convert.ToDouble(readerS["Febrero"].ToString());
                mesesVentas++;
            }
            if (Convert.ToDouble(readerS["Marzo"].ToString()) > 0)
            {
                sumaVentas = sumaVentas + Convert.ToDouble(readerS["Marzo"].ToString());
                mesesVentas++;
            }
            if (Convert.ToDouble(readerS["Abril"].ToString()) > 0)
            {
                sumaVentas = sumaVentas + Convert.ToDouble(readerS["Abril"].ToString());
                mesesVentas++;
            }
            if (Convert.ToDouble(readerS["Mayo"].ToString()) > 0)
            {
                sumaVentas = sumaVentas + Convert.ToDouble(readerS["Mayo"].ToString());
                mesesVentas++;
            }
            if (Convert.ToDouble(readerS["Junio"].ToString()) > 0)
            {
                sumaVentas = sumaVentas + Convert.ToDouble(readerS["Junio"].ToString());
                mesesVentas++;
            }
            if (Convert.ToDouble(readerS["Julio"].ToString()) > 0)
            {
                sumaVentas = sumaVentas + Convert.ToDouble(readerS["Julio"].ToString());
                mesesVentas++;
            }
            if (Convert.ToDouble(readerS["Agosto"].ToString()) > 0)
            {
                sumaVentas = sumaVentas + Convert.ToDouble(readerS["Agosto"].ToString());
                mesesVentas++;
            }
            if (Convert.ToDouble(readerS["Septiembre"].ToString()) > 0)
            {
                sumaVentas = sumaVentas + Convert.ToDouble(readerS["Septiembre"].ToString());
                mesesVentas++;
            }
            if (Convert.ToDouble(readerS["Octubre"].ToString()) > 0)
            {
                sumaVentas = sumaVentas + Convert.ToDouble(readerS["Octubre"].ToString());
                mesesVentas++;
            }
            if (Convert.ToDouble(readerS["Noviembre"].ToString()) > 0)
            {
                sumaVentas = sumaVentas + Convert.ToDouble(readerS["Noviembre"].ToString());
                mesesVentas++;
            }
            if (Convert.ToDouble(readerS["Diciembre"].ToString()) > 0)
            {
                sumaVentas = sumaVentas + Convert.ToDouble(readerS["Diciembre"].ToString());
                mesesVentas++;
            }

        }
        private static  void marcarExistenciaAlterna(double existenciaAlternaE, int ii, int celda, DataGridView dataGridView1)
        {
            if (existenciaAlternaE > 0)
            {
                //dataGridView1.Rows[i].Cells[1+consecutivoAlterno].Style.BackColor = Color.FromArgb(125, 206, 160);
                dataGridView1.Rows[ii].Cells[celda].Style.BackColor = Color.FromArgb(214, 219, 223);
            }
        }

        public static  void ValidarClasificaciones(string clasificacion, string clasifProdAlterno, string productoAlterno, string anio, string mes, string dia, DataGridView dataGridView1,int i,int consecutivoAltgerno, int consecutivoAlterno)
        {
            double existenciaAlterna = 0;
            int sv = 0;
            if (clasificacion == "KWP")
            {
                if (clasifProdAlterno == "CI")
                {
                    dataGridView1.Rows[i].Cells[7].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[27].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 7,dataGridView1);
                }

                if (clasifProdAlterno == "SV")
                {
                    dataGridView1.Rows[i].Cells[8].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[28].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 8, dataGridView1);
                }


                if (clasifProdAlterno == "CA")
                {
                    dataGridView1.Rows[i].Cells[9].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[29].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 9, dataGridView1);
                }

                if (clasifProdAlterno == "AU")
                {
                    dataGridView1.Rows[i].Cells[10].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[30].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 10, dataGridView1);
                }


                if (clasifProdAlterno == "AM")
                {
                    dataGridView1.Rows[i].Cells[11].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[31].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 11, dataGridView1);
                }

            }//FIN KWP
            else if (clasificacion == "AUTOMOTIVE")
            {
                if (clasifProdAlterno == "CI")
                {
                    dataGridView1.Rows[i].Cells[7].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[27].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 7, dataGridView1);
                }

                if (clasifProdAlterno == "SV")
                {
                    dataGridView1.Rows[i].Cells[8].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[28].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 8, dataGridView1);
                }


                if (clasifProdAlterno == "CA")
                {
                    dataGridView1.Rows[i].Cells[9].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[29].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 9, dataGridView1);
                }

                if (clasifProdAlterno == "AM")
                {
                    dataGridView1.Rows[i].Cells[10].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[30].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 10, dataGridView1);
                }


                if (clasifProdAlterno == "VA")
                {
                    dataGridView1.Rows[i].Cells[11].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[31].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 11, dataGridView1);
                }
                if (clasifProdAlterno == "WA")
                {
                    dataGridView1.Rows[i].Cells[12].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[32].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 12, dataGridView1);
                }

            }//FIN DE AUTOMOTIVE

            else if (clasificacion == "CALDERON")
            {
                if (clasifProdAlterno == "CI")
                {
                    dataGridView1.Rows[i].Cells[7].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[27].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 7, dataGridView1);
                }

                if (clasifProdAlterno == "SV")
                {
                    dataGridView1.Rows[i].Cells[8].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[28].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 8, dataGridView1);
                }


                if (clasifProdAlterno == "FR")
                {
                    dataGridView1.Rows[i].Cells[9].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[29].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 9, dataGridView1);
                }

                if (clasifProdAlterno == "WA")
                {
                    dataGridView1.Rows[i].Cells[10].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[30].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 10, dataGridView1);
                }


                if (clasifProdAlterno == "VA")
                {
                    dataGridView1.Rows[i].Cells[11].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[31].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 11, dataGridView1);
                }
                if (clasifProdAlterno == "DA")
                {
                    dataGridView1.Rows[i].Cells[12].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[32].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 12, dataGridView1);
                }

            }//FIN DE CALDERON

            else if (clasificacion == "CIOSA")
            {
                if (clasifProdAlterno == "SV")
                {
                    dataGridView1.Rows[i].Cells[7].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[27].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 7, dataGridView1);
                }

                if (clasifProdAlterno == "CA")
                {
                    dataGridView1.Rows[i].Cells[8].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[28].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 8, dataGridView1);
                }


                if (clasifProdAlterno == "VA")
                {
                    dataGridView1.Rows[i].Cells[9].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[29].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 9, dataGridView1);
                }

                if (clasifProdAlterno == "WA")
                {
                    dataGridView1.Rows[i].Cells[10].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[30].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 10, dataGridView1);
                }


                if (clasifProdAlterno == "A1" || clasifProdAlterno == "A2" || clasifProdAlterno == "A3" || clasifProdAlterno == "A4" || clasifProdAlterno == "A5" || clasifProdAlterno == "A6")
                {
                    dataGridView1.Rows[i].Cells[11].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[31].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 11, dataGridView1);
                }
                if (clasifProdAlterno == "CP")
                {
                    dataGridView1.Rows[i].Cells[12].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[32].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 12, dataGridView1);
                }

            }//FIN DE CIOSA

            else if (clasificacion == "SERVA")
            {
                if (clasifProdAlterno == "CI")
                {
                    dataGridView1.Rows[i].Cells[7].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[27].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 7, dataGridView1);
                }

                if (clasifProdAlterno == "CA")
                {
                    dataGridView1.Rows[i].Cells[8].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[28].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 8, dataGridView1);
                }


                if (clasifProdAlterno == "VA")
                {
                    dataGridView1.Rows[i].Cells[9].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[29].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 9, dataGridView1);
                }

                if (clasifProdAlterno == "DA")
                {
                    dataGridView1.Rows[i].Cells[10].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[30].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 10, dataGridView1);
                }


                if (clasifProdAlterno == "A1" || clasifProdAlterno == "A2" || clasifProdAlterno == "A3" || clasifProdAlterno == "A4" || clasifProdAlterno == "A5" || clasifProdAlterno == "A6")
                {
                    dataGridView1.Rows[i].Cells[11].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[31].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 11, dataGridView1);
                }
                if (clasifProdAlterno == "CP")
                {
                    dataGridView1.Rows[i].Cells[12].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[32].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 12, dataGridView1);
                }

            }//FIN DE SERVA

            else if (clasificacion == "APYMSA")
            {
                if (clasifProdAlterno == "A1" || clasifProdAlterno == "A2" || clasifProdAlterno == "A3" || clasifProdAlterno == "A4" || clasifProdAlterno == "A5" || clasifProdAlterno == "A6")
                {
                    dataGridView1.Rows[i].Cells[7].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[27].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 7, dataGridView1);
                }

                if (clasifProdAlterno == "VA")
                {
                    dataGridView1.Rows[i].Cells[8].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[28].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 8, dataGridView1);
                }


                if (clasifProdAlterno == "CI")
                {
                    dataGridView1.Rows[i].Cells[9].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[29].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 9, dataGridView1);
                }

                if (clasifProdAlterno == "SV")
                {
                    dataGridView1.Rows[i].Cells[10].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[30].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 10, dataGridView1);
                }


                if (clasifProdAlterno == "FR")
                {
                    dataGridView1.Rows[i].Cells[11].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[31].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 11, dataGridView1);
                }
                if (clasifProdAlterno == "DA")
                {
                    dataGridView1.Rows[i].Cells[12].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[32].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 12, dataGridView1);
                }

            }//FIN DE APYMSA

            else if (clasificacion == "CAP")
            {

                if (clasifProdAlterno == "CI")
                {
                    dataGridView1.Rows[i].Cells[7].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[27].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 7, dataGridView1);
                }

                if (clasifProdAlterno == "SV" && sv == 0)
                {
                    dataGridView1.Rows[i].Cells[8].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[28].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 8, dataGridView1);
                    sv++;
                }


                if (clasifProdAlterno == "SV" && sv == 1)
                {
                    dataGridView1.Rows[i].Cells[9].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[29].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 9, dataGridView1);
                    sv++;
                }

            }//FIN DE CAP

            else if (clasificacion == "DAP")
            {
                if (clasifProdAlterno == "GA")
                {
                    dataGridView1.Rows[i].Cells[7].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[27].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 7, dataGridView1);
                }

                if (clasifProdAlterno == "FR")
                {
                    dataGridView1.Rows[i].Cells[8].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[28].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 8, dataGridView1);
                }


                if (clasifProdAlterno == "SV")
                {
                    dataGridView1.Rows[i].Cells[9].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[29].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 9, dataGridView1);
                }

                if (clasifProdAlterno == "CA")
                {
                    dataGridView1.Rows[i].Cells[10].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[30].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 10, dataGridView1);
                }

            }//FIN DE DAP

            else if (clasificacion == "GAMASA")
            {
                if (clasifProdAlterno == "DA")
                {
                    dataGridView1.Rows[i].Cells[7].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[27].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 7, dataGridView1);
                }

                if (clasifProdAlterno == "FR")
                {
                    dataGridView1.Rows[i].Cells[8].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[28].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 8, dataGridView1);
                }

                if (clasifProdAlterno == "SV")
                {
                    dataGridView1.Rows[i].Cells[9].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[29].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 9, dataGridView1);
                }

                if (clasifProdAlterno == "CA")
                {
                    dataGridView1.Rows[i].Cells[10].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[30].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 10, dataGridView1);
                }

            }//FIN DE GAMASA

            else if (clasificacion == "WAI")
            {
                if (clasifProdAlterno == "CI")
                {
                    dataGridView1.Rows[i].Cells[7].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[27].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 7, dataGridView1);
                }

                if (clasifProdAlterno == "SV")
                {
                    dataGridView1.Rows[i].Cells[8].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[28].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 8, dataGridView1);
                }

                if (clasifProdAlterno == "FR")
                {
                    dataGridView1.Rows[i].Cells[9].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[29].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 9, dataGridView1);
                }

                if (clasifProdAlterno == "CA")
                {
                    dataGridView1.Rows[i].Cells[10].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[30].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 10, dataGridView1);
                }
                if (clasifProdAlterno == "AM")
                {
                    dataGridView1.Rows[i].Cells[11].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[31].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 11, dataGridView1);
                }

            }//FIN DE WAI

            else if (clasificacion == "COMERCIALIZADORA LUNA")
            {
                if (clasifProdAlterno == "SV")
                {
                    dataGridView1.Rows[i].Cells[7].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[27].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 7, dataGridView1);
                }

                if (clasifProdAlterno == "CI")
                {
                    dataGridView1.Rows[i].Cells[8].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[28].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 8, dataGridView1);
                }

                if (clasifProdAlterno == "AM")
                {
                    dataGridView1.Rows[i].Cells[9].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[29].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 9, dataGridView1);
                }

                if (clasifProdAlterno == "HO")
                {
                    dataGridView1.Rows[i].Cells[10].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[30].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 10, dataGridView1);
                }
                if (clasifProdAlterno == "A1" || clasifProdAlterno == "A2" || clasifProdAlterno == "A3" || clasifProdAlterno == "A4" || clasifProdAlterno == "A5" || clasifProdAlterno == "A6")
                {
                    dataGridView1.Rows[i].Cells[11].Value = productoAlterno;
                    SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                    dataGridView1.Rows[i].Cells[31].Value = existenciaAlterna;
                    marcarExistenciaAlterna(existenciaAlterna, i, 11, dataGridView1);
                }

            }//FIN DE COMERCIALIZADORA LUNA
            else
            {
                dataGridView1.Rows[i].Cells[7 + consecutivoAlterno].Value = productoAlterno;
                SDK.fRegresaExistencia(productoAlterno, "1", anio, mes, dia, ref existenciaAlterna);
                dataGridView1.Rows[i].Cells[27 + consecutivoAlterno].Value = existenciaAlterna;
                dataGridView1.Rows[i].Cells[27 + consecutivoAlterno].Value = existenciaAlterna;
                marcarExistenciaAlterna(existenciaAlterna, i, 7 + consecutivoAlterno,dataGridView1);
            }
        }
    }
}
