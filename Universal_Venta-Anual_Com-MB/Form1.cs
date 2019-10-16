using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Configuration;
using System.Data.SqlClient;
using OfficeOpenXml;
using System.IO;
using OfficeOpenXml.Style;
using MySql.Data.MySqlClient;
using Microsoft.Office.Interop;
using Excel= Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Threading;

namespace Universal_Venta_Anual_Com_MB
{
    public partial class frmPrincipal : Form
    {
        public static string conexionSql = ConfigurationManager.ConnectionStrings["connSQL"].ConnectionString;
        public static string conexionMysql = ConfigurationManager.ConnectionStrings["ConnMysql"].ConnectionString;
        //SqlConnection conexion = new SqlConnection(conexionSql);

        public static int lError;
        public static int bandera;
        string ruta;
        string fechaUltimoCosto;
        double ultimoCosto = 0.00;
        DateTime f = DateTime.Now;
        int countPro = 0;
        int total = 0;

        public frmPrincipal()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            bgw1.RunWorkerAsync();
            StringBuilder sMensaje = new StringBuilder(512);
            //Elimina la última fila vacía
            dataGridView1.AllowUserToAddRows = false;
           ///Aquí va el código de xonexión a SDK
            Conexiones.conexionSDK(ref bandera,lError, txtMsg);

            if (lError != 0)
            {
                SDK.rError(lError);
                return;
            }
            else
            {
                txtMsg.Text = "Se abrió la empresa correctamente";
                btnFiltrar.Enabled = true;
                btnExcel.Enabled = false;
                                
                CargarListaDeClasificaciones();

            }
        }

        private void CargarListaDeClasificaciones()
        {
            var dt = Datos.ObtenerListaDeClasificaciones();


            DataRow dr = dt.NewRow();
            dr["CVALORCLASIFICACION"] = "SELECCIONA UN PROVEEDOR";
            dt.Rows.InsertAt(dr, 0);

            comboCodigo.ValueMember = "CIDVALORCLASIFICACION";
            comboCodigo.DisplayMember = "CVALORCLASIFICACION";
            comboCodigo.DataSource = dt;
        }

       private void frmPrincipal_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (bandera == 1)
            {
                SDK.fCierraEmpresa();
                SDK.fTerminaSDK();
                //conexion.Close();
            }
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            //Application.DoEvents();
            btnExcel.Enabled = true;

            dataGridView1.Rows.Clear();
            String cadenaS = Conexiones.CadenaConexionVentasAnuales(comboCodigo.SelectedValue.ToString(),txtAnio.Text.Trim());
            using (SqlConnection conexionS = new SqlConnection(conexionSql))
            {
                total = 0;
                countPro = 0;
                DateTime fechaActual = DateTime.Today;
                int numMesActual = 0;
                numMesActual = fechaActual.Month-1;
                double sumaVentas = 0;
                int mesesVentas = 0;
                double sumaExistenciasPrincipal = 0;
                double sumaExistenciasAlterno = 0;
                Boolean tienealterno = false;
                conexionS.Open();
                progressBar1.Maximum = 100;
                progressBar1.Minimum = 0;
                progressBar1.Step = 1;
                int porciento = 0;
                double pEne=0,pFeb=0,pMar=0,pAbr=0,pMay=0,pJun=0,pJul=0,pAgo=0,pSep=0,pOct=0,pNov=0,pDic=0.00;
                double existencia;
                string producto;
                string dia, mes, anio;
                dia = DateTime.Today.Day.ToString();
                mes = DateTime.Today.Month.ToString();
                anio = DateTime.Today.Year.ToString();
                SqlCommand cmdCR = new SqlCommand(cadenaS, conexionS);
                int i = 0;
                SqlDataReader readerCR = cmdCR.ExecuteReader();
                while (readerCR.Read())
                {
                    //Application.DoEvents();//ok
                    total++;
                }
                SqlCommand cmdS = new SqlCommand(cadenaS, conexionS);
                SqlDataReader readerS = cmdS.ExecuteReader();
                while (readerS.Read())
                {
                    pEne = 0; pFeb = 0; pMar = 0; pAbr = 0; pMay = 0; pJun = 0; pJul = 0; pAgo = 0; pSep = 0; pOct = 0; pNov = 0; pDic = 0.00;

                    
                    tienealterno = false;
                    //dataGridView1.Rows.Add();
                    int renglon = dataGridView1.Rows.Add();
                    dataGridView1.Rows[i].Cells[0].Value = Convert.ToString(readerS["CNOMBREPRODUCTO"].ToString()).Trim();
                    dataGridView1.Rows[i].Cells[1].Value = Convert.ToString(readerS["CCODIGOPRODUCTO"].ToString()).Trim();
                    countPro++;
                    double suma = 0.00;
                    double sP = 0.00;
                    sumaExistenciasPrincipal = 0;
                    mesesVentas = 0;
                    sumaVentas = 0;
                    existencia = 0;
                    txtMsg.Text = "Registros insertados: " + countPro;
                    ProgresoCarga();

                    producto = Convert.ToString(readerS["CCODIGOPRODUCTO"].ToString());
                    SDK.fRegresaExistencia(producto.ToString(), "1", anio, mes, dia, ref existencia);

                    //OBTIENE LAS ENTRADAS DEL PRODUCTO PARA RESTARLAS DE LAS VENTAS
                    String cadenaSVP = Conexiones.CadenaConexionVentasAnualesPorProducto(producto.ToString(), txtAnio.Text.Trim());
                    using (SqlConnection conexionSVP = new SqlConnection(conexionSql))
                    {
                        conexionSVP.Open();
                        SqlCommand cmdSVP = new SqlCommand(cadenaSVP, conexionSVP);
                        //int i = 0;
                        SqlDataReader readerSVP = cmdSVP.ExecuteReader();
                        while (readerSVP.Read())
                        {
                            pEne= Convert.ToDouble(readerSVP["Enero"].ToString());
                            pFeb = Convert.ToDouble(readerSVP["Febrero"].ToString());
                            pMar = Convert.ToDouble(readerSVP["Marzo"].ToString());
                            pAbr = Convert.ToDouble(readerSVP["Abril"].ToString());
                            pMay = Convert.ToDouble(readerSVP["Mayo"].ToString());
                            pJun = Convert.ToDouble(readerSVP["Junio"].ToString());
                            pJul = Convert.ToDouble(readerSVP["Julio"].ToString());
                            pAgo = Convert.ToDouble(readerSVP["Agosto"].ToString());
                            pSep = Convert.ToDouble(readerSVP["Septiembre"].ToString());
                            pOct = Convert.ToDouble(readerSVP["Octubre"].ToString());
                            pNov = Convert.ToDouble(readerSVP["Noviembre"].ToString());
                            pDic = Convert.ToDouble(readerSVP["Diciembre"].ToString());
                        }
                    }

                    //FIN DE LAS ENTRADAS DEL PRODUCTO A RESTAR
                    sP = (pEne + pFeb + pMar + pAbr + pMay + pJun + pJul + pAgo + pSep + pOct + pNov + pDic);
                    //sP = 0;

                   suma = Calculos.SumaVentasMeses(readerS) - sP;//Suma las ventas del año

                    dataGridView1.Rows[i].Cells[3].Value = suma;
                    suma = suma + existencia;
                    sumaExistenciasPrincipal = sumaExistenciasPrincipal + existencia;
                    dataGridView1.Rows[i].Cells[2].Value = existencia;
                    dataGridView1.Rows[i].Cells[14].Value = Convert.ToDouble(readerS["Enero"].ToString())-pEne;
                    dataGridView1.Rows[i].Cells[15].Value = Convert.ToDouble(readerS["Febrero"].ToString())-pFeb;
                    dataGridView1.Rows[i].Cells[16].Value = Convert.ToDouble(readerS["Marzo"].ToString())-pMar;
                    dataGridView1.Rows[i].Cells[17].Value = Convert.ToDouble(readerS["Abril"].ToString())-pAbr;
                    dataGridView1.Rows[i].Cells[18].Value = Convert.ToDouble(readerS["Mayo"].ToString())-pMay;
                    dataGridView1.Rows[i].Cells[19].Value = Convert.ToDouble(readerS["Junio"].ToString())-pJun;
                    dataGridView1.Rows[i].Cells[20].Value = Convert.ToDouble(readerS["Julio"].ToString())-pJul;
                    dataGridView1.Rows[i].Cells[21].Value = Convert.ToDouble(readerS["Agosto"].ToString())-pAgo;
                    dataGridView1.Rows[i].Cells[22].Value = Convert.ToDouble(readerS["Septiembre"].ToString())-pSep;
                    dataGridView1.Rows[i].Cells[23].Value = Convert.ToDouble(readerS["Octubre"].ToString())-pOct;
                    dataGridView1.Rows[i].Cells[24].Value = Convert.ToDouble(readerS["Noviembre"].ToString())-pNov;
                    dataGridView1.Rows[i].Cells[25].Value = Convert.ToDouble(readerS["Diciembre"].ToString())-pDic;

                    Calculos.SumasVentas(ref sumaVentas,ref mesesVentas,readerS);//suma las ventas y meses en que se vendieron productos
                    
                    if (suma > 0)
                    {
                        int registrosCompras = 0;

                        using (SqlConnection conexionSCF = new SqlConnection(conexionSql))
                        {
                            conexionSCF.Open();
                            string conCostoFecha=Conexiones.ConnCostoFechaSQLServer(producto);
                            SqlCommand cmdSCF = new SqlCommand(conCostoFecha, conexionSCF);//Cadena de conexion para costo y fecha
                            SqlDataReader readerSCF = cmdSCF.ExecuteReader();
                            if (readerSCF != null)
                            {
                                while (readerSCF.Read())
                                {
                                    //Application.DoEvents();
                                    ultimoCosto = Convert.ToDouble(readerSCF["costo"].ToString());
                                    fechaUltimoCosto = Convert.ToString(readerSCF["cfecha"].ToString()).Substring(0, 10);
                                    registrosCompras++;
                                }
                            }
                            else
                            {
                                fechaUltimoCosto = "";
                                ultimoCosto = 0;
                            }
                            if (registrosCompras == 0)
                            {
                                fechaUltimoCosto = "";
                                ultimoCosto = 0.00;
                            }
                            dataGridView1.Rows[i].Cells[6].Value = Math.Round(ultimoCosto,2);
                            dataGridView1.Rows[i].Cells[26].Value = fechaUltimoCosto.Substring(0, 10);
                        }
                    } //fin del if (suma > 0)

                    //inicio para obtener datos de MB
                    using (MySqlConnection conexionM = new MySqlConnection(conexionMysql))
                    {
                        conexionM.Open();
                        int consecutivoAlterno = 0;
                        double existenciaAlterna = 0.00;
                        sumaExistenciasAlterno = 0;
                        string clasifProdAlterno = "";
                        string clasificacion = comboCodigo.Text.Trim();
                        string productoAlterno = "";
                        //int sv = 0;
                                                
                        string consultaM = @"select articulo,padre from asociados where padre= '" + producto + "'";
                        MySqlCommand cmdM = new MySqlCommand(consultaM, conexionM);
                        MySqlDataReader ReaderM = cmdM.ExecuteReader();
                        if (ReaderM != null)
                        {
                            tienealterno = false;
                        }
                        else
                        {
                            tienealterno = true;
                        }
                        dataGridView1.Rows[i].Cells[7].Value = "";dataGridView1.Rows[i].Cells[8].Value = "";
                        dataGridView1.Rows[i].Cells[9].Value = "";dataGridView1.Rows[i].Cells[10].Value = "";
                        dataGridView1.Rows[i].Cells[11].Value = "";dataGridView1.Rows[i].Cells[12].Value = "";
                        dataGridView1.Rows[i].Cells[13].Value = "";
                        while (ReaderM.Read())
                        {
                            //Application.DoEvents();
                            productoAlterno = "";                         
                            productoAlterno = Convert.ToString(ReaderM["articulo"].ToString()).Trim();
                            clasifProdAlterno = productoAlterno.Substring(0,2);
                            existenciaAlterna = 0;
                            //valida clasificaciones
                            Calculos.ValidarClasificaciones(clasificacion,clasifProdAlterno,productoAlterno,anio,mes,dia,dataGridView1,i,consecutivoAlterno,consecutivoAlterno);
                            sumaExistenciasAlterno = Convert.ToDouble(dataGridView1.Rows[i].Cells[27].Value)+ Convert.ToDouble(dataGridView1.Rows[i].Cells[28].Value)+ Convert.ToDouble(dataGridView1.Rows[i].Cells[29].Value)+ Convert.ToDouble(dataGridView1.Rows[i].Cells[30].Value)+ Convert.ToDouble(dataGridView1.Rows[i].Cells[31].Value)+ Convert.ToDouble(dataGridView1.Rows[i].Cells[32].Value)+ Convert.ToDouble(dataGridView1.Rows[i].Cells[33].Value);
                            
                            //inicio obtener último costo y fecha ultimo costo

                            using (SqlConnection conexionSCFA = new SqlConnection(conexionSql))
                            {
                                double ultimoCostoA = 0.00;
                                string fechaUltimoCostoA;
                                int registrosComprasA = 0;
                                conexionSCFA.Open();
                                string conCostoFecha2 = Conexiones.ConnCostoFechaSQLServer2(dataGridView1.Rows[i].Cells[7 + consecutivoAlterno].Value.ToString().Trim());//conexión para costos de alternos
                                SqlCommand cmdSCFA =new SqlCommand(conCostoFecha2, conexionSCFA);
                                
                                SqlDataReader readerSCFA = cmdSCFA.ExecuteReader();
                                if (readerSCFA != null)
                                {
                                    while (readerSCFA.Read())
                                    {
                                        //Application.DoEvents();
                                        ultimoCostoA = Convert.ToDouble(readerSCFA["costo"].ToString());
                                        fechaUltimoCostoA = Convert.ToString(readerSCFA["cfecha"].ToString()).Substring(0, 10);
                                        registrosComprasA++;
                                    }
                                }
                                else
                                {
                                    fechaUltimoCostoA = "";
                                    ultimoCostoA = 0;
                                }
                                if (registrosComprasA == 0)
                                {
                                    fechaUltimoCostoA = "";
                                    ultimoCostoA = 0.00;
                                }
                                dataGridView1.Rows[i].Cells[34 + consecutivoAlterno].Value = ultimoCostoA;
                                dataGridView1.Rows[i].Cells[41 + consecutivoAlterno].Value = fechaUltimoCosto.Substring(0, 10);
                            }
                            //fin obtener último costo y fecha último costo
                            consecutivoAlterno++;
                        }//fin del While MB alternos
                    }//Fin del using que obtiene códigos hijos de MB

                    double promedioVentas = Math.Round(((sumaVentas-sP) / Convert.ToDouble(numMesActual)),2);
                    double diferenciaPedido = promedioVentas - sumaExistenciasAlterno-sumaExistenciasPrincipal;
                    int mesesPedido = Convert.ToInt16(txtMesesPedido.Text.Trim());
                    //iniciar todos los alternos en cero
                    dataGridView1.Rows[i].Cells[4].Value = "0";
                    dataGridView1.Rows[i].Cells[5].Value = "0";
                    dataGridView1.Rows[i].Cells[49].Value = "0";
                    dataGridView1.Rows[i].Cells[48].Value = "0";
                    if (diferenciaPedido > 0)
                    {
                        dataGridView1.Rows[i].Cells[4].Value = Convert.ToString(Math.Round(promedioVentas - sumaExistenciasPrincipal)*mesesPedido);
                        //dataGridView1.Rows[i].Cells[48].Value = Convert.ToString(Math.Round((promedioVentas - sumaExistenciasPrincipal) * mesesPedido));
                        dataGridView1.Rows[i].Cells[48].Value = Convert.ToString(Math.Round(promedioVentas - sumaExistenciasPrincipal) );

                        dataGridView1.Rows[i].Cells[5].Value = Convert.ToString(Math.Round(promedioVentas - sumaExistenciasPrincipal-sumaExistenciasAlterno)*mesesPedido);
                        //dataGridView1.Rows[i].Cells[49].Value = Convert.ToString(Math.Round((promedioVentas - sumaExistenciasPrincipal-sumaExistenciasAlterno) * mesesPedido));
                        dataGridView1.Rows[i].Cells[49].Value = Convert.ToString(Math.Round(promedioVentas - sumaExistenciasPrincipal - sumaExistenciasAlterno));
                    }
                    else if(diferenciaPedido<0)
                    {
                        dataGridView1.Rows[i].Cells[4].Value = "0";
                        dataGridView1.Rows[i].Cells[48].Value = "0";


                        dataGridView1.Rows[i].Cells[5].Value = "0";
                        dataGridView1.Rows[i].Cells[49].Value = "0";
                    } 

                        i++;
                }//fin del while código principal
            }//fin del 1er Using SQL          
        }//fin del botón excel

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.ColumnIndex == dataGridView1.Columns["Alterno1"].Index)
            {
                var cell = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                cell.ToolTipText = Convert.ToString(string.Format("Existencia: "+ Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[27].Value) + "{0}Fecha de Compra: "+ Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[41].Value) + "{0}Último costo:"+ Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[34].Value) + "", Environment.NewLine));
            }

            if (e.ColumnIndex == dataGridView1.Columns["Alterno2"].Index)
            {
                var cell = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                cell.ToolTipText = Convert.ToString(string.Format("Existencia: " + Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[28].Value) + "{0}Fecha de Compra: " + Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[42].Value) + "{0}Último costo:" + Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[35].Value) + "", Environment.NewLine));
            }

            if (e.ColumnIndex == dataGridView1.Columns["Alterno3"].Index)
            {
                var cell = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                cell.ToolTipText = Convert.ToString(string.Format("Existencia: " + Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[29].Value) + "{0}Fecha de Compra: " + Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[43].Value) + "{0}Último costo:" + Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[36].Value) + "", Environment.NewLine));
            }

            if (e.ColumnIndex == dataGridView1.Columns["Alterno4"].Index)
            {
                var cell = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                cell.ToolTipText = Convert.ToString(string.Format("Existencia: " + Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[30].Value) + "{0}Fecha de Compra: " + Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[44].Value) + "{0}Último costo:" + Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[37].Value) + "", Environment.NewLine));
            }

            if (e.ColumnIndex == dataGridView1.Columns["Alterno5"].Index)
            {
                var cell = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                cell.ToolTipText = Convert.ToString(string.Format("Existencia: " + Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[31].Value) + "{0}Fecha de Compra: " + Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[45].Value) + "{0}Último costo:" + Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[38].Value) + "", Environment.NewLine));
            }

            if (e.ColumnIndex == dataGridView1.Columns["Alterno6"].Index)
            {
                var cell = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                cell.ToolTipText = Convert.ToString(string.Format("Existencia: " + Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[32].Value) + "{0}Fecha de Compra: " + Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[46].Value) + "{0}Último costo:" + Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[39].Value) + "", Environment.NewLine));
            }

            if (e.ColumnIndex == dataGridView1.Columns["Alterno7"].Index)
            {
                var cell = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                cell.ToolTipText = Convert.ToString(string.Format("Existencia: " + Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[33].Value) + "{0}Fecha de Compra: " + Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[47].Value) + "{0}Último costo:" + Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[40].Value) + "", Environment.NewLine));
            }


            if (e.ColumnIndex == dataGridView1.Columns["Ultimo_Costo"].Index)
            {
                var cell = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                cell.ToolTipText = Convert.ToString(string.Format("Fecha de Compra: " + Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[26].Value)));
            }
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.Columns[e.ColumnIndex].Name == "Sug1" || dataGridView1.Columns[e.ColumnIndex].Name == "Sug2")
            {
                dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.FromArgb(171, 235, 198);
            }
        }

        private void btnExcel_Click_1(object sender, EventArgs e)
        {
            GeneraExcel.cargaExcel(dataGridView1,comboCodigo,txtMsg);
            openFile();
        }   

        private void txtMesesPedido_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsDigit(e.KeyChar))
            {
                e.Handled = false;
            }
            else if (Char.IsControl(e.KeyChar))
            {
                e.Handled = false;
            }
            else if (Char.IsSeparator(e.KeyChar))
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
        }

        private void txtMesesPedido_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if (e.KeyData == Keys.Tab)
            {
                int mp = Convert.ToInt16(txtMesesPedido.Text);
                int j = 0;
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    //Application.DoEvents();
                    dataGridView1.Rows[j].Cells[4].Value= Math.Round(Convert.ToDouble(dataGridView1.Rows[j].Cells[48].Value.ToString().Trim())*mp);
                    dataGridView1.Rows[j].Cells[5].Value = Math.Round(Convert.ToDouble(dataGridView1.Rows[j].Cells[49].Value.ToString().Trim()) * mp);
                    j++;
                }

            }
        }
    
        private void button1_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Columns["Alterno1"].Visible==true)
            {
                dataGridView1.Columns["Alterno1"].Visible = false;
                btnSugerido.Text = "+ Sugerido";
            }
            else
            {
                dataGridView1.Columns["Alterno1"].Visible = true;
                btnSugerido.Text = "- Sugerido";
            }

            if (dataGridView1.Columns["Alterno2"].Visible == true)
            {
                dataGridView1.Columns["Alterno2"].Visible = false;
            }
            else
            {
                dataGridView1.Columns["Alterno2"].Visible = true;
            }

            if (dataGridView1.Columns["Alterno3"].Visible == true)
            {
                dataGridView1.Columns["Alterno3"].Visible = false;
            }
            else
            {
                dataGridView1.Columns["Alterno3"].Visible = true;
            }

            if (dataGridView1.Columns["Alterno4"].Visible == true)
            {
                dataGridView1.Columns["Alterno4"].Visible = false;
            }
            else
            {
                dataGridView1.Columns["Alterno4"].Visible = true;
            }

            if (dataGridView1.Columns["Alterno5"].Visible == true)
            {
                dataGridView1.Columns["Alterno5"].Visible = false;
            }
            else
            {
                dataGridView1.Columns["Alterno5"].Visible = true;
            }

            if (dataGridView1.Columns["Alterno6"].Visible == true)
            {
                dataGridView1.Columns["Alterno6"].Visible = false;
            }
            else
            {
                dataGridView1.Columns["Alterno6"].Visible = true;
            }

            if (dataGridView1.Columns["Alterno7"].Visible == true)
            {
                dataGridView1.Columns["Alterno7"].Visible = false;
            }
            else
            {
                dataGridView1.Columns["Alterno7"].Visible = true;
            }
        }

        private void btnAnio_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Columns["Enero"].Visible == true)
            {
                dataGridView1.Columns["Enero"].Visible = false;
                btnAnio.Text = "+ Año";
            }
            else
            {
                dataGridView1.Columns["Enero"].Visible = true;
                btnAnio.Text = "- Año";
            }

            if (dataGridView1.Columns["Febrero"].Visible == true)
            {
                dataGridView1.Columns["Febrero"].Visible = false;
                
            }
            else
            {
                dataGridView1.Columns["Febrero"].Visible = true;
                
            }

            if (dataGridView1.Columns["Marzo"].Visible == true)
            {
                dataGridView1.Columns["Marzo"].Visible = false;
                
            }
            else
            {
                dataGridView1.Columns["Marzo"].Visible = true;
                
            }

            if (dataGridView1.Columns["Abril"].Visible == true)
            {
                dataGridView1.Columns["Abril"].Visible = false;
                
            }
            else
            {
                dataGridView1.Columns["Abril"].Visible = true;
                
            }

            if (dataGridView1.Columns["Mayo"].Visible == true)
            {
                dataGridView1.Columns["Mayo"].Visible = false;
                
            }
            else
            {
                dataGridView1.Columns["Mayo"].Visible = true;
                
            }

            if (dataGridView1.Columns["Junio"].Visible == true)
            {
                dataGridView1.Columns["Junio"].Visible = false;
                
            }
            else
            {
                dataGridView1.Columns["Junio"].Visible = true;
                
            }

            if (dataGridView1.Columns["Julio"].Visible == true)
            {
                dataGridView1.Columns["Julio"].Visible = false;
                
            }
            else
            {
                dataGridView1.Columns["Julio"].Visible = true;
                
            }

            if (dataGridView1.Columns["Agosto"].Visible == true)
            {
                dataGridView1.Columns["Agosto"].Visible = false;
                
            }
            else
            {
                dataGridView1.Columns["Agosto"].Visible = true;
                
            }

            if (dataGridView1.Columns["Septiembre"].Visible == true)
            {
                dataGridView1.Columns["Septiembre"].Visible = false;
                
            }
            else
            {
                dataGridView1.Columns["Septiembre"].Visible = true;
                
            }

            if (dataGridView1.Columns["Octubre"].Visible == true)
            {
                dataGridView1.Columns["Octubre"].Visible = false;
                
            }
            else
            {
                dataGridView1.Columns["Octubre"].Visible = true;
                
            }

            if (dataGridView1.Columns["Noviembre"].Visible == true)
            {
                dataGridView1.Columns["Noviembre"].Visible = false;
                
            }
            else
            {
                dataGridView1.Columns["Noviembre"].Visible = true;
                
            }

            if (dataGridView1.Columns["Diciembre"].Visible == true)
            {
                dataGridView1.Columns["Diciembre"].Visible = false;
                
            }
            else
            {
                dataGridView1.Columns["Diciembre"].Visible = true;
                
            }
        }

        private void openFile()
        {
            string mySheet = @"c:\temp\Pedido_Sugerido_" + comboCodigo.Text + ".xlsx";
            var excelApp = new Excel.Application();
            excelApp.Visible = true;
            Excel.Workbooks books = excelApp.Workbooks;
            Excel.Workbook sheet = books.Open(mySheet);

        }
        private void ProgresoCarga()
        {
            int porciento = Convert.ToInt32((((double)countPro / (double)total) * 100.00));
            if (countPro >= total)
            {
                progressBar1.Value = progressBar1.Maximum;
                lblAvance.Text = progressBar1.Maximum.ToString()+ "%";
            }
            else
            {
                progressBar1.Value = porciento;
                lblAvance.Text = porciento + " %";
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
