using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Universal_Venta_Anual_Com_MB
{
    class Datos
    {
        private static readonly string conexionsqlXis = ConfigurationManager.ConnectionStrings["ConnSQL"].ConnectionString;
        public static DataTable ObtenerListaDeClasificaciones()
        {
            try
            {
                using (SqlConnection conexion = new SqlConnection(conexionsqlXis))
                {
                    conexion.Open();

                    const string sql = "Select CIDVALORCLASIFICACION,CVALORCLASIFICACION from admClasificacionesValores WHERE CIDCLASIFICACION = 25";
                    SqlCommand cmd = new SqlCommand(sql, conexion);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();

                    da.Fill(dt);

                    return dt;
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException("No fué posible obtener la lista de proveedores", ex);
            }

        }


    }
}
