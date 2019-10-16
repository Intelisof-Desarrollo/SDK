using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Universal_Venta_Anual_Com_MB
{
    class Conexiones
    {
        public static void conexionSDK(ref int bandera,int lError,TextBox txt)
        {
            string szRegKeySistema = @"SOFTWARE\\Computación en Acción, SA CV\\CONTPAQ I COMERCIAL";
            string sNombrePAQ = "CONTPAQ I COMERCIAL";
            //int lError;
            string ruta;

            RegistryKey keySistema = Registry.LocalMachine.OpenSubKey(szRegKeySistema);
            object lEntrada = keySistema.GetValue("DirectorioBase");

            // SetCurrentDirectory
            long lResult;
            lResult = SDK.SetCurrentDirectory(lEntrada.ToString());

            lError = SDK.fSetNombrePAQ(sNombrePAQ);
            if (lError != 0)
            {
                SDK.rError(lError);
            }
            else
            {
                
                txt.Text= "Se abrió el SDK correctamente.";
                
                bandera = 1;
            }

            ruta = System.Configuration.ConfigurationManager.AppSettings["ruta"];
            lError = SDK.fAbreEmpresa(ruta);
            //return lError;
        }
        public static string CadenaConexionVentasAnuales(string comboCodigo,string anio)
        {
            
            string Cadena = "select CCODIGOPRODUCTO,CNOMBREPRODUCTO,CVALORCLASIFICACION,UBICACION, " +
            "Case when January is not null then January else 0 end as Enero," +
            "Case when February is not null then February else 0 end as Febrero," +
            "Case when March is not null then March else 0 end as Marzo," +
            "Case when April is not null then April else 0 end as Abril," +
            "Case when May is not null then May else 0 end as Mayo," +
            "Case when June is not null then June else 0 end as Junio," +
            "Case when July is not null then July else 0 end as Julio," +
            "Case when August is not null then August else 0 end as Agosto," +
            "Case when September is not null then September else 0 end as Septiembre," +
            "Case when October is not null then October else 0 end as Octubre," +
            "Case when November is not null then November else 0 end as Noviembre," +
            "Case when December is not null then December else 0 end as Diciembre " +
            "from(SELECT P.CCODIGOPRODUCTO, P.CNOMBREPRODUCTO, C.CVALORCLASIFICACION, CONCAT(CZONA, CPASILLO, CANAQUEL, CREPISA) AS UBICACION," +
            "sum(M.CUNIDADES) as TOTAL, DATENAME(month, M.CFECHA) AS FECHA " +
            "FROM(SELECT * FROM admProductos WHERE CIDVALORCLASIFICACION1 = '" + comboCodigo + "') P " +
            "LEFT JOIN admClasificacionesValores C ON P.CIDVALORCLASIFICACION1 = C.CIDVALORCLASIFICACION " +
            "LEFT JOIN(SELECT CIDPRODUCTO, CZONA, CPASILLO, CANAQUEL, CREPISA FROM admMaximosMinimos WHERE CZONA <> '')MM ON P.CIDPRODUCTO = MM.CIDPRODUCTO " +
            "LEFT JOIN(SELECT CIDDOCUMENTO, CIDPRODUCTO, CUNIDADES, CFECHA FROM admMovimientos WHERE(CIDDOCUMENTODE = 4 or CIDDOCUMENTODE = 33 or CIDDOCUMENTODE = 3)AND CAFECTADOSALDOS <> 0 " + //3=REMISIÓN;4=FACTURA; 33=SALIDA
            "AND CFECHA BETWEEN CONCAT('" + anio + "', '', '01', '01') AND CONCAT('" + anio + "', '', '12', '31')) M  ON P.CIDPRODUCTO = M.CIDPRODUCTO " +
            //"LEFT JOIN(SELECT CIDDOCUMENTO FROM admDocumentos WHERE(CCANCELADO = 0 AND CDEVUELTO = 0) or cidconceptodocumento = 4 or cidconceptodocumento = 5 or cidconceptodocumento = 3 or cidconceptodocumento = 3001 or cidconceptodocumento = 3002 or cidconceptodocumento = 3006 or cidconceptodocumento = 35) D ON M.CIDDOCUMENTO = D.CIDDOCUMENTO " +//Sandino=4,5,3,3001,3004,3009: Tamulte:3,4,5,35,3001,3002,3006
            "LEFT JOIN(SELECT CIDDOCUMENTO, cidconceptodocumento FROM admDocumentos WHERE(CCANCELADO = 0 AND CDEVUELTO = 0)) D ON M.CIDDOCUMENTO = D.CIDDOCUMENTO where cidconceptodocumento = 4 or cidconceptodocumento = 5 or cidconceptodocumento = 3 or cidconceptodocumento = 3001 or cidconceptodocumento = 3002 or cidconceptodocumento = 3006 or cidconceptodocumento = 35"+    //Sandino = 4,5,3,3001,3004,3009: Tamulte: 3,4,5,35,3001,3002,3006
            "group by P.CCODIGOPRODUCTO, P.CNOMBREPRODUCTO, C.CVALORCLASIFICACION, CZONA, CPASILLO, CANAQUEL, CREPISA, DATENAME(month, M.CFECHA))t " +
            "pivot(sum(t.Total) " +
            "for FECHA in (January, February, March, April, May, June, July, August, September, October, November, December)) as PVT";
            return Cadena;
        }
        public static string CadenaConexionVentasAnualesPorProducto(string producto, string anio)
        {
            string cadena = "select CCODIGOPRODUCTO,CNOMBREPRODUCTO,"+
           "Case when January is not null then January else 0 end as Enero,"+
           "Case when February is not null then February else 0 end as Febrero,"+
           "Case when March is not null then March else 0 end as Marzo,"+
           "Case when April is not null then April else 0 end as Abril,"+
           "Case when May is not null then May else 0 end as Mayo,"+
           "Case when June is not null then June else 0 end as Junio,"+
           "Case when July is not null then July else 0 end as Julio,"+
           "Case when August is not null then August else 0 end as Agosto,"+
           "Case when September is not null then September else 0 end as Septiembre,"+
           "Case when October is not null then October else 0 end as Octubre,"+
           "Case when November is not null then November else 0 end as Noviembre,"+
           "Case when December is not null then December else 0 end as Diciembre "+
           "from "+
           "(SELECT P.CCODIGOPRODUCTO, P.CNOMBREPRODUCTO, sum(M.CUNIDADES) as TOTAL, DATENAME(month, M.CFECHA) AS FECHA "+
           "FROM(SELECT * FROM admProductos WHERE CCODIGOPRODUCTO = '"+producto+"') P "+
           "INNER JOIN(SELECT CIDDOCUMENTO, CIDPRODUCTO, CUNIDADES, CFECHA FROM admMovimientos WHERE(CIDDOCUMENTODE = 32)AND CAFECTADOSALDOS <> 0 "+//32=ENTRADAS
           "AND CFECHA BETWEEN CONCAT('" + anio + "', '', '01', '01') AND CONCAT('" + anio + "', '', '12', '31')) M  ON P.CIDPRODUCTO = M.CIDPRODUCTO " +
           "INNER JOIN(SELECT CIDDOCUMENTO FROM admDocumentos WHERE(CCANCELADO = 0 AND CDEVUELTO = 0) AND cidconceptodocumento = 34) D ON M.CIDDOCUMENTO = D.CIDDOCUMENTO "+//SANDINO=3008; TAMULTÉ: 34
           "group by P.CCODIGOPRODUCTO, P.CNOMBREPRODUCTO, DATENAME(month, M.CFECHA))t "+
           "pivot(sum(t.Total) "+
           "for FECHA in (January, February, March, April, May, June, July, August, September, October, November, December)) as PVT";
            return cadena;
        }

        public static string ConnCostoFechaSQLServer(string producto)
        {
            return "select CIDMOVIMIENTO cidproducto, ccostocapturado + CPRECIO as costo, cfecha from admmovimientos where cidmovimiento = " +
                                "(select max(cidmovimiento)from admMovimientos where CIDPRODUCTO = " +
                                "(select cidproducto from admProductos where ccodigoproducto = '" + producto + "') " +
                                "and(CIDDOCUMENTODE = 19 or CIDDOCUMENTODE = 18 or CIDDOCUMENTODE = 32) and cidalmacen = '1')";
        }
        public static string ConnCostoFechaSQLServer2(string producto)
        {
            return "select CIDMOVIMIENTO cidproducto, ccostocapturado + CPRECIO as costo, cfecha from admmovimientos where cidmovimiento = " +
                                    "(select max(cidmovimiento)from admMovimientos where CIDPRODUCTO = " +
                                    "(select cidproducto from admProductos where ccodigoproducto = '" + producto + "') " +
                                    "and(CIDDOCUMENTODE = 19 or CIDDOCUMENTODE = 18 or CIDDOCUMENTODE = 32) and cidalmacen = '1')";
        }
    }
}
