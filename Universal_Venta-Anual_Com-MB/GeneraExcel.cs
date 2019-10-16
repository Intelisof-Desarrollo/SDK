using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;

namespace Universal_Venta_Anual_Com_MB
{
    class GeneraExcel
    {
        public static void cargaExcel(DataGridView dataGridView1,ComboBox comboCodigo,System.Windows.Forms.TextBox txtMsg)
        {
            int i = 0;
            ExcelPackage ExcelPkg = new ExcelPackage();
            ExcelWorksheet wsSheet1 = ExcelPkg.Workbook.Worksheets.Add("Sheet1");
            SaveFileDialog fichero = new SaveFileDialog();
            fichero.Filter = "Excel (*.xls)|*.xls";

            using (ExcelPackage excel = new ExcelPackage())
            {
                ExcelWorksheet hoja = excel.Workbook.Worksheets.Add("Pedidos sugeridos");
                hoja.Cells["a1:z3"].Merge = true;
                hoja.Cells[1, 1].Value = "Reporte de pedidos sugeridos";
                hoja.Cells["a1:z3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                hoja.Cells["a1:z3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                hoja.Cells["a1:z3"].Style.Font.Size = 22;
                hoja.Cells["a1:z3"].Style.Font.Bold = true;

                hoja.Cells[4, 1].Value = "Descripción producto";
                hoja.Cells[4, 2].Value = "Código padre";
                hoja.Cells[4, 3].Value = "Existencia padre";
                hoja.Cells[4, 4].Value = "Venta total";
                hoja.Cells[4, 5].Value = "Sugerido 1";
                hoja.Cells[4, 6].Value = "Sugerido 2";
                hoja.Cells[4, 7].Value = "Último costo padre";
                hoja.Cells[4, 8].Value = "Alterno 1";
                hoja.Cells[4, 9].Value = "Alterno 2";
                hoja.Cells[4, 10].Value = "Alterno 3";
                hoja.Cells[4, 11].Value = "alterno 4";
                hoja.Cells[4, 12].Value = "Alterno 5";
                hoja.Cells[4, 13].Value = "Alterno 6";
                hoja.Cells[4, 14].Value = "Alterno 7";
                hoja.Cells[4, 15].Value = "Enero";
                hoja.Cells[4, 16].Value = "Febrero";
                hoja.Cells[4, 17].Value = "Marzo";
                hoja.Cells[4, 18].Value = "Abril";
                hoja.Cells[4, 19].Value = "Mayo";
                hoja.Cells[4, 20].Value = "Junio";
                hoja.Cells[4, 21].Value = "Julio";
                hoja.Cells[4, 22].Value = "Agosto";
                hoja.Cells[4, 23].Value = "Septiembre";
                hoja.Cells[4, 24].Value = "Octubre";
                hoja.Cells[4, 25].Value = "Noviembre";
                hoja.Cells[4, 26].Value = "Diciembre";


                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    //System.Windows.Forms.Application.DoEvents();
                    hoja.Cells[i + 5, 1].Value = dataGridView1.Rows[i].Cells[0].Value;
                    hoja.Cells[i + 5, 2].Value = dataGridView1.Rows[i].Cells[1].Value;
                    hoja.Cells[i + 5, 3].Value = dataGridView1.Rows[i].Cells[2].Value;
                    hoja.Cells[i + 5, 4].Value = dataGridView1.Rows[i].Cells[3].Value;
                    hoja.Cells[i + 5, 5].Value = dataGridView1.Rows[i].Cells[4].Value;
                    hoja.Cells[i + 5, 6].Value = dataGridView1.Rows[i].Cells[5].Value;
                    hoja.Cells[i + 5, 7].Value = dataGridView1.Rows[i].Cells[6].Value;
                    var commentP = hoja.Cells[i + 5, 7].AddComment("Adicional:", "Exist-Costo-Fecha");
                    var rtP = commentP.RichText.Add("\r\nFech Ult Compra:" + dataGridView1.Rows[i].Cells[26].Value + "\r\n");
                    rtP.Bold = false;
                    commentP.AutoFit = true;

                    hoja.Cells[i + 5, 8].Value = dataGridView1.Rows[i].Cells[7].Value;
                    var comment1 = hoja.Cells[i + 5, 8].AddComment("Adicional:", "Exist-Costo-Fecha");
                    var rt1 = comment1.RichText.Add("\r\nExistencia:" + dataGridView1.Rows[i].Cells[27].Value + "\r\nFech Ult Compra:" + dataGridView1.Rows[i].Cells[41].Value + "\r\nUltimo Costo:" + dataGridView1.Rows[i].Cells[34].Value + "");
                    rt1.Bold = false;
                    comment1.AutoFit = true;
                    hoja.Cells[i + 5, 9].Value = dataGridView1.Rows[i].Cells[8].Value;
                    var comment2 = hoja.Cells[i + 5, 9].AddComment("Adicional:", "Exist-Costo-Fecha");
                    var rt2 = comment2.RichText.Add("\r\nExistencia:" + dataGridView1.Rows[i].Cells[28].Value + "\r\nFech Ult Compra:" + dataGridView1.Rows[i].Cells[42].Value + "\r\nUltimo Costo:" + dataGridView1.Rows[i].Cells[35].Value + "");
                    rt2.Bold = false;
                    comment2.AutoFit = true;
                    hoja.Cells[i + 5, 10].Value = dataGridView1.Rows[i].Cells[9].Value;
                    var comment3 = hoja.Cells[i + 5, 10].AddComment("Adicional:", "Exist-Costo-Fecha");
                    var rt3 = comment3.RichText.Add("\r\nExistencia:" + dataGridView1.Rows[i].Cells[29].Value + "\r\nFech Ult Compra:" + dataGridView1.Rows[i].Cells[43].Value + "\r\nUltimo Costo:" + dataGridView1.Rows[i].Cells[36].Value + "");
                    rt3.Bold = false;
                    comment3.AutoFit = true;
                    hoja.Cells[i + 5, 11].Value = dataGridView1.Rows[i].Cells[10].Value;
                    var comment4 = hoja.Cells[i + 5, 11].AddComment("Adicional:", "Exist-Costo-Fecha");
                    var rt4 = comment4.RichText.Add("\r\nExistencia:" + dataGridView1.Rows[i].Cells[30].Value + "\r\nFech Ult Compra:" + dataGridView1.Rows[i].Cells[44].Value + "\r\nUltimo Costo:" + dataGridView1.Rows[i].Cells[37].Value + "");
                    rt4.Bold = false;
                    comment4.AutoFit = true;
                    hoja.Cells[i + 5, 12].Value = dataGridView1.Rows[i].Cells[11].Value;
                    var comment5 = hoja.Cells[i + 5, 12].AddComment("Adicional:", "Exist-Costo-Fecha");
                    var rt5 = comment5.RichText.Add("\r\nExistencia:" + dataGridView1.Rows[i].Cells[31].Value + "\r\nFech Ult Compra:" + dataGridView1.Rows[i].Cells[45].Value + "\r\nUltimo Costo:" + dataGridView1.Rows[i].Cells[38].Value + "");
                    rt5.Bold = false;
                    comment5.AutoFit = true;
                    hoja.Cells[i + 5, 13].Value = dataGridView1.Rows[i].Cells[12].Value;
                    var comment6 = hoja.Cells[i + 5, 13].AddComment("Adicional:", "Exist-Costo-Fecha");
                    var rt6 = comment6.RichText.Add("\r\nExistencia:" + dataGridView1.Rows[i].Cells[32].Value + "\r\nFech Ult Compra:" + dataGridView1.Rows[i].Cells[46].Value + "\r\nUltimo Costo:" + dataGridView1.Rows[i].Cells[39].Value + "");
                    rt6.Bold = false;
                    comment1.AutoFit = true;
                    hoja.Cells[i + 5, 14].Value = dataGridView1.Rows[i].Cells[13].Value;
                    var comment7 = hoja.Cells[i + 5, 14].AddComment("Adicional:", "Exist-Costo-Fecha");
                    var rt7 = comment7.RichText.Add("\r\nExistencia:" + dataGridView1.Rows[i].Cells[33].Value + "\r\nFech Ult Compra:" + dataGridView1.Rows[i].Cells[47].Value + "\r\nUltimo Costo:" + dataGridView1.Rows[i].Cells[40].Value + "");
                    rt7.Bold = false;
                    comment7.AutoFit = true;
                    hoja.Cells[i + 5, 15].Value = dataGridView1.Rows[i].Cells[14].Value;
                    hoja.Cells[i + 5, 16].Value = dataGridView1.Rows[i].Cells[15].Value;
                    hoja.Cells[i + 5, 17].Value = dataGridView1.Rows[i].Cells[16].Value;
                    hoja.Cells[i + 5, 18].Value = dataGridView1.Rows[i].Cells[17].Value;
                    hoja.Cells[i + 5, 19].Value = dataGridView1.Rows[i].Cells[18].Value;
                    hoja.Cells[i + 5, 20].Value = dataGridView1.Rows[i].Cells[19].Value;
                    hoja.Cells[i + 5, 21].Value = dataGridView1.Rows[i].Cells[20].Value;
                    hoja.Cells[i + 5, 22].Value = dataGridView1.Rows[i].Cells[21].Value;
                    hoja.Cells[i + 5, 23].Value = dataGridView1.Rows[i].Cells[22].Value;
                    hoja.Cells[i + 5, 24].Value = dataGridView1.Rows[i].Cells[23].Value;
                    hoja.Cells[i + 5, 25].Value = dataGridView1.Rows[i].Cells[24].Value;
                    hoja.Cells[i + 5, 26].Value = dataGridView1.Rows[i].Cells[25].Value;
                    i++;
                } //fin del foreach del datagrid
                hoja.Cells.AutoFitColumns();
                excel.SaveAs(new FileInfo(@"c:\temp\Pedido_Sugerido_" + comboCodigo.Text + ".xlsx"));
                MessageBox.Show("Se ha enviado la información satisfactoriamente!");
                txtMsg.Text = "Se ha enviado la información satisfactoriamente!";
                //openFile(comboCodigo,);           
            }//fin del using de excel

        }

        //private void openFile()
        //{
        //    string mySheet = @"c:\temp\Pedido_Sugerido_" + comboCodigo.Text + ".xlsx";
        //    var excelApp = new Excel.Application();
        //    excelApp.Visible = true;
        //    Excel.Workbooks books = excelApp.Workbooks;
        //    Excel.Workbook sheet = books.Open(mySheet);

        //}
    }
}
