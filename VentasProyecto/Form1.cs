using System;
using System.Windows.Forms;
using xls = Microsoft.Office.Interop.Excel;
using System.Windows.Forms.DataVisualization.Charting;

namespace VentasProyecto
{
    public partial class Form1 : Form
    {
        xls.Application a = new xls.Application();
        int i = 3;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            a.Workbooks.Open(Application.StartupPath + @"\FormatoProyecto.xlsx");
            while (a.ActiveWorkbook.ActiveSheet.Cells(i, 1).Value != null)
            {
                i++;
            }
        }

        private void btnLeer_Click(object sender, EventArgs e)
        {
            lvPedido.Items.Clear();
            int x = 3;
            int y = 6;
            lblRecepcion.Text = a.ActiveWorkbook.ActiveSheet.Cells(x, 1).Value.ToString();
            lblInsp.Text = a.ActiveWorkbook.ActiveSheet.Cells(x+1, 1).Value.ToString();
            lblfRec.Text = a.ActiveWorkbook.ActiveSheet.Cells(x, 4).Value.ToString();
            lblfCont.Text = a.ActiveWorkbook.ActiveSheet.Cells(x + 1, 4).Value.ToString();
            lblBodega.Text = a.ActiveWorkbook.ActiveSheet.Cells(x, 7).Value.ToString();
            lblfeCont.Text = a.ActiveWorkbook.ActiveSheet.Cells(x + 1, 7).Value.ToString();
            lblCodigo.Text = a.ActiveWorkbook.ActiveSheet.Cells(x, 9).Value.ToString();

            while (a.ActiveWorkbook.ActiveSheet.Cells(y, 1).Value != null)
            {
                string componente = a.ActiveWorkbook.ActiveSheet.Cells(y, 1).Value.ToString();
                string codigo = a.ActiveWorkbook.ActiveSheet.Cells(y, 2).Value.ToString();
                string pedido = a.ActiveWorkbook.ActiveSheet.Cells(y, 3).Value.ToString();
                string totalrec = a.ActiveWorkbook.ActiveSheet.Cells(y, 4).Value.ToString();
                string totalaprob = a.ActiveWorkbook.ActiveSheet.Cells(y, 5).Value.ToString();
                string porcentajeaprobado = a.ActiveWorkbook.ActiveSheet.Cells(y, 6).Value.ToString();
                string totalrechazado = a.ActiveWorkbook.ActiveSheet.Cells(y, 7).Value.ToString();
                string porcentajerechazado = a.ActiveWorkbook.ActiveSheet.Cells(y, 8).Value.ToString();
                string conforme = a.ActiveWorkbook.ActiveSheet.Cells(y, 9).Value.ToString();
                string mot = a.ActiveWorkbook.ActiveSheet.Cells(y, 10).Value.ToString();

                ListViewItem lista = new ListViewItem(componente);
                lista.SubItems.Add(codigo);
                lista.SubItems.Add(pedido);
                lista.SubItems.Add(totalrec);
                lista.SubItems.Add(totalaprob);
                lista.SubItems.Add(porcentajeaprobado + "%");
                lista.SubItems.Add(totalrechazado);
                lista.SubItems.Add(porcentajerechazado + "%");
                lista.SubItems.Add(conforme);
                lista.SubItems.Add(mot);
                lvPedido.Items.Add(lista);

                y++;
            }

        }

        private void btnInsertar_Click(object sender, EventArgs e)
        {
            string comp = txtComponente.Text;
            string cod = txtCodigo.Text;
            string pedido = txtCantPedido.Text;
            decimal totrec = Convert.ToDecimal(txtTotalRecibido.Text);
            decimal totap = Convert.ToDecimal(txtTotalAprobado.Text);
            decimal porcentaceptado;
            decimal totrech = 0;
            decimal porcentajerech;
            string pedidoconforme;
            string motivo;


            txtComponente.Clear();
            txtCodigo.Clear();
            txtCantPedido.Clear();
            txtTotalRecibido.Clear();
            txtTotalAprobado.Clear();

            totrech = totrec - totap;
            porcentaceptado = ((totap * 100) / totrec);
            porcentajerech = ((totrech * 100) / totrec);

            if (totrech>0)
            {
                pedidoconforme = "NO";
            }
            else
            {
                pedidoconforme = "SÍ";
            }
            if (porcentajerech>15)
            {
                motivo = "NO";
            }
            else
            {
                motivo = "SÍ";
            }
            a.ActiveWorkbook.Worksheets[1].Cells(i, 1).Value = comp;
            a.ActiveWorkbook.Worksheets[1].Cells(i, 2).Value = cod;
            a.ActiveWorkbook.Worksheets[1].Cells(i, 3).Value = pedido;
            a.ActiveWorkbook.Worksheets[1].Cells(i, 4).Value = totrec;
            a.ActiveWorkbook.Worksheets[1].Cells(i, 5).Value = totap;
            a.ActiveWorkbook.Worksheets[1].Cells(i, 6).Value = porcentaceptado;
            a.ActiveWorkbook.Worksheets[1].Cells(i, 7).Value = totrech;
            a.ActiveWorkbook.Worksheets[1].Cells(i, 8).Value = porcentajerech;
            a.ActiveWorkbook.Worksheets[1].Cells(i, 9).Value = pedidoconforme;
            a.ActiveWorkbook.Worksheets[1].Cells(i, 10).Value = motivo;

            i++;
            a.ActiveWorkbook.Save();
            MessageBox.Show("Se agregaron los datos al excel", "Lectura y Escritura", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnCambiar_Click(object sender, EventArgs e)
        {
            int y = 3;

            string nrecep;
            string fecharecep;
            string bodega;
            string inspector;
            string fecharecibo;
            string fechaentrega;
            string codigo;

            nrecep = "Recepción #: " + txtRecep.Text;
            fecharecep = "Fecha de recepción: " + fechaRecep.Text;
            bodega = "Bodega: " + txtBodega.Text;
            inspector = "Inspector: " + txtInspector.Text;
            fecharecibo = "Fecha de recibo de control: " + fechaRecibo.Text;
            fechaentrega = "Fecha de entrega de control: " + fechaEntrega.Text;
            codigo = "Código: " + txtCodigo2.Text;

            txtRecep.Clear();
            fechaRecep.ResetText();
            txtBodega.Clear();
            txtInspector.Clear();
            fechaRecibo.ResetText();
            fechaEntrega.ResetText();
            txtCodigo2.Clear();

            a.ActiveWorkbook.Worksheets[1].Cells(y, 1).Value = nrecep;
            a.ActiveWorkbook.Worksheets[1].Cells(y, 4).Value = fecharecep;
            a.ActiveWorkbook.Worksheets[1].Cells(y, 7).Value = bodega;
            a.ActiveWorkbook.Worksheets[1].Cells(y, 9).Value = codigo;
            a.ActiveWorkbook.Worksheets[1].Cells(y + 1, 1).Value = inspector;
            a.ActiveWorkbook.Worksheets[1].Cells(y + 1, 4).Value = fecharecibo;
            a.ActiveWorkbook.Worksheets[1].Cells(y + 1, 7).Value = fechaentrega;

            a.ActiveWorkbook.Save();
            MessageBox.Show("Se agregaron los datos al excel", "Lectura y Escritura", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            a.ActiveWorkbook.Close();
        }

        private void btnGraficas_Click(object sender, EventArgs e)
        {

            chartTotales.Visible = false;
            chartConforme.Visible = false;
            chartTotales.Series.Clear();
            chartConforme.Series.Clear();
            double totalaprobado = 0;
            double totalrechazado = 0;
            double total = 0;
            double porcentajeaprobado;
            double porcentajerechazado;
            double totalt = 0;
            double porcentajeconforme;
            double porcentajenoconforme;
            double conf = 0;
            double noconf = 0;
          
            int y = 6;

            while (a.ActiveWorkbook.ActiveSheet.Cells(y, 4).Value != null)
            {
                total += a.ActiveWorkbook.ActiveSheet.Cells(y, 4).Value;
                totalaprobado += a.ActiveWorkbook.ActiveSheet.Cells(y, 5).Value;
                totalrechazado += a.ActiveWorkbook.ActiveSheet.Cells(y, 7).Value;
                if (a.ActiveWorkbook.ActiveSheet.Cells(y, 10).Value=="SÍ")
                {
                    conf++;
                }
                if (a.ActiveWorkbook.ActiveSheet.Cells(y, 10).Value == "NO")
                {
                    noconf++;
                }
                y++;
            }

            porcentajeaprobado = ((totalaprobado / total) * 100);
            porcentajerechazado = ((totalrechazado / total) * 100);


            porcentajeconforme = ((conf / y) * 100);
            porcentajenoconforme = ((noconf / y) * 100);

            string[] pedido = { "Aprobado", "Rechazado" };
            double[] totales = { porcentajeaprobado, porcentajerechazado };

            for (int j = 0; j < pedido.Length; j++)
            {
                Series serie = chartTotales.Series.Add(pedido[j]);
                serie.Label = totales[j].ToString();
                serie.Points.Add(totales[j]);
            }

            string[] conformidad = { "SÍ", "NO"};
            double[] totalconf = { porcentajeconforme, porcentajenoconforme };

            for (int k = 0; k < conformidad.Length; k++)
            {
                Series serie = chartConforme.Series.Add(conformidad[k]);
                serie.Label = totalconf[k].ToString();
                serie.Points.Add(totalconf[k]);
            }


            chartTotales.Visible = true;
            chartConforme.Visible = true;
            lblAprobado.Text = porcentajeaprobado + "%";
            lblRechazado.Text = porcentajerechazado + "%";
            lblConformidad.Text = porcentajeconforme + "%";
            lblNoConformidad.Text = porcentajenoconforme + "%";
        }
    }
}