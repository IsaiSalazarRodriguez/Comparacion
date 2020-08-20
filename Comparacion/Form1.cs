using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Data.Odbc;
namespace Comparacion
{
    public partial class Form1 : Form
    {
        private String ruta1;
        private String ruta2;
        private String ruta3;
        Excel archivo1;
        Excel archivo2;
        Excel archivo3;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                ruta1 = openFileDialog1.FileName;
                textBox1.Text = ruta1;
                archivo1 = new Excel(ruta1);
                imprime(archivo1.listaHojas);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                ruta2 = openFileDialog1.FileName;
                textBox2.Text = ruta2;
                archivo2 = new Excel(ruta2);
                imprime2(archivo2.listaHojas);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                ruta3 = openFileDialog1.FileName;
                textBox3.Text = ruta3;
                archivo3 = new Excel();
                Hoja nueva = new Hoja(ruta3, "BASKET");
                archivo3.listaHojas.Add(nueva);
            }
        }

        private void imprime(List<Hoja> libro)
        {
            listBox1.Items.Clear();
            foreach (Hoja lib in libro)
            {
                listBox1.Items.Add(lib.Nombre);
            }
        }
        private void imprime2(List<Hoja> libro)
        {
            listBox2.Items.Clear();
            foreach (Hoja lib in libro)
            {
                listBox2.Items.Add(lib.Nombre);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            int indiceA;
            int indiceB;
            dataGridView3.Rows.Clear();
            
           try
            {
                indiceA = listBox1.SelectedIndex;
                indiceB = listBox2.SelectedIndex;
                archivo1.listaHojas[indiceA].obtenTabla();
                archivo2.listaHojas[indiceB].obtenTabla();
                dataGridView1.DataSource = archivo1.listaHojas[indiceA].modelos;
                dataGridView2.DataSource = archivo2.listaHojas[indiceB].modelos;
                comparaListas();
                dataGridView3.DataSource = archivo3.listaHojas[0].modelosCanasta;
                cambiaCeldas();

            }
            catch (Exception)
            {
                MessageBox.Show("Selecciona un elemento de ambas listas");
            }



        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {

        }

        private void comparaListas()
        {
            bool band1 = true;
            bool band2 = false;
            foreach(Parte par2 in archivo2.listaHojas[listBox2.SelectedIndex].modelos)
            {
                band1 = true;
                foreach(Parte par in archivo1.listaHojas[listBox1.SelectedIndex].modelos)
                {
                    if (par2.parte.Equals(par.parte))
                    {
                        if (par.cantidad != par2.cantidad)
                        {
                            archivo3.listaHojas[0].obtenRenglon(par2.parte, par2.cantidad, "Cambio cantidad de " + archivo2.listaHojas[listBox2.SelectedIndex].Nombre);
                        }
                        band1 = false;
                        break;
                    }
                }
                if (band1)
                {
                    archivo3.listaHojas[0].obtenRenglon(par2.parte, par2.cantidad, "Nueva parte de " + archivo2.listaHojas[listBox2.SelectedIndex].Nombre);
                }
                
            }
        }

        public void exportaraexcel(DataGridView tabla)
        {

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            excel.Application.Workbooks.Add(true);

            int IndiceColumna = 0;

            foreach (DataGridViewColumn col in tabla.Columns) // Columnas
            {

                IndiceColumna++;

                excel.Cells[1, IndiceColumna] = col.Name;

            }

            int IndeceFila = 0;

            foreach (DataGridViewRow row in tabla.Rows) // Filas
            {

                IndeceFila++;

                IndiceColumna = 0;

                foreach (DataGridViewColumn col in tabla.Columns)
                {

                    IndiceColumna++;

                    excel.Cells[IndeceFila + 1, IndiceColumna] = row.Cells[col.Name].Value;

                }

            }

            excel.Visible = true;


        }

        private void button5_Click(object sender, EventArgs e)
        {
            exportaraexcel(dataGridView3);
        }

        private void cambiaCeldas()
        {
            for(int i = 1; i<=dataGridView3.Rows.Count; i++)
            {
                dataGridView3.Rows[i-1].Cells[13].Value = "=D" + i.ToString() + "*K4";
            }
        }
    }
}
