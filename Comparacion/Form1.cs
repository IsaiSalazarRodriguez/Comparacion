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
using Microsoft.Office.Interop.Excel;

namespace Comparacion
{
    public partial class Form1 : Form
    {
        private String ruta1;
        private String ruta2;
        private String ruta3;
        private BindingList<Modelo_canasta> superlista = new BindingList<Modelo_canasta>();
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
                archivo3 = new Excel(ruta3, 0);
                imprime3(archivo3.listaHojas);



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
        private void imprime3(List<Hoja> libro)
        {
            listBox5.Items.Clear();
            foreach (Hoja lib in libro)
            {
                listBox5.Items.Add(lib.Nombre);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Hoja nueva = new Hoja(ruta3, listBox5.Items[listBox5.SelectedIndex].ToString());
            archivo3.listaHojas.Add(nueva);
            int indiceA;
            int indiceB;
            dataGridView3.Rows.Clear();
            listBox4.Items.Clear();
            listBox3.Items.Clear();
            /*try
             {*/
            indiceA = listBox1.SelectedIndex;
            indiceB = listBox2.SelectedIndex;
            archivo1.listaHojas[indiceA].obtenTabla();
            archivo2.listaHojas[indiceB].obtenTabla();
            dataGridView1.DataSource = archivo1.listaHojas[indiceA].modelos;
            dataGridView2.DataSource = archivo2.listaHojas[indiceB].modelos;
            Modelo_canasta gabinete = new Modelo_canasta();

            gabinete.Area = archivo2.listaHojas[listBox2.SelectedIndex].Nombre;
            archivo3.listaHojas[listBox5.SelectedIndex].modelosCanasta.Add(gabinete);

            comparaListas();
            Modelo_canasta total = new Modelo_canasta();
            total.EachNet = "Process Level Total";
            archivo3.listaHojas[listBox5.SelectedIndex].modelosCanasta.Add(total);
            dataGridView3.DataSource = archivo3.listaHojas[listBox5.SelectedIndex].modelosCanasta;
            cambiaCeldas();

            /*}
            catch (Exception)
            {
                MessageBox.Show("Selecciona un elemento de ambas listas");
            }*/



        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {

        }

        private void comparaListas()
        {
            int cont = 0;
            bool band1 = true;
            bool band2 = true;
            foreach (Parte par2 in archivo2.listaHojas[listBox2.SelectedIndex].modelos)
            {
                band1 = true;
                foreach (Parte par in archivo1.listaHojas[listBox1.SelectedIndex].modelos)
                {
                    if (par2.parte.Equals(par.parte))
                    {
                        if (par.cantidad.Equals(par2.cantidad) == false)
                        {
                            if (!archivo3.listaHojas[listBox5.SelectedIndex].obtenRenglon(par2.parte, par2.cantidad, "Cantidad", cont))
                            {
                                listBox4.Items.Add(par2.parte);
                            }
                            else
                            {
                                cont++;
                            }
                        }
                        band1 = false;
                        break;
                    }
                }
                if (band1)
                {
                    if (!archivo3.listaHojas[listBox5.SelectedIndex].obtenRenglon(par2.parte, par2.cantidad, "Agregado", cont))
                    {
                        listBox4.Items.Add(par2.parte);
                    }
                    else
                    {
                        cont++;
                    }
                }

            }
            foreach (Parte partesita in archivo1.listaHojas[listBox1.SelectedIndex].modelos)
            {
                band2 = true;
                foreach (Parte partesota in archivo2.listaHojas[listBox2.SelectedIndex].modelos)
                {
                    if (partesota.parte.Equals(partesita.parte))
                    {
                        band2 = false;

                    }
                }
                if (band2)
                {
                    listBox3.Items.Add(partesita.parte);
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
            for (int i = 2; i <= dataGridView3.Rows.Count - 2; i++)
            {
                dataGridView3.Rows[i - 1].Cells[12].Value = "=D" + (i + 1).ToString() + "*K" + (i + 1).ToString();
                dataGridView3.Rows[i - 1].Cells[2].Value = "=MAX($C$1:C" + (i + 1).ToString() + ")+1";
                dataGridView3.Rows[i - 1].Cells[14].Value = "=D" + (i + 1).ToString() + "*L" + (i + 1).ToString();
                dataGridView3.Rows[i - 1].Cells[17].Value = "=D" + (i + 1).ToString() + "*P" + (i + 1).ToString();
                dataGridView3.Rows[i - 1].Cells[19].Value = "=D" + (i + 1).ToString() + "*Q" + (i + 1).ToString();
                dataGridView3.Rows[i - 1].Cells[21].Value = "=D" + (i + 1).ToString() + "*Q" + (i + 1).ToString();
            }
            dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[12].Value = "=SUM(M3:M" + (dataGridView3.Rows.Count - 1).ToString() + ")";
            dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[14].Value = "=SUM(O3:O" + (dataGridView3.Rows.Count - 1).ToString() + ")";
            dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[17].Value = "=SUM(R3:R" + (dataGridView3.Rows.Count - 1).ToString() + ")";
            dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[19].Value = "=SUM(T3:T" + (dataGridView3.Rows.Count - 1).ToString() + ")";


        }
        private void cambiaLista()
        {
            int inicio = 3;
            String aux;
            for (int i = 2; i < superlista.Count; i++)
            {
                if (superlista[i].EachNet != null)
                {
                    if (superlista[i].EachNet.Equals("Process Level Total"))
                    {
                        superlista[i].TotalList = "=SUM(M" + inicio.ToString() + ":M" + (i + 1).ToString() + ")";
                        superlista[i].TotalNet = "=SUM(O" + inicio.ToString() + ":O" + (i + 1).ToString() + ")";
                        superlista[i].TotXferList = "=SUM(R" + inicio.ToString() + ":R" + (i + 1).ToString() + ")";
                        superlista[i].TotXferNet = "=SUM(T" + inicio.ToString() + ":T" + (i + 1).ToString() + ")";
                        inicio++;
                    }
                    else
                    {
                        superlista[i].Item = "=MAX($C$1:C" + (i + 2).ToString() + ")+1";
                        superlista[i].TotalList = "=D" + (i + 2).ToString() + "*K" + (i + 2).ToString();

                        superlista[i].TotalNet = "=D" + (i + 2).ToString() + "*L" + (i + 2).ToString();
                        superlista[i].TotXferList = "=D" + (i + 2).ToString() + "*P" + (i + 2).ToString();
                        superlista[i].TotXferNet = "=D" + (i + 2).ToString() + "*Q" + (i + 2).ToString();
                        superlista[i].TotInitialXfer = "=D" + (i + 2).ToString() + "*Q" + (i + 2).ToString();
                        inicio++;


                    }
                }
                else
                {
                    inicio++;
                }

            }



        }

        private void button6_Click(object sender, EventArgs e)
        {
            Hoja nueva = new Hoja(ruta3, listBox5.Items[listBox5.SelectedIndex].ToString());
            archivo3.listaHojas.Add(nueva);
            int cont = 0;
            dataGridView3.Rows.Clear();
            listBox3.Items.Clear();
            listBox4.Items.Clear();
            try
            {
                Modelo_canasta gabinete = new Modelo_canasta();
                gabinete.Area = archivo2.listaHojas[listBox2.SelectedIndex].Nombre;
                archivo3.listaHojas[listBox5.SelectedIndex].modelosCanasta.Add(gabinete);
                int indiceB = listBox2.SelectedIndex;
                archivo2.listaHojas[indiceB].obtenTabla();
                dataGridView2.DataSource = archivo2.listaHojas[indiceB].modelos;
                foreach (Parte par2 in archivo2.listaHojas[listBox2.SelectedIndex].modelos)
                {
                    if (!archivo3.listaHojas[listBox5.SelectedIndex].obtenRenglon(par2.parte, par2.cantidad, "Cantidad", cont))
                    {
                        listBox4.Items.Add(par2.parte);
                    }
                    else
                    {
                        cont++;
                    }
                }

                Modelo_canasta total = new Modelo_canasta();
                total.EachNet = "Process Level Total";


                archivo3.listaHojas[listBox5.SelectedIndex].modelosCanasta.Add(total);
                dataGridView3.DataSource = archivo3.listaHojas[listBox5.SelectedIndex].modelosCanasta;
                cambiaCeldas();

            }
            catch (Exception)
            {
                MessageBox.Show("Selecciona un elemento de ambas listas");
            }

        }

        private void button7_Click(object sender, EventArgs e)
        {
            Hoja nueva = new Hoja(ruta3, listBox5.Items[listBox5.SelectedIndex].ToString());
            nueva.obtenBasketMaestra();

            archivo3.listaHojas.Add(nueva);
            int ren = 0;
            Modelo_canasta modeloC;
            dataGridView3.Rows.Clear();
            listBox3.Items.Clear();
            listBox4.Items.Clear();
            superlista.Clear();
            try
            {


                for (int j = 0; j < archivo2.listaHojas.Count; j++)
                {
                    Modelo_canasta gabinete = new Modelo_canasta();
                    gabinete.Area = archivo2.listaHojas[j].Nombre;
                    superlista.Add(gabinete);
                    ren++;
                    archivo2.listaHojas[j].obtenTabla();
                    foreach (Parte par2 in archivo2.listaHojas[j].modelos)
                    {
                        foreach (Modelo_canasta mod in archivo3.listaHojas[archivo3.listaHojas.Count-1].basketMaestra)
                        {
                            
                            if (mod.Model.Contains(par2.parte))
                            {
                                modeloC = new Modelo_canasta();
                                modeloC.Area = mod.Area;
                                modeloC.Leavel = mod.Leavel;
                                modeloC.Item = mod.Item;
                                modeloC.Qty = par2.cantidad;
                                modeloC.ReqDate = mod.ReqDate;
                                modeloC.ProductType = mod.ProductType;
                                modeloC.Model = mod.Model;
                                modeloC.AuxSpec1 = mod.AuxSpec1;
                                modeloC.Description = mod.Description;
                                modeloC.LongDescription = mod.LongDescription;
                                modeloC.EachList = mod.EachList;
                                modeloC.EachNet = mod.EachNet;
                                modeloC.TotalList = mod.EachList;
                                modeloC.Discount = mod.Discount;
                                modeloC.TotalNet = mod.TotalNet;
                                modeloC.EachXferList = mod.EachXferList;
                                modeloC.EachXferNet = mod.EachXferNet;
                                modeloC.TotXferList = mod.TotXferList;
                                modeloC.XferDisc = mod.XferDisc;
                                modeloC.TotXferNet = mod.TotXferNet;
                                modeloC.EachInitialXfer = mod.EachInitialXfer;
                                modeloC.TotInitialXfer = mod.TotInitialXfer;
                                modeloC.VendorCode = mod.VendorCode;
                                modeloC.Weight = mod.Weight;
                                modeloC.MarketGroup = mod.MarketGroup;
                                modeloC.setNet = mod.setNet;
                                modeloC.DiscountA = mod.DiscountA;
                                modeloC.DiscountB = mod.DiscountB;
                                modeloC.DiscountC = mod.DiscountC;
                                modeloC.DiscountD = mod.DiscountD;
                                modeloC.DiscountE = mod.DiscountE;
                                modeloC.LeadTime = mod.LeadTime;
                                modeloC.LifeCycle = mod.LifeCycle;
                                modeloC.Country = mod.Country;
                                modeloC.LineItem = mod.LineItem;
                                modeloC.MfgCurrency = mod.MfgCurrency;
                                modeloC.TagSet = mod.TagSet;
                                modeloC.TagQty = mod.TagQty;
                                modeloC.ModeloJornadas = mod.ModeloJornadas;
                                modeloC.EEC = mod.EEC;
                                modeloC.areaP = mod.areaP;
                                modeloC.Sistema = mod.Sistema;
                                superlista.Add(modeloC);
                                break;

                            }
                        }
                    }

                    Modelo_canasta total = new Modelo_canasta();
                    total.EachNet = "Process Level Total";


                    superlista.Add(total);

                }
                cambiaLista();
                dataGridView3.DataSource = superlista;


            }
            catch (Exception)
            {
                MessageBox.Show("Selecciona un elemento de ambas listas");
            }
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void listBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
        }
    }
}
