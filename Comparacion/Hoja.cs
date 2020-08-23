using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Comparacion
{
    
    public class Hoja
    {
        public String path;
        public String Nombre = "";
        public int indice = 0;
        public List<Parte> modelos = new List<Parte>();
        public BindingList<Modelo_canasta> modelosCanasta = new BindingList<Modelo_canasta>();
        public Hoja(String ruta, String nom)
        {
            path = ruta;
            Nombre = nom;
        }

        public void obtenTabla()
        {
            modelos.Clear();
            SLDocument sl = new SLDocument(path,Nombre);
            int r=1;
            while (!sl.GetCellValueAsString(r, 1).Equals("Cantidad"))
            {
                r++;
            }
            r++;
            r++;
            while (!string.IsNullOrEmpty(sl.GetCellValueAsString(r, 1)))
            {
                Parte modelo = new Parte();
                modelo.cantidad = sl.GetCellValueAsString(r, 1);
                modelo.parte = sl.GetCellValueAsString(r, 2);
                modelo.Descripcion = sl.GetCellValueAsString(r, 3);
                modelos.Add(modelo);
                r++;
            }
            sl.CloseWithoutSaving();
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(sl);
        }

        public bool obtenRenglon(String Nmodelo, String cant, String Operacion)
        {
            
            SLDocument sl = new SLDocument(path, Nombre);
            int r = 4;
            while (!sl.GetCellValueAsString(r, 7).Equals(Nmodelo) && r<2355)
            {
                r++;
            }
            if (sl.GetCellValueAsString(r, 7).Equals(Nmodelo))
            {
                Modelo_canasta modeloC = new Modelo_canasta();
                modeloC.Area = "";
                modeloC.Leavel = Operacion;
                modeloC.Item = "="+sl.GetCellFormula(r, 3);
                modeloC.Qty = cant;
                modeloC.ReqDate = sl.GetCellValueAsString(r, 5);
                modeloC.ProductType = sl.GetCellValueAsString(r, 6);
                modeloC.Model = sl.GetCellValueAsString(r, 7);
                modeloC.AuxSpec1 = sl.GetCellValueAsString(r, 8);
                modeloC.Description = sl.GetCellValueAsString(r, 9);
                modeloC.LongDescription = sl.GetCellValueAsString(r, 10);
                modeloC.EachList = sl.GetCellValueAsString(r, 11);
                modeloC.EachNet = sl.GetCellValueAsString(r, 12);
                modeloC.TotalList = "="+ sl.GetCellFormula(r, 13);
                modeloC.Discount = sl.GetCellValueAsString(r, 14);
                modeloC.TotalNet = "="+sl.GetCellFormula(r, 15);
                modeloC.EachXferList = sl.GetCellValueAsString(r, 16);
                modeloC.EachXferNet = sl.GetCellValueAsString(r, 17);
                modeloC.TotXferList = sl.GetCellFormula(r, 18);
                modeloC.XferDisc = sl.GetCellValueAsString(r, 19);
                modeloC.TotXferNet = "="+sl.GetCellFormula(r, 20);
                modeloC.EachInitialXfer = sl.GetCellValueAsString(r, 21);
                modeloC.TotInitialXfer = "="+sl.GetCellFormula(r, 22);
                modeloC.VendorCode = sl.GetCellValueAsString(r, 23);
                modeloC.Weight = sl.GetCellValueAsString(r, 24);
                modeloC.MarketGroup = sl.GetCellValueAsString(r, 25);
                modeloC.setNet = sl.GetCellValueAsString(r, 26);
                modeloC.DiscountA = sl.GetCellValueAsString(r, 27);
                modeloC.DiscountB = sl.GetCellValueAsString(r, 28);
                modeloC.DiscountC = sl.GetCellValueAsString(r, 28);
                modeloC.DiscountD = sl.GetCellValueAsString(r, 30);
                modeloC.DiscountE = sl.GetCellValueAsString(r, 31);
                modeloC.LeadTime = sl.GetCellValueAsString(r, 32);
                modeloC.LifeCycle = sl.GetCellValueAsString(r, 33);
                modeloC.Country = sl.GetCellValueAsString(r, 34);
                modeloC.LineItem = sl.GetCellValueAsString(r, 35);
                modeloC.MfgCurrency = sl.GetCellValueAsString(r, 36);
                modeloC.TagSet = sl.GetCellValueAsString(r, 37);
                modeloC.TagQty = sl.GetCellValueAsString(r, 38);
                modeloC.ModeloJornadas = sl.GetCellValueAsString(r, 39);
                modeloC.EEC = sl.GetCellValueAsString(r, 40);

                modelosCanasta.Add(modeloC);
                sl.CloseWithoutSaving();
                return true;
            }
            else
            {
                sl.CloseWithoutSaving();
                return false;
            }
            
            
        }
    }
}
