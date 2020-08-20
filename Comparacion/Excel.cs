using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
namespace Comparacion
{
    class Excel
    {
        public List<Hoja> listaHojas = new List<Hoja>();
        _Application excel = new _Excel.Application();
        public Workbook wb;
        public Worksheet ws;
        String path;
        public Excel(String path)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            cargaLista();
            
            
        }

        public Excel() {
            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            KillSpecificExcelFileProcess(path);
        }

        private void cargaLista()
        {
            Hoja aux;
            for(int i = 1; i<wb.Worksheets.Count; i++)
            {

                ws = wb.Worksheets[i];
                if (ws.Name.IndexOf("peso", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    aux = new Hoja(path, ws.Name);
                    aux.Nombre = ws.Name;
                    aux.indice = i;
                    listaHojas.Add(aux);
                }
            }
            wb.Close(0);
            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
            KillSpecificExcelFileProcess(path);
           

        }
        private void KillSpecificExcelFileProcess(string excelFileName)
        {
            var processes = from p in Process.GetProcessesByName("EXCEL")
                            select p;

            foreach (var process in processes)
            {
                if (process.MainWindowTitle == "Microsoft Excel - " + excelFileName)
                    process.Kill();
            }
        }



    }
}
