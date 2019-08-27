using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
namespace Excell
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
               

        private void Button1_Click(object sender, EventArgs e)
        {
            Load(@"F:\Mapeo1.xlsx", 2, 2);

          

        }
        
        public  void Load(string fileName, int sheetNumber, int columns)
        {
            ExcelReader excel = new ExcelReader(fileName, sheetNumber);

            int count = columns; // excel.RowsCount;
            
            for (int i = 0; i < count; i++)
            {               
                List<Field> headers = excel.getFields(i + 2);
                               
                foreach (Field field in headers)
                {                    
                    field.AddField();                   
                }
               
            }

            MessageBox.Show("Se agregó la tabla");
        }

    }
}
