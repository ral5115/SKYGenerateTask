using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SKYGenerateTask
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string filePath = @"C:\Users\Public\Documents\prueba2.xlsx";
             try
            {
                SLDocument s1 = new SLDocument(filePath);

           
                int idRow = 1;
                while (!string.IsNullOrEmpty(s1.GetCellValueAsString(idRow, 1)))
                {
                    var a = s1.GetCellValueAsString(idRow, 1);
                    var b = s1.GetCellValueAsString(idRow, 2);
                    var c = s1.GetCellValueAsString(idRow, 3);
                    idRow++;
                }
                    
            }
            catch (Exception ex)
            {

                throw ex;
            }
           
        }
    }
}
