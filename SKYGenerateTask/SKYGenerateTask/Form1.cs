using Newtonsoft.Json;
using SKYGenerateTask.Clases;
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
            SQLTransaction ejecutar = new SQLTransaction();
            DataSet structure = new DataSet();
            PlaneBuilder planeBuild = new PlaneBuilder();

            string filePath = @"C:\Users\Public\Documents\prueba.xlsx";
             try
            {
                SLDocument s1 = new SLDocument(filePath);           
                int idRow = 2;
                List<DoctoContable> json = new List<DoctoContable>();
                while (!string.IsNullOrEmpty(s1.GetCellValueAsString(idRow, 1)))
                {
                    DoctoContable reg = new DoctoContable();

                    reg.F351_ID_AUXILIAR = s1.GetCellValueAsString(idRow, 1);
                    reg.F351_ID_TERCERO = s1.GetCellValueAsString(idRow, 2);

                    var tipo = s1.GetCellValueAsString(idRow, 4);
                    if (tipo == "db")                    
                        reg.F351_VALOR_DB = s1.GetCellValueAsString(idRow, 3);
                    else
                        reg.F351_VALOR_CR = s1.GetCellValueAsString(idRow, 3);


                    json.Add(reg);
                    idRow++;
                }

                string jsonDocto = JsonConvert.SerializeObject(json);
                jsonDocto = @"[{""Conector"": ""Docto_Contable"", ""Movto_Contable"":" + jsonDocto+"}]";

                structure = ejecutar.GetStruct();


            }
            catch (Exception ex)
            {

                throw ex;
            }
           
        }
    }
}
