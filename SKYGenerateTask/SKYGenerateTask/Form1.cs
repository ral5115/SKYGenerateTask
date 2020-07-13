using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
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

            string filePath = @"C:\Users\Public\Documents\prueba.xlsx";//RUTA DONDE VA A LEER EXCEL
            try
            {
                SLDocument s1 = new SLDocument(filePath);
                int idRow = 2;
                List<DoctoContable> jsonExcel = new List<DoctoContable>();

                while (!string.IsNullOrEmpty(s1.GetCellValueAsString(idRow, 1)))//RECORRE EL EXCEL LEIDO
                {
                    DoctoContable reg = new DoctoContable();

                    var cuenta = s1.GetCellValueAsString(idRow, 1);//VALIDA PARA QUE CUENTA SE VA A IR
                    if (cuenta == "1")
                    {
                        reg.F351_ID_AUXILIAR = "11100503";
                        reg.F351_ID_FE = "110201";
                        reg.F351_ID_TERCERO = "";
                        reg.F351_DOCTO_BANCO = "CG";
                        reg.F351_NRO_DOCTO_BANCO = s1.GetCellValueAsString(idRow, 2);
                    }
                    else if (cuenta == "2")
                    {
                        reg.F351_ID_AUXILIAR = "53051502";
                        reg.F351_ID_FE = "";
                        reg.F351_ID_TERCERO = "890300279";
                        reg.F351_DOCTO_BANCO = "";
                        reg.F351_NRO_DOCTO_BANCO = "0";
                    }
                    else
                    {
                        reg.F351_ID_AUXILIAR = "28050502";
                        reg.F351_ID_FE = "";
                        reg.F351_ID_TERCERO = "890300279";
                        reg.F351_DOCTO_BANCO = "";
                        reg.F351_NRO_DOCTO_BANCO = "0";
                    }
                    reg.F351_NOTAS = "SE RECLASIFICAN PAGOS TU COMPRA DEL DIA " + DateTime.Now.ToString("yyyyMMdd");             
                    
                    var tipo = s1.GetCellValueAsString(idRow, 4);//VALIDA SI ES DEBITO O CREDITO
                    if (tipo == "db")
                        reg.F351_VALOR_DB = s1.GetCellValueAsString(idRow, 3);
                    else
                        reg.F351_VALOR_CR = s1.GetCellValueAsString(idRow, 3);

                    jsonExcel.Add(reg);
                    idRow++;
                }
                string fecha = DateTime.Now.ToString("yyyyMMdd");
                fecha = @"""" + fecha + @"""";//COMILLA DOBLE EN FECHA
                string jsonDocto = JsonConvert.SerializeObject(jsonExcel);//CONVERSION OBJETO EN JSON
                jsonDocto = @"{ ""Conector"": ""Docto_Contable"",""F350_FECHA"":" + fecha //ADICION DE ENCABEZADO EN JSON DE DOCTO CONTABLE
                          + @",""F350_NOTAS"": ""SE RECLASIFICAN PAGOS TU COMPRA DEL DIA "+ DateTime.Now.ToString("yyyyMMdd")
                          + @" "",""Movto_Contable"":" + jsonDocto + "}";

                structure = ejecutar.GetStruct();//CONSULTA LA ESTRUCTURA

                StringBuilder plane = new StringBuilder();
                int consectLine = 1;
                JObject jsonValue = JObject.Parse(jsonDocto);
                List<JObject> value = new List<JObject>();
                value.Add(jsonValue);

                if (value != null)//ARMA EL PLANO CON BASE EN EL JSON Y LA ESTRUCTURA
                {
                    plane.Append(planeBuild.BuildInitial(structure, value[0]));//construye linea inicial

                    for (int j = 0; j < value.Count; j++)//recorre la lista de registros a enviar
                    {

                        string ConectorType = (string)value[j]["Conector"];//extrae el nombre del conector
                        JObject json = value[j];//extrae json del conector a enviar

                        plane.Append(planeBuild.BuildMasters(structure, json, ref consectLine));//construye encabezados o maestros
                        string Pano = plane.ToString();
                        plane.Append(planeBuild.BuildDetails(structure, json, ref consectLine));//construye movimientos


                    }

                    plane.Append(planeBuild.BuildFinal(structure, value[0], ref consectLine));//construye linea final


                    string Plano = plane.ToString();
                    var SavePlane = new StreamWriter($@"C:\Users\Public\Documents\DoctoContable{DateTime.Now.ToString("ddMMyyyy")}.txt");
                    SavePlane.WriteLine(Plano);
                    SavePlane.Close();

                }


            }
            catch (Exception ex)
            {

                throw ex;
            }
            finally
            {
                Application.Exit();
            }
           
        }
    }
}
