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

            string filePath = @"C:\Interfaz_tu_compra\Excel_de_carque";//RUTA DONDE VA A LEER EXCEL
            try
            {
                //MessageBox.Show("inicio proceso");

                DirectoryInfo di = new DirectoryInfo(filePath);
                foreach (var fi in di.GetFiles("*.xlsx"))
                {
                    //MessageBox.Show("leyo archivo");

                    var file = filePath + @"\" + fi.Name;                                       
                    SLDocument s1 = new SLDocument(file);
                    int idRow = 2;
                    List<DoctoContable> jsonExcel = new List<DoctoContable>();

                    while (!string.IsNullOrEmpty(s1.GetCellValueAsString(idRow, 1)))//RECORRE EL EXCEL LEIDO
                    {
                        DoctoContable reg = new DoctoContable();

                        var cuenta = s1.GetCellValueAsString(idRow, 1);//VALIDA PARA QUE CUENTA SE VA A IR
                        if (cuenta == "11100503")
                        {
                            reg.F351_ID_AUXILIAR = "11100503";
                            reg.F351_ID_FE = "110201";
                            reg.F351_ID_TERCERO = "";
                            reg.F351_DOCTO_BANCO = "CG";
                            reg.F351_NRO_DOCTO_BANCO = s1.GetCellValueAsDateTime(idRow, 3).ToString("yyyyMMdd");
                        }
                        else if (cuenta == "53051502")
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
                        if (tipo == "DB")
                            reg.F351_VALOR_DB = s1.GetCellValueAsString(idRow, 2);
                        else
                            reg.F351_VALOR_CR = s1.GetCellValueAsString(idRow, 2);

                        jsonExcel.Add(reg);
                        idRow++;
                    }

                    

                    string fecha = DateTime.Now.ToString("yyyyMMdd");
                    fecha = @"""" + fecha + @"""";//COMILLA DOBLE EN FECHA
                    string jsonDocto = JsonConvert.SerializeObject(jsonExcel);//CONVERSION OBJETO EN JSON
                    jsonDocto = @"{ ""Conector"": ""Docto_Contable"",""F350_FECHA"":" + fecha //ADICION DE ENCABEZADO EN JSON DE DOCTO CONTABLE
                              + @",""F350_NOTAS"": ""SE RECLASIFICAN PAGOS TU COMPRA DEL DIA " + DateTime.Now.ToString("yyyyMMdd")
                              + @" "",""Movto_Contable"":" + jsonDocto + "}";

                    structure = ejecutar.GetStruct();//CONSULTA LA ESTRUCTURA

                    StringBuilder plane = new StringBuilder();
                    int consectLine = 1;
                    JObject jsonValue = JObject.Parse(jsonDocto);
                    List<JObject> value = new List<JObject>();
                    value.Add(jsonValue);

                    //MessageBox.Show("armo json");

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

                        //MessageBox.Show("armo plano");

                        string Plano = plane.ToString();
                        var SavePlane = new StreamWriter($@"C:\Interfaz_tu_compra\Txt_unoee\{fi.Name}.txt");
                        SavePlane.WriteLine(Plano);
                        SavePlane.Close();

                        //MessageBox.Show("guardo plano");

                        fi.MoveTo($@"C:\Interfaz_tu_compra\Excel_de_carque\procesados\{fi.Name}");

                        //MessageBox.Show("movio plano");

                    }
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                throw ex;
            }
            finally
            {
                Application.Exit();
            }
           
        }
    }
}
