using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SKYGenerateTask.Clases
{
    class PlaneBuilder
    {
        StringBuilder plane;

        
        public string BuildInitial(DataSet structure, JObject json)
        {

            try
            {
                plane = new StringBuilder("<Linea>000000100000001");
                DataRow[] structureDetail = structure.Tables[1].Select("desc_seccion = 'Inicial'");//consulta estructura de linea inicial

                //valida su se esta enviando variable o fija la compañia
                if (structureDetail[structureDetail.Length - 1]["Fuente"].ToString() == "")
                {
                    plane.Append(structureDetail[structureDetail.Length - 1]["ValorFijo"].ToString());
                }
                else
                {
                    plane.Append(((string)json[structureDetail[structureDetail.Length - 1]["Fuente"].ToString()]).PadLeft(3, '0'));

                }

            }
            catch (Exception)
            {

                throw;
            }
            plane.Append("</Linea>");
            return plane.ToString();
        }

        public string BuildMasters(DataSet structure, JObject json, ref int consectLine)
        {
            try
            {


                plane = new StringBuilder();
                DataRow[] structureDetail = structure.Tables[1].Select("desc_seccion = '" + json["Conector"].ToString() + "'");//extrae estructura del encabezado o maestro
                plane.AppendLine();
                consectLine++;
                plane.Append("<Linea>" + (consectLine).ToString().PadLeft(7, '0'));//genera consecutivo de linea

                for (int i = 1; i < structureDetail.Length; i++)//recorre la estructura armando la linea
                {
                    //valida que sea fijo o variable
                    if (structureDetail[i]["Fuente"].ToString() == "")
                    {
                        int length = (int)(structureDetail[i]["Tamano"]);


                        //valida que sea numerico o alfanumerico
                        if (structureDetail[i]["Tipo"].ToString() == "Alfanumerico")
                        {
                            plane.Append(structureDetail[i]["ValorFijo"].ToString().PadRight(length, ' '));
                        }
                        else
                        {
                            if (structureDetail[i]["Tipo"].ToString() == "Decimal")
                            {
                                int ent = int.Parse((structureDetail[i]["Enteros"]).ToString());
                                int dec = int.Parse((structureDetail[i]["Decimales"]).ToString());
                                string[] decimalVal = structureDetail[i]["ValorFijo"].ToString().Split('.');
                                if (decimalVal.Count() > 1)
                                    plane.Append($"{decimalVal[0].ToString().PadLeft(ent, '0')}.{decimalVal[1].ToString().PadRight(dec, '0')}");
                                else
                                    plane.Append($"{decimalVal[0].ToString().PadLeft(ent, '0')}.{("0").PadRight(dec, '0')}");
                            }
                            else
                            {
                                plane.Append(structureDetail[i]["ValorFijo"].ToString().PadLeft(length, '0'));
                            }

                        }

                    }
                    else
                    {
                        int length = (int)(structureDetail[i]["Tamano"]);//extrae tamaño del campo
                                                                         //valida que sea numerico o alfanumerico
                        if (structureDetail[i]["Tipo"].ToString() == "Alfanumerico")
                        {
                            var t = structureDetail[i]["Fuente"].ToString();
                            plane.Append(((string)json[structureDetail[i]["Fuente"].ToString()]).PadRight(length, ' '));
                        }
                        else
                        {
                            if (structureDetail[i]["Tipo"].ToString() == "Decimal")
                            {
                                int ent = int.Parse((structureDetail[i]["Enteros"]).ToString());
                                int dec = int.Parse((structureDetail[i]["Decimales"]).ToString());
                                string[] decimalVal = json[structureDetail[i]["Fuente"].ToString()].ToString().Split('.');
                                if (decimalVal.Count() > 1)
                                    plane.Append($"{decimalVal[0].ToString().PadLeft(ent, '0')}.{decimalVal[1].ToString().PadRight(dec, '0')}");
                                else
                                    plane.Append($"{decimalVal[0].ToString().PadLeft(ent, '0')}.{("0").PadRight(dec, '0')}");
                            }
                            else if (structureDetail[i]["Tipo"].ToString() == "DecimalConSigno")
                            {
                                int ent = int.Parse((structureDetail[i]["Enteros"]).ToString());
                                int dec = int.Parse((structureDetail[i]["Decimales"]).ToString());
                                string[] decimalVal = json[structureDetail[i]["Fuente"].ToString()].ToString().Split('.');
                                if (decimalVal.Count() > 1)
                                    plane.Append($"+{decimalVal[0].ToString().PadLeft(ent, '0')}.{decimalVal[1].ToString().PadRight(dec, '0')}");
                                else
                                    plane.Append($"+{decimalVal[0].ToString().PadLeft(ent, '0')}.{("0").PadRight(dec, '0')}");
                            }
                            else
                            {
                                plane.Append(((string)json[structureDetail[i]["Fuente"].ToString()]).PadLeft(length, '0'));
                            }
                        }

                    }
                    var a = plane.ToString();
                }

            }
            catch (Exception e)
            {
                throw e;
            }
            plane.Append("</Linea>");
            return plane.ToString();
        }

        public string BuildDetails(DataSet structure, JObject json, ref int consectLine)
        {

            try
            {
                plane = new StringBuilder();
                int sections = structure.Tables[3].Rows.Count;//extrae cantidad de secciones a recorrer de los detalles
                for (int id = 0; id < sections; id++)//recorre las secciones
                {
                    DataRow[] structureDetail = structure.Tables[2].Select("desc_seccion = '" + structure.Tables[3].Rows[id]["desc_seccion"].ToString() + "'");//extrae los detalle de la seccion en curso

                    var detailexist = json[structureDetail[0]["desc_seccion"].ToString()];//verifica que se esten enviando detalles en el json
                    if (detailexist != null)
                    {
                        int details = json[structureDetail[0]["desc_seccion"].ToString()].Count();//extrae la cantidad de detalle actual en el json

                        for (int t = 0; t < details; t++)//recorre los detalles enviados en el json
                        {
                            plane.AppendLine();
                            consectLine++;
                            plane.Append("<Linea>" + (consectLine).ToString().PadLeft(7, '0'));//asigna consecutivo de linea

                            for (int i = 1; i < structureDetail.Length; i++)//recorre estructura del detalle armando la linea
                            {

                                //valida que sea fijo o variable 
                                if (structureDetail[i]["Fuente"].ToString() == "")
                                {
                                    int length = (int)(structureDetail[i]["Tamano"]);
                                    //valida que sea numerico o alfanumerico
                                    if (structureDetail[i]["Tipo"].ToString() == "Alfanumerico")
                                    {
                                        plane.Append(structureDetail[i]["ValorFijo"].ToString().PadRight(length, ' '));
                                    }
                                    else
                                    {
                                        if (structureDetail[i]["Tipo"].ToString() == "Decimal")
                                        {

                                            int ent = int.Parse((structureDetail[i]["Enteros"]).ToString());
                                            int dec = int.Parse((structureDetail[i]["Decimales"]).ToString());
                                            string[] decimalVal = structureDetail[i]["ValorFijo"].ToString().Split('.');
                                            if (decimalVal.Count() > 1)
                                                plane.Append($"{decimalVal[0].ToString().PadLeft(ent, '0')}.{decimalVal[1].ToString().PadRight(dec, '0')}");
                                            else
                                                plane.Append($"{decimalVal[0].ToString().PadLeft(ent, '0')}.{("0").PadRight(dec, '0')}");
                                        }
                                        else
                                        {
                                            plane.Append(structureDetail[i]["ValorFijo"].ToString().PadLeft(length, '0'));
                                        }
                                    }
                                }
                                else
                                {
                                    int length = (int)(structureDetail[i]["Tamano"]);

                                    if (structureDetail[i]["Tipo"].ToString() == "Alfanumerico")
                                    {
                                        plane.Append(((string)json[structureDetail[i]["desc_seccion"].ToString()][t][structureDetail[i]["Fuente"].ToString()]).PadRight(length, ' '));
                                    }
                                    else
                                    {
                                        //valida que sea decimal
                                        if (structureDetail[i]["Tipo"].ToString() == "Decimal")
                                        {

                                            int ent = int.Parse((structureDetail[i]["Enteros"]).ToString());
                                            int dec = int.Parse((structureDetail[i]["Decimales"]).ToString());
                                            string[] decimalVal = json[structureDetail[i]["desc_seccion"].ToString()][t][structureDetail[i]["Fuente"].ToString()].ToString().Split('.');
                                            if (decimalVal.Count() > 1)
                                                plane.Append($"{decimalVal[0].ToString().PadLeft(ent, '0')}.{decimalVal[1].ToString().PadRight(dec, '0')}");
                                            else
                                                plane.Append($"{decimalVal[0].ToString().PadLeft(ent, '0')}.{("0").PadRight(dec, '0')}");

                                        }
                                        else if (structureDetail[i]["Tipo"].ToString() == "DecimalConSigno")
                                        {
                                            int ent = int.Parse((structureDetail[i]["Enteros"]).ToString());
                                            int dec = int.Parse((structureDetail[i]["Decimales"]).ToString());
                                            string[] decimalVal = json[structureDetail[i]["desc_seccion"].ToString()][t][structureDetail[i]["Fuente"].ToString()].ToString().Split('.');
                                            if (decimalVal.Count() > 1)
                                                plane.Append($"+{decimalVal[0].ToString().PadLeft(ent, '0')}.{decimalVal[1].ToString().PadRight(dec, '0')}");
                                            else
                                                plane.Append($"+{decimalVal[0].ToString().PadLeft(ent, '0')}.{("0").PadRight(dec, '0')}");
                                        }
                                        else
                                        {
                                            plane.Append(((string)json[structureDetail[i]["desc_seccion"].ToString()][t][structureDetail[i]["Fuente"].ToString()]).PadLeft(length, '0'));
                                        }
                                    }

                                }
                                var resultado = plane.ToString();
                            }
                            plane.Append("</Linea>");
                        }
                    }



                }
            }
            catch (Exception)
            {

                throw;
            }

            return plane.ToString();
        }

        public string BuildFinal(DataSet structure, JObject json, ref int consectLine)
        {

            try
            {
                plane = new StringBuilder();
                DataRow[] structureDetail = structure.Tables[1].Select("desc_seccion = 'Final'");
                plane.AppendLine();
                consectLine++;
                plane.Append("<Linea>" + (consectLine).ToString().PadLeft(7, '0'));

                for (int i = 0; i < structureDetail.Length; i++)
                {
                    if (structureDetail[i]["Fuente"].ToString() == "")
                    {
                        plane.Append(structureDetail[i]["ValorFijo"].ToString());
                    }
                    else
                    {
                        plane.Append(((string)json[structureDetail[i]["Fuente"].ToString()]).PadLeft(3, '0'));

                    }
                }
                plane.Append("</Linea>");

            }
            catch (Exception)
            {

                throw;
            }

            return plane.ToString();
        }

        //public string SendInformationWS(string xml)
        //{
        //    try
        //    {
        //        string xmlSend;
        //        wsSIESA.WSUNOEESoap ws = new wsSIESA.WSUNOEESoapClient(wsSIESA.WSUNOEESoapClient.EndpointConfiguration.WSUNOEESoap);
        //        wsSIESA.ImportarXMLRequest request = new wsSIESA.ImportarXMLRequest();


        //        xmlSend = "<Importar>";
        //        xmlSend += "<NombreConexion>unoee</NombreConexion>";
        //        xmlSend += "<IdCia>1</IdCia>";
        //        xmlSend += "<Usuario>auxcontable</Usuario>";
        //        xmlSend += "<Clave>1234</Clave>";
        //        xmlSend += "<Datos>";
        //        xmlSend += xml;
        //        xmlSend += "</Datos>";
        //        xmlSend += "</Importar>";

        //        request.printTipoError = 0;
        //        request.pvstrDatos = xmlSend;
        //        var result = ws.ImportarXMLAsync(request).Result;
        //        if (result.printTipoError == 0)
        //        {
        //            return "Importacion Exitosa";
        //        }
        //        else
        //        {
        //            return result.ImportarXMLResult.Nodes[1].ToString();
        //        }


        //    }
        //    catch (Exception ex)
        //    {

        //        return ex.Message;
        //    }



        //}
    }
}
