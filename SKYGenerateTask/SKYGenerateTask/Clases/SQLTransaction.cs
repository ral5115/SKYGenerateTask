using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SKYGenerateTask.Clases
{
    class SQLTransaction
    {
        public DataSet GetStruct()
        {            
            DataSet result = new DataSet();
            SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["Con"].ToString());
            connection.Open();
            SqlCommand command = new SqlCommand();
            SqlDataAdapter daData = new SqlDataAdapter();
            command.Connection = connection;
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "sp_select_struct";
            try
            {
                daData.SelectCommand = command;
                daData.Fill(result);
                result.Tables[0].TableName = "Struct";
                result.Tables[1].TableName = "StructMasters";
                result.Tables[2].TableName = "StructDetails";
            }
            catch (Exception e)
            {

                throw e;
            }
            finally
            {
                command.Connection.Close();
            }



            return result;
        }


    }
}
