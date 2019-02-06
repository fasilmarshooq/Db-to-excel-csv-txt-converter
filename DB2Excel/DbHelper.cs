using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace DB2Excel
{
    class DbHelper
    {
        
        public static DataTable ConnectToDB()
        {
            DataTable dataTable = new DataTable();
            SqlConnection Conn = new SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ToString());
            SqlCommand command = new SqlCommand(ConfigurationManager.AppSettings["SqlQuery"].ToString(), Conn);
            Conn.Open();
            SqlDataAdapter da = new SqlDataAdapter(command);
            da.Fill(dataTable);
            Conn.Close();
            da.Dispose();
            return dataTable;


        }


    }

}
