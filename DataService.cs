using System;
using System.Collections.Generic;
using System.Text;

using System.Data;
using System.Data.SqlClient;


namespace Excell
{
    class DataService
    {
        private static string dbConnectionString = @"Server=DESKTOP-LGFTI19\SQLEXPRESS; Database=KeyValues; Trusted_Connection=True;";
        
        internal static DataTable GetFields()
        {           
            SqlConnection connection = new SqlConnection(dbConnectionString);

            string sql = "SELECT * FROM Fields";
            SqlDataAdapter adapter = new SqlDataAdapter(sql, connection);            
            DataTable fieldsTable = new DataTable();
            adapter.Fill(fieldsTable);

            return fieldsTable;
        }

        internal  static void AddField(string UID, string fieldType, string fieldKey, string fieldValue)
        {

            using (SqlConnection connection = new SqlConnection(dbConnectionString))
            {
                String query = "INSERT INTO Fields (UID,ObjectType, FieldKey, FieldValue)" +
                               " VALUES (@UID, @ObjectType, @FieldKey, @FieldValue)";
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@UID", UID);
                    command.Parameters.AddWithValue("@ObjectType", fieldType);
                    command.Parameters.AddWithValue("@FieldKey", fieldKey);
                    command.Parameters.AddWithValue("@FieldValue", fieldValue);

                    connection.Open();
                    command.ExecuteNonQuery();

                }
            }

        }


    }
}
