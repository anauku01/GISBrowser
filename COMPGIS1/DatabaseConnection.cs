using System;
using System.Data;
using System.Data.OleDb;

namespace COMPGIS1
{
    class DatabaseConnection
    {
        public void DC()
        {
            // The connection string assumes that the Access 
            // Northwind.mdb is located in the c:\Data folder.
            string connectionString = ("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=C:\\GISData\\Comp_BP_Data_Lasalle_v5(1).mdb; Jet OLEDB:System Database=C:\\GISData\\V08.mdw; User ID=developer; Password=waiheke");

            // Provide the query string with a parameter placeholder.
            string queryString =
                "SELECT Boring_Location_ID, Notes, Location_Description from Soil_Boring_Locations "
                    //+ "WHERE Boring_Location_ID > 12 "
                    + "ORDER BY Boring_Location_ID DESC;";

            // Specify the parameter value.
            int paramValue = 5;

            // Create and open the connection in a using block. This
            // ensures that all resources will be closed and disposed
            // when the code exits.
            using (OleDbConnection connection =
                new OleDbConnection(connectionString))
            {
                // Create the Command and Parameter objects.
                OleDbCommand command = new OleDbCommand(queryString, connection);
                command.Parameters.AddWithValue("@boringlocation", paramValue);

                // Open the connection in a try/catch block. 
                // Create and execute the DataReader, writing the result
                // set to the console window.
                try
                {
                    connection.Open();
                    OleDbDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        System.Diagnostics.Debug.WriteLine("\t{0}\t{1}\t{2}", reader[0], reader[1], reader[2]);
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                Console.ReadLine();
            }
        }
    }
}
