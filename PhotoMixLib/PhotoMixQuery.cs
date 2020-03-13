using System;
using System.Data;

//$ using Microsoft.Webstore.WstClient;
using System.Data.SqlClient;
using Msn.Framework;
//$ using Msn.Framework.Webstore;

namespace Msn.PhotoMix
{
    public class PhotoMixQuery : IDisposable
	{
        private SqlConnection connection = null;
        protected SqlCommand command = null;
        private SqlDataReader reader = null;
        private SqlDataAdapter dataAdapter = null;

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            Close();
        }

        public SqlParameterCollection Parameters
        {
            get { return command.Parameters; }
        }

        public PhotoMixQuery(string sql)
            :
            this(sql, CommandType.StoredProcedure)
        {
        }

        public PhotoMixQuery(string sql, CommandType commandType)
        {
            connection = new SqlConnection(GetConnectionString());
            command = new SqlCommand(sql, connection);
            command.CommandType = commandType;
        }

        public PhotoMixQuery(string sql, CommandType commandType, bool forDataSet)
        {
            connection = new SqlConnection(GetConnectionString());            
            command = new SqlCommand(sql, connection);
            command.CommandType = commandType;
            if (forDataSet)
                dataAdapter = new SqlDataAdapter(sql, connection);

        }
        
        public DataSet GetDataSet()
        {            
            // Bind the command to the data adapter
            dataAdapter.SelectCommand = command;
            
            // create the DataSet 
            DataSet dataSet = new DataSet();
            
            // fill the DataSet using our DataAdapter 
            dataAdapter.Fill(dataSet);

            return dataSet;
        }

        public SqlDataReader Reader
        {
            get
            {
                if (reader == null)
                {
                    connection.Open();
                    reader = command.ExecuteReader();
                }
                return reader;
            }
        }

		public int Execute()
		{
            int rows;
            try
            {
                connection.Open();
                rows = command.ExecuteNonQuery();
            }
            catch (Exception exception)
            {
                throw exception;
            }
            finally
            {
                Close();
            }
            return rows;
        }

        public void Close()
        {
            if (reader != null)
            {
                reader.Close();
                reader = null;
            }
            if (connection != null)
            {
                connection.Close();
                connection = null;
            }
        }

        private string GetConnectionString()
        {
			return Config.GetDefaultConnectionString();
        }
    }
}
