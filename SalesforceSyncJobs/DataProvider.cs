using System;
using System.Collections.Specialized;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Diagnostics;
using System.Text;
using System.Xml;
using System.Configuration;
using System.Web;

namespace SalesforceSyncJobs
{
    /// <summary>
    /// Summary description for DataProvider.
    /// </summary>
    public class DataProvider : IDisposable
    {
        /// <summary>
        /// Class level SqlConnection
        /// </summary>
        private SqlConnection _MyConnection = null;

        /// <summary>
        /// Class level SqlTransaction
        /// </summary>
        private SqlTransaction _MyTransaction = null;

        /// <summary>
        /// Flag to indicate the DataProvider is currently in a SqlTransaction
        /// </summary>
        private bool _IsInTransaction = false;

        /// <summary>
        /// Connection string provided in the constructor.
        /// </summary>
        private string _ConnectionString = "";

        /// <summary>
        /// Track whether Dispose has been called.
        /// </summary>
        private bool Disposed = false;


        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="pConnectionKey">Connection string key.</param>
        public DataProvider(string pConnectionString)
        {
            try
            {
                _ConnectionString = pConnectionString;

                //Create and initialize the SqlConnection object
                _MyConnection = new SqlConnection(_ConnectionString);

            }
            catch (Exception e)
            {
                throw new Exception("DataProvider failed in constructor." + e.ToString(), e);
            }
            finally
            {
            }
        }

        /// <summary>
        /// Connection string attribute
        /// </summary>
        public string ConnectionString
        {
            get { return _ConnectionString; }
            set
            {
                _ConnectionString = value;
                if (_MyConnection != null) _MyConnection.ConnectionString = _ConnectionString;
            }
        }

        /// <summary>
        /// Destructor.
        /// </summary>
        /// <remarks>
        /// Use C# destructor syntax for finalization code.
        /// This destructor will run only if the Dispose method 
        /// does not get called. 
        /// It gives your base class the opportunity to finalize.
        /// Do not provide destructors in types derived from this class.
        /// </remarks>
        ~DataProvider()
        {
            // Do not re-create Dispose clean-up code here.
            // Calling Dispose(false) is optimal in terms of
            // readability and maintainability.
            Dispose(false);
        }

        /// <summary>
        /// Implement IDisposable. 
        /// </summary>
        /// <remarks>
        /// Do not make this method virtual.
        /// A derived class should not be able to override this method. 
        /// </remarks>
        public void Dispose()
        {
            Dispose(true);
            // This object will be cleaned up by the Dispose method.
            // Therefore, you should call GC.SupressFinalize to
            // take this object off the finalization queue 
            // and prevent finalization code for this object
            // from executing a second time.
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Internal dispose method.
        /// </summary>
        /// <param name="Disposing">Flag to indicate component is being disposed.</param>
        /// <remarks>
        /// Dispose(bool disposing) executes in two distinct scenarios.
        /// If disposing equals true, the method has been called directly
        /// or indirectly by a user's code. Managed and unmanaged resources
        /// can be disposed.
        /// If disposing equals false, the method has been called by the 
        /// runtime from inside the finalizer and you should not reference 
        /// other objects. Only unmanaged resources can be disposed.
        /// </remarks>
        private void Dispose(bool Disposing)
        {
            // Check to see if Dispose has already been called.
            if (!this.Disposed)
            {
                // If disposing equals true, dispose all managed 
                // and unmanaged resources.
                if (Disposing)
                {
                    if (_IsInTransaction)
                    {
                        //If the transaction is still active when the Dispose method is called, Rollback must be called.
                        //The state of the transaction isn't guaranteed to be correct at this point
                        TransRollback();
                    }

                    if (_MyConnection != null)
                    {
                        //Close the connection if not already closed
                        if (_MyConnection.State != ConnectionState.Closed)
                        {
                            _MyConnection.Close();
                        }
                        //Dispose the connection
                        _MyConnection.Dispose();
                        _MyConnection = null;
                    }
                }

                // Call the appropriate methods to clean up unmanaged resources here.
            }
            Disposed = true;
        }

        /// <summary>
        /// This function executes a non-query command against the database.
        /// </summary>
        /// <param name="pCommand">Command to execute.</param>
        /// <param name="pCommandType">System.Data.CommandType indicating the type of command.</param>
        /// <param name="pInputs">A NameValueCollection containing the name of the parameter and value of the parameter. If no parameters are required pass null for this parameter.</param>
        /// <returns>Boolean. True for success, False for failure.</returns>
        public bool DoNonQuery(string pCommand, CommandType pCommandType, NameValueCollection pInputs)
        {
            SqlCommand MyCmd = null;
            bool pResult = false;
            try
            {
                //Create a new command object
                MyCmd = _MyConnection.CreateCommand();
                MyCmd.CommandType = pCommandType;
                MyCmd.CommandText = pCommand;

                // If this is part of a transaction then connect it here to this command.
                if (_IsInTransaction)
                {
                    MyCmd.Transaction = _MyTransaction;
                }

                //Assure the SqlConnection is open
                OpenConnection();

                //Set parameters
                SetParameters(ref MyCmd, pInputs);

                //Execute the query
                MyCmd.ExecuteNonQuery();

                //It didn't blow up so return True.
                pResult = true;
            }
            catch (Exception e)
            {
                throw new Exception("DoNonQuery failed in DataProvider." + e.ToString(), e);
            }
            finally
            {
                //Close the SqlConnection
                CloseConnection();
                MyCmd = null;
            }
            return pResult;
        }

        private int GetTimeoutSetting()
        {
            try
            {
                int iTimeout = 2000;
                //int.TryParse(ConfigurationManager.AppSettings["SQL_QUERY_TIMEOUT"], out iTimeout);
                return iTimeout;
            }
            catch (Exception ex)
            {
                throw new Exception("Unable to get timeout setting. " + ex.Message);
            }
        }
        /// <summary>
        /// This function executes a non-query command against the database.
        /// </summary>
        /// <param name="pCommand">Command to execute.</param>
        /// <param name="pCommandType">System.Data.CommandType indicating the type of command.</param>
        /// <param name="pInputs">A NameValueCollection containing the name of the parameter and value of the parameter. If no parameters are required pass null for this parameter.</param>
        /// <returns>Boolean. True for success, False for failure.</returns>
        public bool DoNonQuery(string pCommand, CommandType pCommandType, NameValueCollection pInputs, NameValueCollection pInputTypes)
        {
            SqlCommand MyCmd = null;
            bool pResult = false;
            try
            {
                //Create a new command object
                MyCmd = _MyConnection.CreateCommand();
                MyCmd.CommandType = pCommandType;
                MyCmd.CommandText = pCommand;
                MyCmd.CommandTimeout = GetTimeoutSetting();

                // If this is part of a transaction then connect it here to this command.
                if (_IsInTransaction)
                {
                    MyCmd.Transaction = _MyTransaction;
                }

                //Assure the SqlConnection is open
                OpenConnection();

                //Set parameters
                SetParameters(ref MyCmd, pInputs, pInputTypes);

                //Execute the query
                MyCmd.ExecuteNonQuery();

                //It didn't blow up so return True.
                pResult = true;
            }

            catch (Exception e)
            {
                throw new Exception(MyCmd.CommandText + " DoNonQuery failed in DataProvider." + e.ToString(), e);
            }
            finally
            {
                //Close the SqlConnection
                CloseConnection();
                MyCmd = null;
            }
            return pResult;
        }

        // This function is the same as the DoNonQuery function but instead it returns an output parameter. This parameter is specified
        // as the fourth argument. The Stored Proc must declare an output parameter of the same name. This does not return the 
        // "RETURN VALUE" of a stored proc.
        public object DoNonQueryGetInputOutputParam(string pCommand, CommandType pCommandType, NameValueCollection pInputs, string RetValId, int size)
        {
            SqlCommand MyCmd = null;
            //bool pResult = false;
            object output = null;
            try
            {
                //Create a new command object
                MyCmd = _MyConnection.CreateCommand();
                MyCmd.CommandType = pCommandType;
                MyCmd.CommandText = pCommand;

                // If this is part of a transaction then connect it here to this command.
                if (_IsInTransaction)
                {
                    MyCmd.Transaction = _MyTransaction;
                }

                //Assure the SqlConnection is open
                OpenConnection();

                //Set parameters
                SetParameters(ref MyCmd, pInputs);
                MyCmd.Parameters["@" + RetValId].Direction = ParameterDirection.InputOutput;
                if (size > 0)
                    MyCmd.Parameters["@" + RetValId].Size = size;
                //Execute the query
                MyCmd.ExecuteNonQuery();
                MyCmd.UpdatedRowSource = UpdateRowSource.OutputParameters;
                output = MyCmd.Parameters["@" + RetValId].Value;
            }
            catch (Exception e)
            {
                throw new Exception(MyCmd.CommandText + " DoNonQuery failed in DataProvider." + e.ToString(), e);
            }
            finally
            {
                //Close the SqlConnection
                CloseConnection();
                MyCmd = null;
            }
            return output;
        }

        // This function is the same as the DoNonQuery function but instead it returns an output parameter. This parameter is specified
        // as the fourth argument. The Stored Proc must declare an output parameter of the same name. This does not return the 
        // "RETURN VALUE" of a stored proc.
        public object DoNonQueryGetOutputParam(string pCommand, CommandType pCommandType, NameValueCollection pInputs, string RetValId, int size)
        {
            SqlCommand MyCmd = null;
            //bool pResult = false;
            object output = null;
            try
            {
                //Create a new command object
                MyCmd = _MyConnection.CreateCommand();
                MyCmd.CommandType = pCommandType;
                MyCmd.CommandText = pCommand;

                // If this is part of a transaction then connect it here to this command.
                if (_IsInTransaction)
                {
                    MyCmd.Transaction = _MyTransaction;
                }

                //Assure the SqlConnection is open
                OpenConnection();

                //Set parameters
                SetParameters(ref MyCmd, pInputs);
                MyCmd.Parameters["@" + RetValId].Direction = ParameterDirection.Output;
                if (size > 0)
                    MyCmd.Parameters["@" + RetValId].Size = size;
                //Execute the query
                MyCmd.ExecuteNonQuery();
                MyCmd.UpdatedRowSource = UpdateRowSource.OutputParameters;
                output = MyCmd.Parameters["@" + RetValId].Value;
            }
            catch (Exception e)
            {
                throw new Exception(MyCmd.CommandText + " DoNonQuery failed in DataProvider." + e.ToString(), e);
            }
            finally
            {
                //Close the SqlConnection
                CloseConnection();
                MyCmd = null;
            }
            return output;
        }

        // This function is the same as the DoNonQuery function but instead it returns an output parameter. This parameter is specified
        // as the fourth argument. The Stored Proc must declare an output parameter of the same name. This does not return the 
        // "RETURN VALUE" of a stored proc.
        public NameValueCollection DoNonQueryGetMultipleOutputParam(string pCommand, CommandType pCommandType, NameValueCollection pInputs, NameValueCollection pOutputs)
        {
            SqlCommand MyCmd = null;
            //bool pResult = false;
            NameValueCollection output = new NameValueCollection();
            try
            {
                //Create a new command object
                MyCmd = _MyConnection.CreateCommand();
                MyCmd.CommandType = pCommandType;
                MyCmd.CommandText = pCommand;

                // If this is part of a transaction then connect it here to this command.
                if (_IsInTransaction)
                {
                    MyCmd.Transaction = _MyTransaction;
                }

                //Assure the SqlConnection is open
                OpenConnection();

                //Set parameters
                SetParameters(ref MyCmd, pInputs);

                foreach (string pKey in pOutputs.Keys)
                {
                    MyCmd.Parameters[pKey].Direction = ParameterDirection.Output;
                    MyCmd.Parameters[pKey].Size = Convert.ToInt32(pOutputs[pKey]);
                }

                //Execute the query
                MyCmd.ExecuteNonQuery();
                MyCmd.UpdatedRowSource = UpdateRowSource.OutputParameters;

                foreach (string pKey in pOutputs.Keys)
                {
                    output.Add(pKey, MyCmd.Parameters[pKey].Value.ToString());
                }
            }
            catch (Exception e)
            {
                throw new Exception(MyCmd.CommandText + " DoNonQuery failed in DataProvider." + e.ToString(), e);
            }
            finally
            {
                //Close the SqlConnection
                CloseConnection();
                MyCmd = null;
            }
            return output;
        }
        /// <summary>
        /// This function executes a comand against the database and returns the resulting DataReader.
        /// </summary>
        /// <param name="pCommand">Command to execute.</param>
        /// <param name="pCommandType">System.Data.CommandType indicating the type of command.</param>
        /// <param name="pInputs">A NameValueCollection containing the name of the parameter and value of the parameter. If no parameters are required pass null for this parameter.</param>
        /// <returns>System.Data.SqlClient.SqlDataReader containing the results of the command.</returns>
        public SqlDataReader GetDataReader(string pCommand, CommandType pCommandType, NameValueCollection pInputs)
        {
            SqlCommand MyCmd = null;
            SqlDataReader MyDr = null;
            try
            {
                //Create a new command object
                MyCmd = _MyConnection.CreateCommand();
                MyCmd.CommandType = pCommandType;
                MyCmd.CommandText = pCommand;

                // If this is part of a transaction then connect it here to this command.
                if (_IsInTransaction)
                {
                    MyCmd.Transaction = _MyTransaction;
                }

                //Assure the SqlConnection is open
                OpenConnection();

                //Set parameters
                SetParameters(ref MyCmd, pInputs);

                //Execute the DataReader
                //NOTE: CommandBehavior.CloseConnection will close the SqlConnection when 
                //the DataReader is consumed by the calling code.

                MyDr = MyCmd.ExecuteReader(CommandBehavior.CloseConnection);
                return MyDr;
            }
            catch (Exception e)
            {
                throw new Exception(MyCmd.CommandText + " GetDataReader failed in DataProvider.", e);
            }
        }

        /// <summary>
        /// This function executes a command against the database and returns the resulting DataSet.
        /// </summary>
        /// <param name="pCommand">Command to execute.</param>
        /// <param name="pCommandType">System.Data.CommandType indicating the type of command.</param>
        /// <param name="pInputs">A NameValueCollection containing the name of the parameter and value of the parameter. If no parameters are required pass null for this parameter.</param>
        /// <returns>System.Data.DataSet containing the results of the command.</returns>
        public DataSet GetDataSet(string pCommand, CommandType pCommandType, NameValueCollection pInputs)
        {
            SqlCommand MyCmd = null;
            SqlDataAdapter MyAdapter = null;
            DataSet MyDs = new DataSet();
            try
            {
                //Create a new command object
                MyCmd = _MyConnection.CreateCommand();
                MyCmd.CommandType = pCommandType;
                MyCmd.CommandText = pCommand;
                MyCmd.CommandTimeout = GetTimeoutSetting();

                // If this is part of a transaction then connect it here to this command.
                if (_IsInTransaction)
                {
                    MyCmd.Transaction = _MyTransaction;
                }

                //Assure the SqlConnection is open
                OpenConnection();

                //Set parameters
                SetParameters(ref MyCmd, pInputs);

                //Use a DataAdapter to capture the data into a DataSet.
                MyAdapter = new SqlDataAdapter(MyCmd);

                //Fill the DataSet using the DataAdapter

                MyAdapter.Fill(MyDs);
                return MyDs;
            }
            catch (Exception e)
            {
                throw new Exception(MyCmd.CommandText + " GetDataSet failed in DataProvider.", e);
            }
            finally
            {
                //Assure the SqlConnection is closed
                CloseConnection();

                if (MyAdapter != null)
                {
                    MyAdapter.Dispose();
                }

                MyCmd = null;
                MyAdapter = null;
            }
        }

        /// <summary>
        /// This function executes a command against the database and returns the resulting DataSet.
        /// </summary>
        /// <param name="pCommand">Command to execute.</param>
        /// <param name="pCommandType">System.Data.CommandType indicating the type of command.</param>
        /// <param name="pInputs">A NameValueCollection containing the name of the parameter and value of the parameter. If no parameters are required pass null for this parameter.</param>
        /// <returns>System.Data.DataSet containing the results of the command.</returns>
        public DataSet GetDataSetOutputParam(string pCommand, CommandType pCommandType, NameValueCollection pInputs, string RetValId, int size, out object value)
        {
            SqlCommand MyCmd = null;
            SqlDataAdapter MyAdapter = null;
            DataSet MyDs = new DataSet();
            value = "";
            try
            {
                //Create a new command object
                MyCmd = _MyConnection.CreateCommand();
                MyCmd.CommandType = pCommandType;
                MyCmd.CommandText = pCommand;
                MyCmd.CommandTimeout = GetTimeoutSetting();

                // If this is part of a transaction then connect it here to this command.
                if (_IsInTransaction)
                {
                    MyCmd.Transaction = _MyTransaction;
                }

                //Assure the SqlConnection is open
                OpenConnection();

                //Set parameters
                SetParameters(ref MyCmd, pInputs);
                MyCmd.Parameters["@" + RetValId].Direction = ParameterDirection.Output;
                if (size > 0)
                    MyCmd.Parameters["@" + RetValId].Size = size;

                //Use a DataAdapter to capture the data into a DataSet.
                MyAdapter = new SqlDataAdapter(MyCmd);

                //Fill the DataSet using the DataAdapter

                MyAdapter.Fill(MyDs);

                MyCmd.UpdatedRowSource = UpdateRowSource.OutputParameters;
                value = MyCmd.Parameters["@" + RetValId].Value;

                return MyDs;
            }
            catch (Exception e)
            {
                throw new Exception(MyCmd.CommandText + " GetDataSet failed in DataProvider.", e);
            }
            finally
            {
                //Assure the SqlConnection is closed
                CloseConnection();

                if (MyAdapter != null)
                {
                    MyAdapter.Dispose();
                }

                MyCmd = null;
                MyAdapter = null;
            }
        }

        /// <summary>
        /// This function executes a command against the database and returns the resulting DataSet.
        /// </summary>
        /// <param name="pCommand">Command to execute.</param>
        /// <param name="pCommandType">System.Data.CommandType indicating the type of command.</param>
        /// <param name="pInputs">A NameValueCollection containing the name of the parameter and value of the parameter. If no parameters are required pass null for this parameter.</param>
        /// <param name="pOutputs">A NameValueCollection containing the name of the parameter and value of the parameter. If no parameters are required pass null for this parameter.</param>
        /// <returns>System.Data.DataSet containing the results of the command.</returns>
        public DataSet GetDataSetMultipleOutputParam(string pCommand, CommandType pCommandType, NameValueCollection pInputs, NameValueCollection pOutputs, out NameValueCollection result)
        {
            SqlCommand MyCmd = null;
            SqlDataAdapter MyAdapter = null;
            DataSet MyDs = new DataSet();
            result = new NameValueCollection();

            try
            {
                //Create a new command object
                MyCmd = _MyConnection.CreateCommand();
                MyCmd.CommandType = pCommandType;
                MyCmd.CommandText = pCommand;
                MyCmd.CommandTimeout = GetTimeoutSetting();

                // If this is part of a transaction then connect it here to this command.
                if (_IsInTransaction)
                {
                    MyCmd.Transaction = _MyTransaction;
                }

                //Assure the SqlConnection is open
                OpenConnection();

                //Set parameters
                SetParameters(ref MyCmd, pInputs);

                foreach (string pKey in pOutputs.Keys)
                {
                    MyCmd.Parameters[pKey].Direction = ParameterDirection.Output;
                    MyCmd.Parameters[pKey].Size = Convert.ToInt32(pOutputs[pKey]);
                }

                //Use a DataAdapter to capture the data into a DataSet.
                MyAdapter = new SqlDataAdapter(MyCmd);

                //Fill the DataSet using the DataAdapter

                MyAdapter.Fill(MyDs);

                MyCmd.UpdatedRowSource = UpdateRowSource.OutputParameters;

                foreach (string pKey in pOutputs.Keys)
                {
                    result.Add(pKey, MyCmd.Parameters[pKey].Value.ToString());
                }

                return MyDs;
            }
            catch (Exception e)
            {
                throw new Exception(MyCmd.CommandText + " GetDataSet failed in DataProvider.", e);
            }
            finally
            {
                //Assure the SqlConnection is closed
                CloseConnection();

                if (MyAdapter != null)
                {
                    MyAdapter.Dispose();
                }

                MyCmd = null;
                MyAdapter = null;
            }
        }

        /// <summary>
        /// This function executes a command against the database and returns the value in 1st row & 1st column of the resultset.
        /// </summary>
        /// <param name="pCommand">Command to execute.</param>
        /// <param name="pCommandType">System.Data.CommandType indicating the type of command.</param>
        /// <param name="pInputs">A NameValueCollection containing the name of the parameter and value of the parameter. If no parameters are required pass null for this parameter.</param>
        /// <returns>Object representing the return value of the command.</returns>
        public object GetScalar(string pCommand, CommandType pCommandType, NameValueCollection pInputs)
        {
            SqlCommand MyCmd = null;
            object pResult = null;
            try
            {
                //Create a new command object
                MyCmd = _MyConnection.CreateCommand();
                MyCmd.CommandType = pCommandType;
                MyCmd.CommandText = pCommand;

                // If this is part of a transaction then connect it here to this command.
                if (_IsInTransaction)
                {
                    MyCmd.Transaction = _MyTransaction;
                }

                //Assure the SqlConnection is open
                OpenConnection();

                //Set parameters
                SetParameters(ref MyCmd, pInputs);

                //Capture the results of ExecuteScalar() into an object.
                pResult = MyCmd.ExecuteScalar();
                // fix added because if we return a nulll and are assigning to a variable, an exception will be thrown.
                // Now, we return a blank, which should not cause an exception to be thrown.
                if (pResult == null)
                    pResult = "";
            }
            catch (Exception e)
            {
                throw new Exception(MyCmd.CommandText + " GetScalar failed in DataProvider.", e);
            }
            finally
            {
                //Close the SqlConnection
                CloseConnection();
                MyCmd = null;
            }
            return pResult;
        }

        /// <summary>
        /// This function executes a command against the database and returns the resulting XML.
        /// </summary>
        /// <param name="pCommand">Command to execute.</param>
        /// <param name="pCommandType">System.Data.CommandType indicating the type of command.</param>
        /// <param name="pInputs">A NameValueCollection containing the name of the parameter and value of the parameter. If no parameters are required pass null for this parameter.</param>
        /// <returns>An XML string of data.</returns>
        public string GetXML(string pCommand, CommandType pCommandType, NameValueCollection pInputs)
        {
            return GetXML(pCommand, pCommandType, pInputs, "DOCUMENT");
        }

        /// <summary>
        /// This function executes a command against the database and returns the resulting XML with the specified root tag.
        /// </summary>
        /// <param name="pCommand">Command to execute.</param>
        /// <param name="pCommandType">System.Data.CommandType indicating the type of command.</param>
        /// <param name="pInputs">A NameValueCollection containing the name of the parameter and value of the parameter. If no parameters are required pass null for this parameter.</param>
        /// <param name="pRootTag">Value of the opening and closing root tags.</param>
        /// <returns>An XML string of data.</returns>
        public string GetXML(string pCommand, CommandType pCommandType, NameValueCollection pInputs, string pRootTag)
        {
            SqlCommand MyCmd = null;
            XmlReader MyXmlReader = null;
            string MyXml = "";
            try
            {
                //Remove any xml formatting specified.
                pRootTag = pRootTag.Replace("/", "");
                pRootTag = pRootTag.Replace("<", "");
                pRootTag = pRootTag.Replace(">", "");

                //Create a new command object
                MyCmd = _MyConnection.CreateCommand();
                MyCmd.CommandType = pCommandType;
                MyCmd.CommandText = pCommand;
                MyCmd.CommandTimeout = GetTimeoutSetting();

                // If this is part of a transaction then connect it here to this command.
                if (_IsInTransaction)
                {
                    MyCmd.Transaction = _MyTransaction;
                }

                //Assure the SqlConnection is open
                OpenConnection();

                //Set parameters
                SetParameters(ref MyCmd, pInputs);

                MyXmlReader = MyCmd.ExecuteXmlReader();

                if (MyXmlReader.Read())
                {
                    if (pRootTag == "")
                    {
                        while (MyXmlReader.ReadState != System.Xml.ReadState.EndOfFile)
                        {
                            MyXml = MyXmlReader.ReadOuterXml();
                        }
                    }
                    else
                    {
                        //Read the XML string
                        MyXml = MyXmlReader.ReadInnerXml();
                    }

                    if (MyXml != "")
                    {
                        if (pRootTag != "")
                            //Format the xml document with the root node tags
                            MyXml = String.Format("<{0}>{1}</{2}>", pRootTag, MyXml, pRootTag);
                    }
                }

                MyXmlReader.Close();
                return MyXml;
            }
            catch (Exception e)
            {
                throw new Exception(MyCmd.CommandText + " GetXML failed in DataProvider." + e.Message, e);
            }
            finally
            {
                //Close the SqlConnection
                CloseConnection();

                if (MyXmlReader.ReadState != ReadState.Closed)
                {
                    MyXmlReader.Close();
                }

                MyXmlReader = null;
                MyCmd = null;
            }
        }

        #region Transaction handling routines
        /// <summary>
        /// This routine will start a new transaction.
        /// </summary>
        /// <remarks>The SqlConnection will not be closed while the transaction is active.</remarks>
        public void TransBegin()
        {
            try
            {
                //Assure the SqlConnection is open
                OpenConnection();

                //Start a new SqlTransaction
                _MyTransaction = _MyConnection.BeginTransaction();

                //Flag the transaction active
                _IsInTransaction = true;
            }
            catch (Exception e)
            {
                throw new Exception("TransBegin failed in DataProvider." + e.ToString(), e);
            }
        }

        /// <summary>
        /// This routine will commit all changes since the TransBegin() call.
        /// </summary>
        public void TransCommit()
        {
            try
            {
                if (_IsInTransaction)
                {
                    //Commit all changes while in the transaction
                    _MyTransaction.Commit();
                }
            }
            catch (Exception e)
            {
                throw new Exception("TransCommit failed in DataProvider." + e.ToString(), e);
            }
            finally
            {
                if (_MyTransaction != null)
                {
                    _MyTransaction.Dispose();
                    _MyTransaction = null;
                }

                //Flag the transaction as inactive
                _IsInTransaction = false;
                CloseConnection();
            }
        }

        /// <summary>
        /// This routine will rollback anything executed since the TransBegin() call.
        /// </summary>
        public void TransRollback()
        {
            try
            {
                if (_IsInTransaction)
                {
                    //Rollback anything executed within the transaction
                    _MyTransaction.Rollback();
                }
            }
            catch (Exception e)
            {
                throw new Exception("TransRollback failed in DataProvider." + e.ToString(), e);
            }
            finally
            {
                if (_MyTransaction != null)
                {
                    _MyTransaction.Dispose();
                    _MyTransaction = null;
                }

                //Flag the transaction as inactive
                _IsInTransaction = false;
                CloseConnection();
            }
        }
        #endregion

        #region Connection handling routines
        /// <summary>
        /// Helper routine to open the SqlConnection for this instance of the DataProvider.
        /// </summary>
        private void OpenConnection()
        {
            try
            {
                if (_MyConnection != null)
                {
                    if (_MyConnection.State != ConnectionState.Open)
                    {
                        //Attempt to reopen the connection
                        _MyConnection.Open();
                    }
                }
                else
                {
                    //Create a new connection using the connection string passed to this instance of the DataProvider
                    _MyConnection = new SqlConnection(_ConnectionString);
                    _MyConnection.Open();
                }

            }
            catch (Exception ex)
            {
                throw new Exception("OpenConnection failed in DataProvider." + ex.ToString(), ex);
            }
        }

        /// <summary>
        /// Helper routine to close the SqlConnection for this instance of the DataProvider.
        /// </summary>
        private void CloseConnection()
        {
            try
            {
                //If a transaction is currently active the connection cannot be closed.
                if (!_IsInTransaction)
                {
                    if (_MyConnection != null)
                    {
                        if (_MyConnection.State != ConnectionState.Closed)
                        {
                            //Close the connection
                            _MyConnection.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("CloseConnection failed in DataProvider." + ex.ToString(), ex);
            }
        }
        #endregion

        /// <summary>
        /// Helper routine to populate the parameters for the given SqlCommand.
        /// </summary>
        /// <paramref name="pCmd">SqlCommand</paramref>
        /// <param name="pInputs">NameValueCollection containing the parameters provided by the business tier</param>
        private void SetParameters(ref SqlCommand pCmd, NameValueCollection pInputs)
        {
            string sKey = "";
            try
            {
                SqlCommandBuilder.DeriveParameters(pCmd);
                if (pInputs != null && pInputs.Count > 0)
                {
                    // If this is part of a transaction then we can't use DeriveParameters.
                    if (_IsInTransaction)
                    {
                        pCmd.Parameters.Clear();
                        foreach (string pKey in pInputs.Keys)
                        {
                            //Populate the parameter
                            sKey = pKey;
                            //pCmd.Parameters.Add(pKey, pInputs[pKey]); deprecated in .net 2.0 - 6/30/2008 btomlin
                            pCmd.Parameters.AddWithValue(pKey, pInputs[pKey]);
                        }
                    }
                    else
                    {
                        //Derive the parameters for this command
                        //SqlCommandBuilder.DeriveParameters(pCmd);
                        foreach (string pKey in pInputs.Keys)
                        {
                            sKey = pKey;		//for debugging
                            //Populate the parameter
                            switch (pCmd.Parameters[pKey].SqlDbType)
                            {
                                case SqlDbType.Bit:
                                    pCmd.Parameters[pKey].Value = Convert.ToBoolean(pInputs[pKey]);
                                    break;
                                case SqlDbType.Int:
                                    pCmd.Parameters[pKey].Value = Convert.ToInt32(pInputs[pKey]);
                                    break;
                                case SqlDbType.BigInt:
                                    pCmd.Parameters[pKey].Value = Convert.ToInt64(pInputs[pKey]);
                                    break;
                                case SqlDbType.SmallInt:
                                    pCmd.Parameters[pKey].Value = Convert.ToInt16(pInputs[pKey]);
                                    break;
                                case SqlDbType.Money:
                                    pCmd.Parameters[pKey].Value = Convert.ToDecimal(pInputs[pKey]);
                                    break;
                                case SqlDbType.SmallMoney:
                                    pCmd.Parameters[pKey].Value = Convert.ToDecimal(pInputs[pKey]);
                                    break;
                                case SqlDbType.Decimal:
                                    pCmd.Parameters[pKey].Value = Convert.ToDecimal(pInputs[pKey]);
                                    break;
                                case SqlDbType.UniqueIdentifier:
                                    Guid guid = new Guid(pInputs[pKey]);
                                    pCmd.Parameters[pKey].Value = guid;
                                    break;
                                default:
                                    pCmd.Parameters[pKey].Value = pInputs[pKey];
                                    break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                try
                {
                    throw new Exception(pCmd.CommandText + " SetParameters failed to set all parameters for the provided SqlCommand. Key=" + sKey, ex);
                }
                catch
                {

                }
            }
        }

        /// <summary>
        /// Helper routine to populate the parameters for the given SqlCommand.
        /// </summary>
        /// <paramref name="pCmd">SqlCommand</paramref>
        /// <param name="pInputs">NameValueCollection containing the parameters provided by the business tier</param>
        private void SetParameters(ref SqlCommand pCmd, NameValueCollection pInputs, NameValueCollection pInputTypes)
        {
            string sKey = "";
            try
            {
                if (pInputs != null)
                {
                    // If this is part of a transaction then we can't use DeriveParameters.
                    if (_IsInTransaction)
                    {
                        pCmd.Parameters.Clear();
                        foreach (string pKey in pInputs.Keys)
                        {
                            sKey = pKey;		//for debugging
                            //Populate the parameter
                            //pCmd.Parameters.Add(pKey, pInputs[pKey]); deprecated in .net 2.0 - 6/30/2008 btomlin
                            pCmd.Parameters.AddWithValue(pKey, pInputs[pKey]);
                        }
                    }
                    else
                    {
                        //Derive the parameters for this command
                        SqlCommandBuilder.DeriveParameters(pCmd);
                        foreach (string pKey in pInputs.Keys)
                        {
                            sKey = pKey;		//for debugging
                            //Populate the parameter
                            switch (pCmd.Parameters[pKey].SqlDbType)
                            {
                                case SqlDbType.Bit:
                                    pCmd.Parameters[pKey].Value = Convert.ToBoolean(pInputs[pKey]);
                                    break;
                                case SqlDbType.Int:
                                    pCmd.Parameters[pKey].Value = Convert.ToInt32(pInputs[pKey]);
                                    break;
                                case SqlDbType.BigInt:
                                    pCmd.Parameters[pKey].Value = Convert.ToInt64(pInputs[pKey]);
                                    break;
                                case SqlDbType.SmallInt:
                                    pCmd.Parameters[pKey].Value = Convert.ToInt16(pInputs[pKey]);
                                    break;
                                case SqlDbType.Money:
                                    pCmd.Parameters[pKey].Value = Convert.ToDecimal(pInputs[pKey]);
                                    break;
                                case SqlDbType.SmallMoney:
                                    pCmd.Parameters[pKey].Value = Convert.ToDecimal(pInputs[pKey]);
                                    break;
                                case SqlDbType.Decimal:
                                    pCmd.Parameters[pKey].Value = Convert.ToDecimal(pInputs[pKey]);
                                    break;
                                default:
                                    if (pInputTypes[pKey] == "Text")
                                    {
                                        pCmd.Parameters[pKey].SqlDbType = SqlDbType.Text;
                                        pCmd.Parameters[pKey].Size = pInputs[pKey].Length;
                                    }
                                    pCmd.Parameters[pKey].Value = pInputs[pKey];
                                    break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(pCmd.CommandText + " SetParameters failed to set all parameters for the provided SqlCommand. Key=" + sKey, ex);
            }
        }
    }
}

