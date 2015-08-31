using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Data.OleDb;

namespace Gewatec.Database.Odbc
{
    /// <summary>
    /// Summary description for ODBCDSN.
    /// </summary>
    public class OdbcDsn
    {
        internal OdbcDsn()
        {
        }

        internal Dictionary<string, string> Data { get; private set; }


        public static OdbcDsn ParseForODBCDSN(string dsnName, string dsnDriverName,
            Dictionary<string, string> data)
        {
            OdbcDsn odbcdsn = new OdbcDsn()
            {
                DSNName = dsnName,
                DSNDriverName = dsnDriverName,
                Data = data
            };

            // For each element defined for a typical DSN get
            // its value.
            foreach (var dsnElement in data)
            {
                string v = dsnElement.Value;

                switch (dsnElement.Key.ToLower())
                {
                    case "description":
                        odbcdsn.Description = v;
                        break;
                    case "server":
                    case "servername":
                        odbcdsn.ServerName = v;
                        break;
                    case "driver":
                        odbcdsn.DSNDriverPath = v;
                        break;
                    case "uid":
                        odbcdsn.UID = v;
                        break;
                    case "pwd":
                        odbcdsn.Password = v;
                        break;
                    case "databasename":
                        odbcdsn.DatabaseName = v;
                        break;
                    case "commlinks":
                        odbcdsn.CommLinks = v;
                        break;
                    case "host":
                        odbcdsn.Host = v;
                        break;
                }
            }

            return odbcdsn;
        }

        public string DSNName { get; private set; }
        public string DSNDriverName { get; private set; }
        public string DSNDriverPath { get; private set; }
        public string Description { get; private set; }
        public string ServerName { get; private set; }
        public string DatabaseName { get; private set; }
        public string UID { get; private set; }   // Username
        public string Password { get; private set; }
        public string CommLinks { get; private set; }
        public string Host { get; private set; }
        public bool UserDSN { get; internal set; }

        public override string ToString()
        {
            return DSNName;
        }



        private static Dictionary<string, string> String2Dict(string s)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();

            foreach (string element in s.Split(';'))
            {
                string[] parts = element.Split('=');
                result.Add(parts[0], parts[1]);
            }
            return result;
        }



        // Sybase SQL Anywhere 12
        private OleDbConnectionStringBuilder SybaseConnectionString()
        {
            var builder = new OleDbConnectionStringBuilder();

            builder.Provider = "SAOLEDB.12";
            builder.DataSource = this.ServerName;
            builder.Add("Initial Catalog", this.DatabaseName);
            builder.Add("User ID", this.UID);
            if (!String.IsNullOrEmpty(Password))
            {
                builder.Add("Password", this.Password);
                builder.PersistSecurityInfo = true;
            }
            builder.DataSource = this.ServerName;

            if (!String.IsNullOrEmpty(this.Host))   // 32bit ODBC
                builder.Add("Location", this.Host);
            else if (!String.IsNullOrEmpty(this.CommLinks))  // 64bit ODBC
            {
                // example: TCPIP{IP=PC-2015;DoBroad=No;ServerPort=8888}
                string s = CommLinks;

                if (s.StartsWith("TCPIP"))
                {
                    s = s.Replace("TCPIP", "");
                    if (s.Length > 3)
                    {
                        s = s.Substring(1, s.Length - 2); // remove surrounding { and }

                        var dict = String2Dict(s);

                        string location = dict["IP"];
                        if (dict.ContainsKey("ServerPort"))
                            location += ":" + dict["ServerPort"];
                        builder.Add("Location", location);
                    }
                    else
                    {
                        builder.Add("Location", "localhost");
                    }
                }
            }
            return builder;
        }

        /// <summary>
        /// return a ConnectionStringBuilder object or null if the driver name is not supported
        /// </summary>
        public DbConnectionStringBuilder ConnectionStringBuilder
        {
            get
            {
                switch (DSNDriverName)
                {
                    case "SQL Anywhere 12":
                        return SybaseConnectionString();
                    // TODO: other Database Types

                    default:
                        return null;
                }
            }
        }


        /// <summary>
        /// return a OLE-DB ConnectionString
        /// </summary>
        public string ConnectionString
        {
            get
            {
                var builder = ConnectionStringBuilder;
                if (builder != null)
                    return builder.ConnectionString;
                return null;
            }
        }





    }
}
