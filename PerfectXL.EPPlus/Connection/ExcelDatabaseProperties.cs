using System;
using System.Xml;

namespace OfficeOpenXml.Connection
{
    /// <summary>
    /// Command type for the database connection
    /// </summary>
    public enum CommandType
    {
        Unidentified = 0,
        Cube = 1,
        SqlStatement = 2,
        Table = 3,
        DefaultInformation = 4,
        WebBased = 5
    }

    /// <summary>
    /// Database properties for a connection
    /// </summary>
    public class ExcelDatabaseProperties
    {
        /// <summary>
        /// Attribute: connection
        /// ECMA Description: The string used to initiate a session with a data source.
        /// </summary>
        public string ConnectionString { get; }
        /// <summary>
        /// Attribute: commandType
        /// ECMA Description: Specifies the custom data source command type. Values are passed to the custom data source provider.
        /// </summary>
        public CommandType CommandType { get; }
        /// <summary>
        /// Attribute: command
        /// ECMA Description: The string containing the database command to pass to the data provider that will interact with the external source in order to retrieve data.
        /// </summary>
        public string CommandString { get; }
        /// <summary>
        /// Attribute: serverCommand
        /// ECMA Description: Specifies a second command text string that is persisted when PivotTable server-based page fields are in use.
        ///                   For ODBC connections, serverCommand is usually a broader query than command(no WHERE clause is present in the former). Based on these 2 commands, parameter UI can
        ///                   be populated and parameterized queries can be constructed.
        /// </summary>
        public string ServerCommand { get; }

        /// <summary>
        /// Database properties for a connection
        /// </summary>
        /// <param name="dbPrElement">this as XmlElement</param>
        public ExcelDatabaseProperties(XmlElement dbPrElement)
        {
            CommandString = dbPrElement.GetAttribute("command");
            string commandTypeAttribute = dbPrElement.GetAttribute("commandType");
            CommandType = commandTypeAttribute == "" ? CommandType.Unidentified : (CommandType)Enum.Parse(typeof(CommandType), commandTypeAttribute, false);
            ConnectionString = dbPrElement.GetAttribute("connection");
            ServerCommand = dbPrElement.GetAttribute("serverCommand");
        }
    }
}