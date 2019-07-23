using System;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Connection
{

    /// <summary>
    /// Source types corresponding to ECMA, except 100
    /// </summary>
    public enum ConnectionSourceType
    {
        Unidentified = 0,
        ODBCBasedSource = 1,
        DAO = 2,
        ApplicationDefinedConnectionFile = 3,
        WebQuery = 4,
        OleDb = 5,
        TextBasedSource = 6,
        ADO = 7,
        DSP = 8,
        DataModelSource = 100
    }
    public class ExcelConnection : XmlHelper, IDisposable
    {
        #region XmlAttributes
        /// <summary>
        /// Attribute:          background
        /// ECMA Description:   Indicates whether the connection can be refreshed in the background (asynchronously).
        ///                     true if preferred usage of the connection is to refresh asynchronously in the background;
        ///                     false if preferred usage of the connection is to refresh synchronously in the foreground.
        ///                     This flag should be intentionally ignored in specific cases
        /// </summary>
        public bool BackgroundRefresh => TopNode.Attributes["background"]?.Value == "1";

        /// <summary>
        /// Attribute:          deleted
        /// ECMA Description:   Indicates whether the associated workbook connection has been deleted. true if the connection has been deleted; otherwise, false.
        ///                     Deleted connections contain only the attributes name and deleted=true, all other information is removed from the SpreadsheetML file.
        ///                     If a new connection is created with the same name as a deleted connection, then the deleted connection is overwritten by the new connection.
        /// </summary>
        public bool DeletedConnection => TopNode.Attributes["deleted"]?.Value == "1";
        /// <summary>
        /// Attribute:          deleted
        /// ECMA Description:   Specifies the user description for this connection.
        /// </summary>
        public string Description { get; }
        /// <summary>
        /// Attribute:          id
        /// ECMA Description:   Specifies The unique identifier of this connection.
        /// </summary>
        public int Id { get; }
        /// <summary>
        /// Attribute:          interval
        /// ECMA Description:   Specifies the number of minutes between automatic refreshes of the connection. When
        ///                     this attribute is not present, the connection is not automatically refreshed.
        /// </summary>
        public int AutomaticRefreshInterval => int.TryParse(TopNode.Attributes["interval"]?.Value, out int interval) ? interval : 0;
        /// <summary>
        /// Attribute:          keepAlive
        /// ECMA Description:   true when the spreadsheet application should make efforts to keep the connection
        ///                     open.When false, the application should close the connection after retrieving the
        ///                     information.This corresponds to the MaintainConnection property of a PivotCache object.
        /// </summary>
        public bool KeepConnectionOpen => TopNode.Attributes["keepAlive"]?.Value == "1";
        /// <summary>
        /// Attribute:          minRefreshableVersion
        /// ECMA Description:   For compatibility with legacy spreadsheet applications. This represents the minimum
        ///                     version # that is required to be able to correctly refresh the data connection. This
        ///                     attribute applies to connections that are used by a QueryTable.
        /// </summary>
        public string Name { get; }
        /// <summary>
        /// Attribute:          new
        /// ECMA Description:   true if the connection has not been refreshed for the first time; otherwise, false. This
        ///                     state can happen when the user saves the file before a query has finished returning.
        /// </summary>
        public bool IsNewConnection => TopNode.Attributes["new"]?.Value == "1";
        /// <summary>
        /// Attribute:          ocdFile
        /// ECMA Description:   Specifies the full path to external connection file from which this connection was created.If a connection fails during an attempt to refresh data, and
        ///                     reconnectionMethod=1, then the spreadsheet application will try again using information from the external connection file instead of the connection object
        ///                     embedded within the workbook.
        ///                     This is a benefit for data source and spreadsheetML document manageability. If the definition in the external connection file is changed(e.g., because of a database server
        ///                     name change), then the workbooks that made use of that connection will fail to connect with their internal connection information, and reload the new connection information
        ///                     from this file.
        ///                     This attribute is cleared by the spreadsheet application when the user manually edits the connection definition within the workbook.Can be expressed in URI or system-specific
        ///                     file path notation.
        /// </summary>
        public string ConnectionFile => TopNode.Attributes["ocdFile"]?.Value ?? "";
        /// <summary>
        /// Attribute:          resfreshOnLoad
        /// ECMA Description:   true if this connection should be refreshed when opening the file; otherwise, false.
        /// </summary>
        public bool RefreshOnOpen => TopNode.Attributes["refreshOnLoad"]?.Value == "1";

        /// <summary>
        /// Attribute:          excludeFromRefreshAll
        /// </summary>
        public bool RefreshOnRefreshAll => GetXmlNodeString("d:extLst/d:ext//@excludeFromRefreshAll") != "1";

        /// <summary>
        /// Attribute:          saveData
        /// ECMA Description:   true if the external data fetched over the connection to populate a table is to be saved with the workbook; otherwise, false.
        ///                     This exists for data security purposes - if no external data is saved in (or "cached") in the workbook, then current user credentials can be required every time to retrieve the
        ///                     relevant data, and people won't see the data the workbook author had last been using before saving the file.
        /// </summary>
        public bool SaveData => TopNode.Attributes["saveData"]?.Value == "1";
        /// Attribute:          savePassword
        /// ECMA Description:   true if the password is to be saved as part of the connection string; otherwise, False.
        /// </summary>
        public bool SavePassword => TopNode.Attributes["savePassword"]?.Value == "1";
        /// <summary>
        /// Attribute:          singleSignOnId
        /// ECMA Description:   Identifier for Single Sign On (SSO) used for authentication between an intermediate spreadsheetML server and the external data source.
        /// </summary>
        public string SingleSignOnId => TopNode.Attributes["singleSignOnId"]?.Value ?? "";

        /// <summary>
        /// Attribute:          sourceFile
        /// ECMA Description:   Used when the external data source is file-based. When a connection to such a data source fails, the spreadsheet application attempts to connect directly to this file. Can be
        ///                     expressed in URI or system-specific file path notation.
        /// </summary> 
        public string SourceDatabaseFile => TopNode.Attributes["sourceFile"]?.Value ?? "";
        /// <summary>
        /// Attribute:          type
        /// ECMA Description:   Specifies the data source type.
        /// </summary>
        public ConnectionSourceType ConnectionSourceType { get; }
        /// <summary>
        /// Database properties of the connection if available
        /// </summary>
        public ExcelDatabaseProperties DatabaseProperties { get; private set; }
        #endregion

        public ExcelConnection(XmlNamespaceManager nameSpaceManager, XmlElement connectionElement) : base(nameSpaceManager, connectionElement)
        {
            int.TryParse(connectionElement.GetAttribute("id"), out int connectionId);
            Id = connectionId;
            Name = connectionElement.GetAttribute("name");
            ConnectionSourceType = (ConnectionSourceType)Enum.Parse(typeof(ConnectionSourceType), connectionElement.GetAttribute("type"));
            Description = connectionElement.GetAttribute("description");

            if (connectionElement.SelectSingleNode("d:dbPr", nameSpaceManager) is XmlElement dbPrElement)
            {
                DatabaseProperties = new ExcelDatabaseProperties(dbPrElement);
            }
        }

        public void Dispose()
        {
            DatabaseProperties = null;
        }
    }
}