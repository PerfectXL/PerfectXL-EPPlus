using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Connection
{
    public class ExcelConnections : XmlHelper, IEnumerable<ExcelConnection>, IDisposable
    {
        private Dictionary<int, ExcelConnection> _connections;

        public IEnumerator<ExcelConnection> GetEnumerator()
        {
            return _connections.Values.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _connections.Values.GetEnumerator();
        }

        public ExcelConnections(ExcelPackage package, XmlNamespaceManager nameSpaceManager)
            : base(nameSpaceManager, null)
        {
            _connections = new Dictionary<int, ExcelConnection>();

            Uri connectionsUri = package.Workbook.ConnectionsUri;
            
            if (!package.Package.PartExists(connectionsUri)) return;

            XmlDocument connectionsXml = new XmlDocument();

            LoadXmlSafe(connectionsXml, package.Package.GetPart(connectionsUri).GetStream());
            TopNode = connectionsXml.DocumentElement;
            foreach (XmlElement connectionElement in TopNode.SelectNodes("//d:connection", NameSpaceManager))
            {
                var connection = new ExcelConnection(NameSpaceManager, connectionElement);
                _connections.Add(connection.Id, connection);
            }
        }

        public ExcelConnection GetConnectionById(int id)
        {
            return _connections.TryGetValue(id, out ExcelConnection connection) ? connection : null;
        }

        public void Dispose()
        {
            foreach (ExcelConnection connection in _connections.Values)
            {
                ((IDisposable)connection).Dispose();
            }

            _connections = null;
        }
    }
}