using System.Linq;
using System.Web;
using System.Xml;
using NorthwindHttpHandler.DataAccess.GeneratedEntities;

namespace NorthwindHttpHandler.ReportGenerator
{
	public class XmlGenerator
	{
		public void WriteXmlToResponse(HttpResponse httpResponse, IQueryable<Order> orders)
		{
			using (XmlWriter writer = XmlWriter.Create(httpResponse.OutputStream))
			{
				writer.WriteStartDocument();
				writer.WriteStartElement("Orders");

				foreach (Order order in orders)
				{
					writer.WriteStartElement("Order");

					writer.WriteElementString(nameof(order.OrderID), order.OrderID.ToString());
					writer.WriteElementString(nameof(order.CustomerID), order.CustomerID.ToString());
					writer.WriteElementString(nameof(order.OrderDate), order.OrderDate.ToString());
					writer.WriteElementString(nameof(order.ShipCountry), order.ShipCountry);

					writer.WriteEndElement();
				}

				writer.WriteEndElement();
				writer.WriteEndDocument();
			}
		}
	}
}