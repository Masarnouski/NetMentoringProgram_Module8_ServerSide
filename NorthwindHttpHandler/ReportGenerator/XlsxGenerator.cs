using System.Linq;
using ClosedXML.Excel;
using NorthwindHttpHandler.DataAccess.GeneratedEntities;

namespace NorthwindHttpHandler.ReportGenerator
{
	public class XlsxGenerator
	{
		public XLWorkbook CreateWorkbook(IQueryable<Order> orders)
		{
			XLWorkbook workbook = new XLWorkbook();

			IXLWorksheet worksheet = workbook.Worksheets.Add("Northwind Worksheet");

			this.CreateHeader(worksheet);
			this.GenerateOrders(worksheet, orders);

			return workbook;
		}

		private void CreateHeader(IXLWorksheet worksheet)
		{
			worksheet.Cell("A1").Value = nameof(Order.OrderID);
			worksheet.Cell("B1").Value = nameof(Order.CustomerID);
			worksheet.Cell("C1").Value = nameof(Order.OrderDate);
			worksheet.Cell("D1").Value = nameof(Order.ShipCountry);
		}

		private void GenerateOrders(IXLWorksheet worksheet, IQueryable<Order> orders)
		{
			int rowIndex = 2;
			foreach(Order order in orders)
			{
				worksheet.Cell("A" + rowIndex.ToString()).Value = order.OrderID;
				worksheet.Cell("B" + rowIndex.ToString()).Value = order.CustomerID;
				worksheet.Cell("C" + rowIndex.ToString()).Value = order.OrderDate;
				worksheet.Cell("D" + rowIndex.ToString()).Value = order.ShipCountry;

				rowIndex++;
			}
		}
	}
}