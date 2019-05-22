using System.Linq;
using System.Web;
using ClosedXML.Extensions;
using NorthwindHttpHandler.DataAccess.GeneratedEntities;

namespace NorthwindHttpHandler.ReportGenerator
{
    public class Generator
    {
        private readonly IQueryable<Order> _orders;
        private readonly XlsxGenerator _xlsxGenerator;
        private readonly XmlGenerator _xmlGenerator;

        public Generator(IQueryable<Order> orders)
        {
            _orders = orders;
            _xlsxGenerator = new XlsxGenerator();
            _xmlGenerator = new XmlGenerator();
        }

        public void CreateReport(HttpResponse httpResponse, ReportFormat format)
        {
            if (format == ReportFormat.Xlsx)
            {
                using (var wb = _xlsxGenerator.CreateWorkbook(_orders))
                {
                    wb.DeliverToHttpResponse(httpResponse, "NorthwindOrders.xlsx");
                }
            }
            else
            {
                httpResponse.ContentType = "application/xml";
                _xmlGenerator.WriteXmlToResponse(httpResponse, _orders);
            }
        }
    }
}
