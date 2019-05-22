using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Web;
using NorthwindHttpHandler.DataAccess;
using NorthwindHttpHandler.Exceptions;
using NorthwindHttpHandler.ReportGenerator;

namespace NorthwindHttpHandler
{
	public class NorthwindHandler : IHttpHandler
	{
		private readonly string[] _xmlAcceptTypes = { "text/xml", "application/xml" };

		public bool IsReusable => true;

		public void ProcessRequest(HttpContext context)
		{
            
			NameValueCollection queryString;
			if (context.Request.QueryString.HasKeys())
			{
				queryString = context.Request.QueryString;
			}
			else
			{
				queryString = this.ParseRequestBody(context.Request);
			}

			if (queryString == null)
			{
				queryString = new NameValueCollection();
			}

			ReportFormat format = this.ParseReportFormat(context.Request);
			using (DataModel model = new DataModel())
			{

                OrdersFilter filter = new OrdersFilter(model.Orders.AsQueryable(), queryString);

			    var filteredOrdersQuery = filter.GetFilteredOrdersQuery();

			    Generator generator = new Generator(filteredOrdersQuery);

                try
				{
					generator.CreateReport(context.Response, format);
				}
				catch (InvalidRequestException ex)
				{
					context.Response.StatusCode = 400;
					context.Response.Output.WriteLine(ex.Message);
				}
			}
		}

		private ReportFormat ParseReportFormat(HttpRequest request)
		{
			if (request.AcceptTypes != null
				&& _xmlAcceptTypes.Intersect(request.AcceptTypes).Any())
			{
				return ReportFormat.Xml;
			}

			return ReportFormat.Xlsx;
		}

		private NameValueCollection ParseRequestBody(HttpRequest request)
		{
			string input;

			using (StreamReader reader = new StreamReader(request.InputStream))
			{
				input = reader.ReadToEnd();
			}

			return HttpUtility.ParseQueryString(input);
		}
	}
}