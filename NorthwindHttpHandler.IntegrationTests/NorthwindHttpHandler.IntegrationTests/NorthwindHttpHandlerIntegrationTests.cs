using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;

namespace NorthwindHttpHandler.IntegrationTests
{
	public class NorthwindHttpHandlerIntegrationTests
	{
		private const string Host = "http://localhost/Report/";

		public void RunTests()
		{
			this.ProcessRequest_WhenDateFromAndDateToSpecified_BadRequestReturned();
			this.ProcessRequest_WhenDateToIsInvalid_BadRequestReturned();
			this.ProcessRequest_WhenDateFromIsInvalid_BadRequestReturned();
			this.ProcessRequest_WhenTakeIsInvalid_BadRequestReturned();
			this.ProcessRequest_WhenSkipIsInvalid_BadRequestReturned();

			this.ProcessRequest_WhenXmlIsRequested_XmlIsReturned("application/xml");
			this.ProcessRequest_WhenXmlIsRequested_XmlIsReturned("text/xml");
			this.ProcessRequest_WhenXlsxIsRequested_XlsxIsReturned();
			this.ProcessRequest_WhenAcceptHeaderIsEmpty_XlsxIsReturned();
			this.ProcessRequest_WhenQueryStringIsValid_ResponseIsOK();
			this.ProcessRequest_WhenRequestBodyIsValid_ResponseIsOK();
		}

		#region Negative tests
		private void ProcessRequest_WhenDateFromAndDateToSpecified_BadRequestReturned()
		{
			HttpResponseMessage actual;
			using (HttpClient httpClient = new HttpClient())
			{
				Uri url = new Uri(Host + "?dateTo=12.05.2019&dateFrom=12.05.2018");
				actual = httpClient.GetAsync(url).Result;
			}

			Assert.AreEqual(actual.StatusCode, HttpStatusCode.BadRequest);
		}

		private void ProcessRequest_WhenDateToIsInvalid_BadRequestReturned()
		{
			HttpResponseMessage actual;
			using (HttpClient httpClient = new HttpClient())
			{
				Uri url = new Uri(Host + "?dateTo=fakeErrorDate");
				actual = httpClient.GetAsync(url).Result;
			}

			Assert.AreEqual(actual.StatusCode, HttpStatusCode.BadRequest);
		}

		private void ProcessRequest_WhenDateFromIsInvalid_BadRequestReturned()
		{
			HttpResponseMessage actual;
			using (HttpClient httpClient = new HttpClient())
			{
				Uri url = new Uri(Host + "?dateFrom=fakeErrorDate");
				actual = httpClient.GetAsync(url).Result;
			}

			Assert.AreEqual(actual.StatusCode, HttpStatusCode.BadRequest);
		}

		private void ProcessRequest_WhenTakeIsInvalid_BadRequestReturned()
		{
			HttpResponseMessage actual;
			using (HttpClient httpClient = new HttpClient())
			{
				Uri url = new Uri(Host + "?take=-1");
				actual = httpClient.GetAsync(url).Result;
			}

			Assert.AreEqual(actual.StatusCode, HttpStatusCode.BadRequest);
		}

		private void ProcessRequest_WhenSkipIsInvalid_BadRequestReturned()
		{
			HttpResponseMessage actual;
			using (HttpClient httpClient = new HttpClient())
			{
				Uri url = new Uri(Host + "?skip=-1");
				actual = httpClient.GetAsync(url).Result;
			}

			Assert.AreEqual(actual.StatusCode, HttpStatusCode.BadRequest);
		}
		#endregion

		#region Positive tests
		private void ProcessRequest_WhenXmlIsRequested_XmlIsReturned(string accept)
		{
			HttpResponseMessage actual;
			using (HttpClient httpClient = new HttpClient())
			{
				httpClient.DefaultRequestHeaders.Add("Accept", accept);
				actual = httpClient.GetAsync(new Uri(Host)).Result;
			}

			Assert.AreEqual(actual.StatusCode, HttpStatusCode.OK);
			Assert.AreEqual(actual.Content.Headers.ContentType.ToString(), "application/xml");
		}

		private void ProcessRequest_WhenXlsxIsRequested_XlsxIsReturned()
		{
			HttpResponseMessage actual;
			using (HttpClient httpClient = new HttpClient())
			{
				httpClient.DefaultRequestHeaders.Add("Accept", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
				actual = httpClient.GetAsync(new Uri(Host)).Result;
			}

			Assert.AreEqual(actual.StatusCode, HttpStatusCode.OK);
			Assert.AreEqual(actual.Content.Headers.ContentType.ToString(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
		}

		private void ProcessRequest_WhenAcceptHeaderIsEmpty_XlsxIsReturned()
		{
			HttpResponseMessage actual;
			using (HttpClient httpClient = new HttpClient())
			{
				actual = httpClient.GetAsync(new Uri(Host)).Result;
			}

			Assert.AreEqual(actual.StatusCode, HttpStatusCode.OK);
			Assert.AreEqual(actual.Content.Headers.ContentType.ToString(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
		}

		private void ProcessRequest_WhenQueryStringIsValid_ResponseIsOK()
		{
			HttpResponseMessage actual;
			using (HttpClient httpClient = new HttpClient())
			{
				Uri url = new Uri(Host + "?customer=TOMSP&dateTo=12.05.2000&take=5&skip=5");
				actual = httpClient.GetAsync(url).Result;
			}

			Assert.AreEqual(actual.StatusCode, HttpStatusCode.OK);
		}

		private void ProcessRequest_WhenRequestBodyIsValid_ResponseIsOK()
		{
			HttpResponseMessage actual;
			using (HttpClient httpClient = new HttpClient())
			{
				FormUrlEncodedContent content = new FormUrlEncodedContent(new[]
				{
					new KeyValuePair<string, string>("customer", "TOMSP"),
					new KeyValuePair<string, string>("dateTo", "12.05.2000"),
					new KeyValuePair<string, string>("take", "5"),
					new KeyValuePair<string, string>("skip", "5")
				});

				actual = httpClient.PostAsync(new Uri(Host), content).Result;
			}

			Assert.AreEqual(actual.StatusCode, HttpStatusCode.OK);
		}
		#endregion
	}
}
