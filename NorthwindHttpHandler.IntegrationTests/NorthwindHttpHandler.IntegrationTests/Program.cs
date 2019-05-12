using System;

namespace NorthwindHttpHandler.IntegrationTests
{
	public class Program
	{
		public static void Main(string[] args)
		{
			var tests = new NorthwindHttpHandlerIntegrationTests();

			tests.RunTests();

			Console.WriteLine("All tests have passed!");
			Console.ReadLine();
		}
	}
}
