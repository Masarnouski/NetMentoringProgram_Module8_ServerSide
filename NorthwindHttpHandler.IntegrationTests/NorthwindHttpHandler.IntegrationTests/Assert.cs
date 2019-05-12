using System;

namespace NorthwindHttpHandler.IntegrationTests
{
	public static class Assert
	{
		public static void AreEqual(object actual, object expected)
		{
			if (!object.Equals(actual, expected))
			{
				throw new Exception($"actual: {actual} and expected: {expected} are not equal");
			}
		}
	}
}
