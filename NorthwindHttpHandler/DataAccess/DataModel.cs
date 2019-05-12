using System.Data.Entity;
using NorthwindHttpHandler.DataAccess.GeneratedEntities;

namespace NorthwindHttpHandler.DataAccess
{
	public partial class DataModel : DbContext
	{
		public DataModel()
			: base("name=DataModel")
		{
		}

		public virtual DbSet<Order> Orders { get; set; }

		protected override void OnModelCreating(DbModelBuilder modelBuilder)
		{
			modelBuilder.Entity<Order>()
				.Property(e => e.CustomerID)
				.IsFixedLength();

			modelBuilder.Entity<Order>()
				.Property(e => e.Freight)
				.HasPrecision(6, 2);
		}
	}
}
