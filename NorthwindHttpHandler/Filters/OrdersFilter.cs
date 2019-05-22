using System;
using System.Collections.Specialized;
using System.Linq;
using NorthwindHttpHandler.DataAccess.GeneratedEntities;
using NorthwindHttpHandler.Exceptions;

namespace NorthwindHttpHandler
{
    public class OrdersFilter
    {
        private IQueryable<Order> _orders;
        private readonly NameValueCollection _queryString;
        public OrdersFilter(IQueryable<Order> orders, NameValueCollection queryString)
        {
            _queryString = queryString;
            _orders = orders;
        }
        public IQueryable<Order> GetFilteredOrdersQuery()
        {
            _orders = _orders.OrderBy(order => order.OrderID);

            this.FilterByCustomerId();
            this.FilterByDate();
            this.FilterByTake();
            this.FilterBySkip();

            return _orders;
        }

        private void FilterByCustomerId()
        {
            string customerId = _queryString["customer"];
            if (!string.IsNullOrEmpty(customerId))
            {
                _orders = _orders.Where(order => order.CustomerID.Equals(customerId, StringComparison.OrdinalIgnoreCase));
            }
        }

        private void FilterByDate()
        {
            string dateFrom = _queryString["dateFrom"];
            string dateTo = _queryString["dateTo"];
            if (!string.IsNullOrEmpty(dateFrom) && !string.IsNullOrEmpty(dateTo))
            {
                throw new InvalidRequestException("dateFrom and dateTo cannot be specified at the same time");
            }
            else if (!string.IsNullOrEmpty(dateFrom))
            {
                if (!DateTime.TryParse(dateFrom, out DateTime date))
                {
                    throw new InvalidRequestException("dateFrom is invalid");
                }

                _orders = _orders.Where(order => order.OrderDate.Value >= date);
            }
            else if (!string.IsNullOrEmpty(dateTo))
            {
                if (!DateTime.TryParse(dateTo, out DateTime date))
                {
                    throw new InvalidRequestException("dateTo is invalid");
                }

                _orders = _orders.Where(order => order.OrderDate.Value <= date);
            }
        }

        private void FilterByTake()
        {
            if (string.IsNullOrEmpty(_queryString["take"]))
            {
                return;
            }

            if (!int.TryParse(_queryString["take"], out int take)
                || take < 0)
            {
                throw new InvalidRequestException("take parameter must be a non-negative integer");
            }

            _orders = _orders.Take(take);
        }

        private void FilterBySkip()
        {
            if (string.IsNullOrEmpty(_queryString["skip"]))
            {
                return;
            }

            if (!int.TryParse(_queryString["skip"], out int skip)
                || skip < 0)
            {
                throw new InvalidRequestException("skip parameter must be a non-negative integer");
            }

            _orders = _orders.Skip(skip);
        }
    }
}