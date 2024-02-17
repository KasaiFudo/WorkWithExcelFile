using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using TaskThree.files;

namespace TaskThree
{
    class Controller
    {
        public List<Product> Products {get;set;}
        public List<Client> Clients { get;set;}
        public List<Order> Orders { get;set;}

        public Controller(List<Product> products, List<Client> clients, List<Order> requests)
        {
            this.Products = products;
            this.Clients = clients;
            Orders = requests;
        }

        public void ReadClientsFromProduct(string nameProduct)
        {
            var query = from order in Orders
                        join client in Clients on order.CustomerID equals client.CustomerID
                        join product in Products on order.ProductID equals product.ProductID
                        where product.Name == nameProduct
                        select new
                        {
                            ClientName = client.OrganizationName,
                            Amount = order.ProductAmount,
                            Price = product.Price,
                            OrderDate = order.Date
                        };
            foreach (var result in query)
            {
                Console.WriteLine($"Клиент: {result.ClientName}, Количество: {result.Amount}, Цена: {result.Price}, Дата заказа: {result.OrderDate}");
            }
        }
        public void ReadTheGoldClient(int year, int month)
        {
            var query = from order in Orders
                        where order.Date.Year == year && order.Date.Month == month
                        group order by order.CustomerID into grouped
                        orderby grouped.Sum(o => o.ProductAmount) descending
                        select new
                        {
                            CustomerID = grouped.Key,
                            TotalOrders = grouped.Sum(o => o.ProductAmount)
                        };

            // Получение информации о золотом клиенте
            var goldCustomer = query.FirstOrDefault();
            if (goldCustomer != null)
            {
                Console.WriteLine($"Золотой клиент: {Clients.FirstOrDefault(c => c.CustomerID == goldCustomer.CustomerID)?.OrganizationName}, Количество заказов: {goldCustomer.TotalOrders}");
            }
        }
        public void RenameClient(string organizationNameToUpdate, string newContact)
        {
            var clientToUpdate = Clients.FirstOrDefault(c => c.OrganizationName == organizationNameToUpdate);
            if (clientToUpdate != null)
            {
                clientToUpdate.Contact = newContact;
                List<Client> newList = new List<Client>();

                foreach (var client in Clients)
                {
                    if (client.OrganizationName == clientToUpdate.OrganizationName)
                    {
                        newList.Add(clientToUpdate);
                    }
                    else
                    {
                        newList.Add(client);
                    }
                }
                Clients = newList;
                Console.WriteLine("Данные переписаны");
            }
            else
            {
                Console.WriteLine($"Клиент с организацией {organizationNameToUpdate} не найден.");
            }
        }
    }
}
