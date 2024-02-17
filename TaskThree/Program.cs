using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using TaskThree;
using TaskThree.files;

namespace Program
{
    class Program //entity framework
    {
        static void Main(string[] args)
        {
            start:
            Console.WriteLine("Добрый день! Добро пожаловать в приложение.");
            Console.WriteLine("Прошу вас не забывать, что пока что приложение поддерживает лишь одну таблицу!");
            Console.WriteLine("Однако вы можете поддержать проект на Bus*censure*.com и Patre*censure*.ru и т.д.");
            Console.WriteLine("Пожалуйста, введите путь до файла. Это должен быть полный путь до файла: ");
            var importer = new Importer(Console.ReadLine());
            if (importer.FilePath == null)
            {
                throw new Exception("Файл по заданному пути не найден");
            }

            var listProducts = importer.ImportExel<Product>((string)"Товары");
            var listCients = importer.ImportExel<Client>((string)"Клиенты");
            var listRequests = importer.ImportExel<Order>((string)"Заявки");
            if(listProducts == null || listCients == null || listRequests == null) 
            { 
                throw new Exception("Не найдена одна из таблиц. Проверьте, верный ли файл по заданному пути");
            }
            Controller controller = new Controller(listProducts, listCients, listRequests);
            Console.Clear();
            menu:
            Console.WriteLine("Отлично! Таблицы подключены. Выберите операцию: ");
            Console.WriteLine("1 - Найти клиентов, заказавших определенный товар. Нужно ввести название товара");
            Console.WriteLine("2 - Найти 'золотого клиента' за определенные год и месяц. Нужно ввести год и месяц");
            Console.WriteLine("3 - Переименоавть контакт организации. Нужно ввести название организации");
            Console.WriteLine("4 - Назад. Вернуться к подключаемому файлу");
            Console.WriteLine("5 - Отчистить консоль");
            Console.WriteLine("6 - Сохранить и выйти");
            int operation = int.Parse(Console.ReadLine());
            switch (operation)
            {
                case 1:
                    Console.WriteLine("Введите название товара: ");
                    string name = Console.ReadLine();
                    controller.ReadClientsFromProduct(name);
                    goto menu;
                case 2:
                    Console.WriteLine("Введите год в формате гггг: ");
                    int year = int.Parse(Console.ReadLine());
                    Console.WriteLine("Введите введите месяц 1-12: ");
                    int month = int.Parse(Console.ReadLine());
                    controller.ReadTheGoldClient(year, month);
                    goto menu;
                case 3:
                    Console.WriteLine("Введите организацию: ");
                    string organization = Console.ReadLine();
                    Console.WriteLine("Введите необходимое к замене ФИО: ");
                    string fullName = Console.ReadLine();
                    controller.RenameClient(organization, fullName);
                    goto menu;
                case 4:
                    Console.Clear();
                    goto start;
                case 5:
                    Console.Clear();
                    goto menu;
                case 6:
                    importer.ExportExel("Клиенты", controller.Clients);
                    break;
                default:
                    Console.WriteLine("Выбрана неверная операция. ");
                    goto menu;

            }
        }
    }
}