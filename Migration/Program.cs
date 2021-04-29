using System;

namespace Migration
{
    class Program
    {
        public readonly static string connectionString = "";
        public readonly static string connectionStringForLive = "";

        public readonly static string orderPath = @"C:\Users\Muhsin\Desktop\TODO\AyakkabiDunyasi\Migration\gece3eadar\OrderItemMart 2021-Nisan2021.xlsx;";
        public readonly static string orderItemPath = @"C:\Users\Muhsin\Desktop\TODO\AyakkabiDunyasi\Migration\gece3eadar\OrderMart2021-Nisan2021.xlsx;";
        public readonly static string customerPath = @"C:\Users\Muhsin\Desktop\TODO\AyakkabiDunyasi\Migration\gece3eadar\second\Mart2021 -Nisan2021.xlsx;";
        public readonly static string customerRolePath = @"C:\Users\Muhsin\Desktop\TODO\AyakkabiDunyasi\Migration\gece3eadar\Ids.xlsx;";
        public readonly static string addressPath = @"C:\Users\Muhsin\Desktop\TODO\AyakkabiDunyasi\Migration\gece3eadar\adres\Mart 2021 - Nisan 2021.xlsx;";

        static void Main(string[] args)
        {
            //Order.ImportOrder(connectionStringForLive, orderPath, "Sayfa1");
            //Order.ImportOrderItem(connectionStringForLive, orderItemPath, "Sayfa1");
            //Customer.ImportCustomer(connectionStringForLive, customerPath, "Sayfa1");
            //Customer.ImportCustomerRoles(connectionStringForLive, customerRolePath, "Sayfa1");
            Address.ImportAddress(connectionString, addressPath, "Sayfa1");
            //Address.ImportAddress(connectionString, addressPath, "Temmuz 2017 - Aralık 2017");

            Console.WriteLine("Finish!!");
        }
    }
}
