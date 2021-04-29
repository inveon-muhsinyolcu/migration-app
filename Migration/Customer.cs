using System;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace Migration
{
    public class Customer
    {
        public static void ImportCustomer(string connectionString, string filePath, string sheetName)
        {
            DataTable dt = new DataTable();

            //bir tane customer schema '.'n select atiyoruz/
            using (var adapter = new SqlDataAdapter($"SELECT TOP 0 * FROM [Customer]", new SqlConnection(connectionString)))
            {
                adapter.Fill(dt);
            };

            int i = 1;
            string conn =
  $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath}" +
  @"Extended Properties='Excel 8.0;HDR=Yes;'";
            using (OleDbConnection connection = new OleDbConnection(conn))
            {
                connection.Open();
                //OleDbCommand command = new OleDbCommand("select * from [Mart 2017 - Mart 2018$]", connection); //part1
                OleDbCommand command = new OleDbCommand($"select * from [{sheetName}$]", connection); //part2
                using (OleDbDataReader dr = command.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        i++;

                        var excelId = dr[1];

                        if (string.IsNullOrEmpty(excelId.ToString()) || excelId.ToString() == "0")
                            continue;

                        DataRow row = dt.NewRow();

                        var CreatedOnUtc = dr[4];
                        CreatedOnUtc = !string.IsNullOrEmpty(CreatedOnUtc.ToString()) ? CreatedOnUtc : DateTime.Now;

                        var LastActivityDateUtc = dr[5];
                        LastActivityDateUtc = !string.IsNullOrEmpty(LastActivityDateUtc.ToString()) ? LastActivityDateUtc : DateTime.Now;

                        var allowSms = dr[13];
                        var allowEmail = dr[11];
                        allowSms = !string.IsNullOrEmpty(allowSms.ToString()) ? allowSms : false;
                        allowEmail = !string.IsNullOrEmpty(allowEmail.ToString()) ? allowEmail : false;

                        var active = dr[6];
                        active = !string.IsNullOrEmpty(active.ToString()) ? active : false;

                        var firstName = dr[2];
                        firstName = Helpers.SetMaxLength(firstName, 49);

                        var lastName = dr[3];
                        lastName = Helpers.SetMaxLength(lastName, 49);

                        var mobilePhone = dr[10];
                        mobilePhone = Helpers.SetMaxLength(mobilePhone, 30);

                        row["Id"] = dr[1];
                        row["Username"] = dr[8]; //email
                        row["CustomerGuid"] = Guid.NewGuid();
                        row["Email"] = dr[8];
                        row["FirstName"] = firstName;
                        row["LastName"] = lastName;
                        row["MobilePhone"] = mobilePhone;
                        row["ErpCode"] = dr[7];
                        row["CreatedOnUtc"] = CreatedOnUtc ?? DateTime.Now;
                        row["LastActivityDateUtc"] = LastActivityDateUtc ?? DateTime.Now;
                        row["Password"] = dr[9];
                        row["PasswordFormatId"] = 0;
                        row["IsTaxExempt"] = false;
                        row["AffiliateId"] = 0;
                        row["LastIpAddress"] = dr[15];
                        row["VendorId"] = 0;
                        row["HasShoppingCartItems"] = false;
                        row["Active"] = active;
                        row["Deleted"] = false;
                        row["IsSystemAccount"] = false;
                        row["IsEmailValidated"] = false;
                        row["AllowEmailCommunication"] = allowEmail;
                        row["AllowSmsCommunication"] = allowSms;
                        row["StateProvinceId"] = 0;
                        row["CountyId"] = 0;
                        row["OrderCount"] = 0;
                        row["CountryId"] = 0;
                        row["LanguageId"] = 2;
                        row["SelectedPickUpInStore"] = false;
                        row["UseRewardPointsDuringCheckout"] = false;
                        row["LanguageAutomaticallyDetected"] = false;
                        row["CurrencyId"] = 0;
                        row["IsPhoneValidated"] = false;

                        dt.Rows.Add(row);
                    }
                }
            }

            var test = i;
            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connectionString, SqlBulkCopyOptions.KeepIdentity))
            {
                try
                {
                    //bulkCopy.BatchSize = 50000;
                    bulkCopy.BulkCopyTimeout = 0;
                    bulkCopy.DestinationTableName = "[Customer]";
                    bulkCopy.WriteToServer(dt);
                }
                catch (Exception ex) 
                {

                    throw;
                }
            }
        }
        public static void ImportCustomerRoles(string connectionString, string filePath, string sheetName)
        {
            DataTable dt = new DataTable();

            //bir tane customer schema '.'n select atiyoruz/
            using (var adapter = new SqlDataAdapter($"SELECT TOP 0 * FROM [Customer_CustomerRole_Mapping]", new SqlConnection(connectionString)))
            {
                adapter.Fill(dt);
            };

            string conn =
  $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath}" +
  @"Extended Properties='Excel 8.0;HDR=Yes;'";
            using (OleDbConnection connection = new OleDbConnection(conn))
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand($"select * from [{sheetName}$]", connection); 
                using (OleDbDataReader dr = command.ExecuteReader())
                {
                    int i = 1;
                    while (dr.Read())
                    {
                        i++;

                        var Id = dr[0];

                        DataRow row = dt.NewRow();

                        row["Customer_Id"] = Id;
                        row["CustomerRole_Id"] = 3; 

                        dt.Rows.Add(row);
                    }
                }
            }


            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connectionString, SqlBulkCopyOptions.KeepIdentity))
            {
                //bulkCopy.BatchSize = 50000;
                bulkCopy.BulkCopyTimeout = 0;
                bulkCopy.DestinationTableName = "[Customer_CustomerRole_Mapping]";
                bulkCopy.WriteToServer(dt);
            }
        }
    }
}
