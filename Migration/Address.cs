using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace Migration
{
    public class Address
    {
        public static void ImportAddress(string connectionString, string filePath, string sheetName)
        {
            #region Address Excel Import

            int i = 1;
            List<string> ids = new List<string>();
            string conn =
  $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath}" +
  @"Extended Properties='Excel 8.0;HDR=Yes;Connection Lifetime=3;Max Pool Size=2000;'";

            try
            {
                #region Datatable
                DataTable dt = new DataTable();
                dt.Columns.Add("Id", typeof(Int32));
                dt.Columns.Add("FirstName", typeof(string));
                dt.Columns.Add("LastName", typeof(string));
                dt.Columns.Add("Email", typeof(string));
                dt.Columns.Add("Company", typeof(string));
                dt.Columns.Add("CountryId", typeof(int));
                dt.Columns.Add("StateProvinceId", typeof(int));
                dt.Columns.Add("City", typeof(string));
                dt.Columns.Add("Address1", typeof(string));
                dt.Columns.Add("Address2", typeof(string));
                dt.Columns.Add("ZipPostalCode", typeof(string));
                dt.Columns.Add("PhoneNumber", typeof(string));
                dt.Columns.Add("FaxNumber", typeof(string));
                dt.Columns.Add("CustomAttributes", typeof(string));
                dt.Columns.Add("CreatedOnUtc", typeof(DateTime));
                dt.Columns.Add("IsBukoli", typeof(bool));
                dt.Columns.Add("BukoliPointCode", typeof(string));
                dt.Columns.Add("TaxOffice", typeof(string));
                dt.Columns.Add("TaxNumber", typeof(string));
                dt.Columns.Add("IdentificationNumber", typeof(string));
                dt.Columns.Add("MobilePhoneNumber", typeof(string));
                dt.Columns.Add("ErpCode", typeof(string));
                dt.Columns.Add("CityCountry", typeof(string));
                dt.Columns.Add("CountyId", typeof(int));
                dt.Columns.Add("IsCorporate", typeof(bool));
                dt.Columns.Add("PhysicalStoreId", typeof(int));
                dt.Columns.Add("Title", typeof(string));
                dt.Columns.Add(new DataColumn { ColumnName = "StreetId", DataType = typeof(int), AllowDBNull = true });
                dt.Columns.Add("ApartmentName", typeof(string));
                dt.Columns.Add("ApartmentNo", typeof(string));
                dt.Columns.Add("RoomNumber", typeof(string));
                dt.Columns.Add("FloorNumber", typeof(string));
                dt.Columns.Add("DistrictId", typeof(int));
                dt.Columns.Add("Code", typeof(string));
                dt.Columns.Add("CountyName", typeof(string));


                DataTable dtAddressMapping = new DataTable();
                dtAddressMapping.Columns.Add("Customer_Id", typeof(int));
                dtAddressMapping.Columns.Add("Address_Id", typeof(int));
                #endregion

                using (OleDbConnection connection = new OleDbConnection(conn))
                {
                    connection.Open();
                    OleDbCommand command = new OleDbCommand($"select * from [{sheetName}$]", connection);
                    using (OleDbDataReader dr = command.ExecuteReader())
                    {

                        while (dr.Read())
                        {
                            var addressId = dr[1];
                            var userId = dr[2];

                            var addresserpRefCode = dr[3];
                            addresserpRefCode = Helpers.SetMaxLength(addresserpRefCode, 99);

                            var addressName = dr[4];
                            addressName = Helpers.SetMaxLength(addressName, 49);

                            var firstname = dr[5];
                            var lastname = dr[6];


                            //var cityId = dr[7];
                            int? localCityId = 0;
                            //var districtCode = dr[7];
                            int? localDistrictId = 0;
                            string localEmail = "";

                            var town = dr[8];
                            var address = dr[10];


                            var postalCode = dr[11];
                            var bilePhone = dr[12];
                            var isBilling = dr[13];
                            //var isCompany = dr[13];
                            var companyName = dr[14];
                            var taxOffice = dr[15];
                            var taxNumber = dr[16];

                            taxOffice = Helpers.SetMaxLength(taxOffice, 499);
                            taxNumber = Helpers.SetMaxLength(taxNumber, 29);

                            i++;
                            if (!string.IsNullOrEmpty(addressId.ToString()))
                            {
                                try
                                {
                                    //StateId ve CountyID işlemi bu şekilde çok uzun sürdürüyor bunun için bir kolona bu değeri basıp
                                    //Sql tarafında update etmek daha hızlı bir çözüm oluyor.
                                    //Örnek sql kodunu method sonunda bulabilirsiniz.
                                    #region Comment
                                    //try
                                    //{
                                    //    using (SqlConnection con = new SqlConnection(connectionString))
                                    //    {
                                    //        con.Open();
                                    //        Console.WriteLine("Connection Established Successfully");

                                    //        LogMessage(Environment.NewLine + "Çalıştırılan Query'ler > ");
                                    //        using (var query = con.CreateCommand())
                                    //        {
                                    //            query.CommandText = "SELECT Id, StateId FROM County WHERE Name like '%" + town + "%'";
                                    //            SqlDataReader data = query.ExecuteReader();

                                    //            if (data.Read())
                                    //            {
                                    //                int? stateProvinceId = data["StateId"] as int? ?? default(int?);
                                    //                int? countyId = data["Id"] as int? ?? default(int?);
                                    //                if (stateProvinceId > 0)
                                    //                    localCityId = stateProvinceId;
                                    //                if (countyId > 0)
                                    //                    localDistrictId = countyId;
                                    //            }

                                    //        }
                                    //        if (!string.IsNullOrEmpty(userId.ToString()))
                                    //        {
                                    //            using (var query = con.CreateCommand())
                                    //            {
                                    //                query.CommandText = "SELECT Email FROM Customer WHERE Id = " + userId;
                                    //                SqlDataReader data = query.ExecuteReader();

                                    //                string email = "";
                                    //                if (data.Read())
                                    //                {
                                    //                    email = data["Email"] as string ?? default(string);
                                    //                }
                                    //                if (!string.IsNullOrEmpty(email))
                                    //                    localEmail = email;

                                    //            }
                                    //        }
                                    //    }
                                    //}
                                    //catch (Exception ex)
                                    //{
                                    //    //continue;
                                    //}
                                    #endregion


                                    #region Addres DataRowu Setleme
                                    DataRow row = dt.NewRow();
                                    row["Id"] = addressId;
                                    row["FirstName"] = firstname;
                                    row["LastName"] = lastname;
                                    row["Email"] = localEmail;
                                    row["Company"] = companyName;
                                    row["CountryId"] = 77;
                                    row["StateProvinceId"] = localCityId;
                                    row["City"] = null;
                                    row["Address1"] = address;
                                    row["Address2"] = userId;
                                    row["ZipPostalCode"] = postalCode;
                                    row["PhoneNumber"] = bilePhone;
                                    row["FaxNumber"] = null;
                                    row["CustomAttributes"] = town;
                                    row["CreatedOnUtc"] = DateTime.Now;
                                    row["IsBukoli"] = false;
                                    row["BukoliPointCode"] = null;
                                    row["TaxOffice"] = taxOffice;
                                    row["TaxNumber"] = taxNumber;
                                    row["IdentificationNumber"] = null;
                                    row["MobilePhoneNumber"] = null;
                                    row["ErpCode"] = addresserpRefCode;
                                    row["CityCountry"] = null;
                                    row["CountyId"] = localDistrictId;
                                    row["IsCorporate"] = false;
                                    row["PhysicalStoreId"] = 0;
                                    row["Title"] = addressName;
                                    row["StreetId"] = (object)DBNull.Value;
                                    row["ApartmentName"] = null;
                                    row["ApartmentNo"] = null;
                                    row["RoomNumber"] = null;
                                    row["FloorNumber"] = null;
                                    row["DistrictId"] = (object)DBNull.Value;
                                    row["Code"] = null;
                                    row["CountyName"] = null;

                                    dt.Rows.Add(row);
                                    #endregion

                                    try
                                    {
                                        #region CustomerAddres DataRowu Setleme
                                        if (!string.IsNullOrEmpty(userId.ToString()))
                                        {
                                            DataRow addressMappingRow = dtAddressMapping.NewRow();
                                            addressMappingRow["Customer_Id"] = userId;
                                            addressMappingRow["Address_Id"] = addressId;

                                            dtAddressMapping.Rows.Add(addressMappingRow);
                                        }
                                        #endregion
                                    }
                                    catch (Exception ex)
                                    {
                                    }

                                }
                                catch (Exception ex)
                                {
                                }
                            }
                            else
                            {

                            }
                        }
                    }
                }

                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connectionString, SqlBulkCopyOptions.KeepIdentity))
                {
                    //bulkCopy.BatchSize = 50000;
                    bulkCopy.BulkCopyTimeout = 0;
                    bulkCopy.DestinationTableName = "[Address]";
                    bulkCopy.WriteToServer(dt);
                }

                //using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connectionString, SqlBulkCopyOptions.KeepIdentity))
                //{
                //    //bulkCopy.BatchSize = 50000;
                //    bulkCopy.BulkCopyTimeout = 0;
                //    bulkCopy.DestinationTableName = "[CustomerAddresses]";
                //    bulkCopy.WriteToServer(dtAddressMapping);
                //}

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            #endregion

            //Address2 kısmı için
            //Bazı alanlarda telefon numarası basıldığı için uzunluğu 10 olanlara null atadık
            //  update "Address" set Address2 = null
            //  where Len(Address2) = 10

            //  UPDATE a
            //  SET a.Email = c.Email
            //  FROM "Address" AS a
            //  INNER JOIN Customer AS c
            //  ON a.Address2 = c.Id

            ////--CustomAttributes alanı için
            //--StateProvinceId güncelleme
            //      UPDATE a
            //      SET a.StateProvinceId = c.StateId
            //      FROM "Address" AS a
            //      INNER JOIN County AS c
            //      ON a.CustomAttributes = c."Name"


            //--CountyId güncelleme
            //    UPDATE a
            //    SET a.CountyId = c.Id
            //    FROM "Address" AS a
            //    INNER JOIN County AS c
            //    ON a.CustomAttributes = c."Name"
        }
    }
}
