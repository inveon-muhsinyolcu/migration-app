using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace Migration
{
    public class Order
    {
        public static void ImportOrder(string connectionString, string filePath, string sheetName)
        {
            if (connectionString == null || filePath == null || sheetName == null)
                return;

            #region order Excel Import
            DataTable dt = new DataTable();

            int i = 1;
            //Guid kolonu exception fıratıyordu. Bu sebeple aşağıda tablo klonunu alıp, onun üzerinde işlem yaparak tek tek kolon eklemekten ve hatadan kurtulmuş olduk.
            using (var adapter = new SqlDataAdapter($"SELECT TOP 0 * FROM [Order]", new SqlConnection(connectionString)))
            {
                adapter.Fill(dt);
            };

            string conn = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath}" +
                        @"Extended Properties='Excel 8.0;HDR=Yes;'";
            using (OleDbConnection connection = new OleDbConnection(conn))
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand($"select * from [{sheetName}$]", connection);
                using (OleDbDataReader dr = command.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        i++;
                        var orderId = dr[0];
                        var userId = dr[3];

                        if (orderId == null || userId == null)
                            continue;

                        var orderDate = dr[1];

                        if (string.IsNullOrEmpty(orderDate.ToString()))
                            orderDate = DateTime.Now;

                        var orderNumber = dr[2];
                        //var name = dr[4];
                        //var lastName = dr[5];
                        //var email = dr[6];
                        //var shipMobilepPhone = dr[7];
                        var orderStatus = dr[4];
                        if (string.IsNullOrEmpty(orderStatus.ToString()))
                            orderStatus = "onaylandı";

                        var shipAdressId = dr[5];
                        var bilAddressId = dr[6];
                        if (string.IsNullOrEmpty(bilAddressId.ToString()))
                            bilAddressId = 0;

                        var shipperId = dr[7]; //todo burada id paylaşılmamış.
                        //var paymentmethodSystemName = dr[11]; //todo
                        //var cartBankName =  dr[12];
                        var cardNumber = dr[10];
                        var cardOwner = dr[11];
                        var shippingCost = dr[12];
                        var interestCost = dr[13];
                        //var installMeenCount = dr[18];
                        var totalAmount = dr[14];
                        if (string.IsNullOrEmpty(totalAmount.ToString()))
                            totalAmount = 0;

                        var paymentCurrency = dr[16];
                        //var userRole = dr[20];
                        var ipAddress = dr[18];

                        var orderDiscount = dr[15];
                        if (string.IsNullOrEmpty(orderDiscount.ToString()))
                            orderDiscount = 0;

                        var refundedAmount = dr[19];
                        if (string.IsNullOrEmpty(refundedAmount.ToString()))
                            refundedAmount = 0;

                        var giftNote = dr[20];
                        if (!string.IsNullOrEmpty(giftNote.ToString()))
                            giftNote = Helpers.SetMaxLength(giftNote, 499);

                        int orderStatusId = 0;
                        var excelStatus = orderStatus.ToString().ToLower();


                        switch (excelStatus)
                        {
                            case "iade edildi":
                            case "iptal edildi":
                            case "onaylandı":
                                orderStatusId = 30; //Delivered - DeliveredFromTheStore
                                break;
                            case "kargolandı":
                            default:
                                orderStatusId = 20; //DeliveredToCargo
                                break;
                        }

                        if (!string.IsNullOrEmpty(orderId.ToString()) /*&&/* userRole != null /*&& userRole.ToString() != "Guest"*/)
                        {
                            DataRow row = dt.NewRow();
                            var guidID = Guid.NewGuid();
                            row["OrderGuid"] = guidID;
                            ///
                            row["Id"] = orderId;
                            row["StoreId"] = 1;
                            row["CustomerId"] = userId;
                            row["BillingAddressId"] = bilAddressId;
                            row["ShippingAddressId"] = shipAdressId;
                            row["PickUpInStore"] = false;
                            row["OrderStatusId"] = orderStatusId;
                            row["ShippingStatusId"] = 35;
                            row["PaymentStatusId"] = 30;
                            //row["PaymentMethodSystemName"] = paymentmethodSystemName;
                            row["CustomerCurrencyCode"] = paymentCurrency;
                            row["CurrencyRate"] = 1.0;
                            row["CustomerTaxDisplayTypeId"] = 0;
                            row["VatNumber"] = null;
                            row["OrderSubtotalInclTax"] = 0.0;
                            row["OrderSubtotalExclTax"] = 0.0;
                            row["OrderSubTotalDiscountInclTax"] = 0.0;
                            row["OrderSubTotalDiscountExclTax"] = 0.0;
                            row["OrderShippingInclTax"] = shippingCost;
                            row["OrderShippingExclTax"] = shippingCost;
                            row["PaymentMethodAdditionalFeeInclTax"] = interestCost;
                            row["PaymentMethodAdditionalFeeExclTax"] = interestCost;
                            row["OrderTax"] = 0.0;
                            row["orderDiscount"] = orderDiscount;
                            row["OrderTotal"] = totalAmount;
                            row["RefundedAmount"] = refundedAmount;
                            row["RewardPointsWereAdded"] = 0.0;
                            row["CheckoutAttributeDescription"] = "";
                            row["CheckoutAttributesXml"] = null;
                            row["CustomerLanguageId"] = 2;
                            row["AffiliateId"] = 0;
                            row["CustomerIp"] = ipAddress;
                            row["AllowStoringCreditCardNumber"] = 0;
                            row["CardType"] = null;
                            //row["CardName"] = cartBankName;
                            //row["CardNumber"] = cardNumber;
                            row["MaskedCreditCardNumber"] = cardNumber;
                            row["CardCvv2"] = null;
                            row["CardExpirationMonth"] = null;
                            row["CardExpirationYear"] = null;
                            row["PaymentCardName"] = null;
                            //row["PaymentBankName"] = cartBankName;
                            row["PaymentGatewayName"] = "";
                            row["PaymentInstallmentCount"] = 0;
                            row["PaymentTypeErpCode"] = null;
                            row["AuthorizationTransactionId"] = null;
                            row["AuthorizationTransactionCode"] = null;
                            row["AuthorizationTransactionResult"] = null;
                            row["CaptureTransactionId"] = null;
                            row["CaptureTransactionResult"] = null;
                            row["SubscriptionTransactionId"] = null;
                            row["PaidDateUtc"] = DateTime.Now;
                            row["ShippingRateComputationMethodSystemName"] = "";
                            row["CustomValuesXml"] = null;
                            row["Deleted"] = false;
                            row["CreatedOnUtc"] = orderDate;
                            row["BukoliInvoiceNo"] = null;
                            row["FraudStatusId"] = 0;
                            row["OrderNote"] = null;
                            row["GiftNote"] = giftNote;
                            row["IsExported"] = true;
                            row["ExportedOnUtc"] = DateTime.Now;
                            row["ErpCode"] = null;
                            row["BillingNumber"] = null;
                            row["OrderTypeId"] = 0;
                            row["Is3DOrder"] = 0;
                            row["IsMobileOrder"] = 0;
                            row["RelatedReferral"] = null;
                            row["OrderNumber"] = orderNumber;
                            row["ReturnOrderTrackingNumber"] = null;
                            row["LoyaltyCardNumber"] = null;
                            row["SMSVerificationCode"] = null;
                            row["IsImpersonated"] = false;
                            row["ImpersonatedById"] = DBNull.Value;
                            row["ImpersonatedStoreId"] = DBNull.Value;
                            row["ImpersonatedStaffId"] = DBNull.Value;
                            row["PaymentMethodAdditionalFeePercent"] = 0.0;
                            row["ReturnedOnUtc"] = DBNull.Value;
                            row["PhysicalStorePosDeviceId"] = DBNull.Value;
                            row["ErpStatusId"] = DBNull.Value;
                            row["ErpTransactionCode"] = DBNull.Value;
                            row["ErpTransactionDate"] = DBNull.Value;
                            row["ErpInvoiceNumber"] = DBNull.Value;
                            row["ErpCode2"] = DBNull.Value;
                            row["OrderErpInformation"] = DBNull.Value;
                            row["OrderFreight"] = 0.0;
                            row["UsedLoyaltyPointTotal"] = 0.0;
                            row["PaymentCurrencyCode"] = paymentCurrency;
                            row["CargoLineId"] = DBNull.Value;
                            row["ReturnExportedOnUtc"] = DBNull.Value;
                            row["ReturnOrderErpInformation"] = DBNull.Value;
                            row["CheckOutSessionValue"] = DBNull.Value;
                            row["PaymentMethodUsedPoint"] = 0.0;
                            row["UsedLoyaltyPointMultiplir"] = 1.0;
                            row["InvoiceUrl"] = DBNull.Value;
                            row["DeliveryBarcode"] = DBNull.Value;
                            row["ShippingMethod"] = shipperId;
                            row["UsedLoyaltyPointMultiplier"] = 0.0;

                            dt.Rows.Add(row);
                        }
                        else
                        {

                        }
                    }
                }
            }
            var test = i;
            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connectionString, SqlBulkCopyOptions.KeepIdentity))
            {
                //bulkCopy.BatchSize = 50000;
                bulkCopy.BulkCopyTimeout = 0;
                bulkCopy.DestinationTableName = "[Order]";
                bulkCopy.WriteToServer(dt);
            }
            #endregion
        }

        public static void ImportOrderItem(string connectionString, string filePath, string sheetName)
        {
            if (connectionString == null || filePath == null || sheetName == null)
                return;

            List<string> notFoundProducts = new List<string>();

            DataTable dt = new DataTable();
            using (var adapter = new SqlDataAdapter($"SELECT TOP 0 * FROM [OrderItem]", new SqlConnection(connectionString)))
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
                OleDbCommand command = new OleDbCommand($"select * from [{sheetName}$]", connection);
                using (OleDbDataReader dr = command.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        i++;
                        var orderId = dr[0];
                        var orderItemId = dr[1];
                        var sku = dr[14];
                        sku = Helpers.SetMaxLength(sku, 49);
                        var productName = dr[3];
                        var quantity = dr[4];

                        var productPrice = dr[5];
                        if (string.IsNullOrEmpty(productPrice.ToString()))
                            productPrice = 0;

                        var salePrice = dr[6];
                        var taxRate = dr[7];
                        var discount = dr[8];
                        var lineTotal = dr[9];
                        var taxAmount = dr[10];
                        var interestCost = dr[11];
                        var orderTotal = dr[12];
                        var currency = dr[13];
                        var barcode = dr[2];
                        barcode = Helpers.SetMaxLength(barcode, 99);
                        var status = dr[15];

                        var trackingNumber = dr[16];
                        var invoiceNumber = dr[17];

                        int productId = 0;

                        //using (SqlConnection con = new SqlConnection(connectionString))
                        //{
                        //    con.Open();
                        //    Console.WriteLine("Connection Established Successfully > " + sku);

                        //    using (var query = con.CreateCommand())
                        //    {
                        //        query.CommandText = "SELECT Id FROM Product2 WHERE Sku = '" + sku + "'";
                        //        SqlDataReader data = query.ExecuteReader();

                        //        if (data.Read())
                        //        {
                        //            int id = data.GetInt32(0);
                        //            int? productIdValue = data["Id"] as int? ?? default(int?);
                        //            if (productIdValue > 0)
                        //                productId = id;
                        //            else
                        //                notFoundProducts.Add(sku.ToString());
                        //        }

                        //    }
                        //}


                        int statusId = 0;
                        var excelStatus = status.ToString().ToLower();
                        switch (excelStatus)
                        {
                            case "iade edildi":
                                statusId = 70;
                                break;
                            case "iptal edildi":
                                statusId = 80;
                                break;
                            case "onaylandı":
                                statusId = 90;
                                break;
                            case "kargolandı":
                            default:
                                statusId = 50;
                                break;
                        }

                        if (!string.IsNullOrEmpty(orderItemId.ToString()))
                        {
                            DataRow row = dt.NewRow();
                            row["Id"] = int.Parse(orderItemId.ToString());
                            row["OrderItemGuid"] = Guid.NewGuid();
                            row["OrderId"] = orderId;
                            row["ProductId"] = 0;
                            row["Sku"] = sku;
                            row["Quantity"] = quantity;
                            row["UnitPriceInclTax"] = productPrice;
                            row["UnitPriceExclTax"] = productPrice;
                            row["PriceInclTax"] = salePrice;
                            row["PriceExclTax"] = salePrice;
                            row["DiscountAmountInclTax"] = discount;
                            row["DiscountAmountExclTax"] = discount;
                            row["OriginalProductCost"] = 0.0;
                            row["AttributeDescription"] = DBNull.Value;
                            row["AttributesXml"] = DBNull.Value;
                            row["DownloadCount"] = 0;
                            row["IsDownloadActivated"] = false;
                            row["LicenseDownloadId"] = 0;
                            row["ItemWeight"] = DBNull.Value;
                            row["RentalStartDateUtc"] = DBNull.Value;
                            row["RentalEndDateUtc"] = DBNull.Value;
                            row["ShoppingCartDealId"] = DBNull.Value;
                            row["PromotionTypeId"] = 0;
                            row["ProductName"] = productName;
                            row["ProductCost"] = 0.0;
                            row["ProductOldPrice"] = 0.0;
                            row["IntegrationId"] = 0;
                            row["ReturnQuantity"] = 0;
                            row["IsChildItem"] = false;
                            row["ParentId"] = 0;
                            row["CanceledQuantity"] = 0;
                            row["ProductPrice"] = productPrice;
                            row["TierPrice"] = DBNull.Value;
                            row["InventoryStatusId"] = 0;
                            row["PackagedQuantity"] = 0;
                            row["TierPriceCustomerRoleName"] = DBNull.Value;
                            row["TierPriceCorporateCustomerName"] = DBNull.Value;
                            row["TierPriceCorporateCustomerDescription"] = DBNull.Value;
                            row["OrderItemStatusId"] = statusId;
                            row["ModifiedOnUtc"] = DateTime.Now;
                            row["Barcode"] = barcode;
                            row["PreventErpExport"] = false;
                            row["TrackingUrl"] = trackingNumber.ToString();
                            row["ErpLineId"] = DBNull.Value;
                            row["ProductNote"] = DBNull.Value;
                            row["IsGift"] = DBNull.Value;
                            row["GiftNote"] = DBNull.Value;
                            row["WareHouseAddress"] = DBNull.Value;
                            row["WareHouseCode"] = DBNull.Value;
                            row["ReturnExportOnUtc"] = DBNull.Value;
                            row["ReturnBillingNumber"] = DBNull.Value;
                            row["InvoiceNumber"] = invoiceNumber;
                            row["InvoiceDate"] = DBNull.Value;
                            row["InvoiceQty"] = DBNull.Value;
                            row["ReturnItemStoreCode"] = DBNull.Value;
                            row["InvoiceLineId"] = DBNull.Value;

                            dt.Rows.Add(row);
                        }

                    }

                    Helpers.LogMessage("NotFound Products>>" + Environment.NewLine);
                    Helpers.LogMessage(String.Join(',', notFoundProducts));

                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connectionString, SqlBulkCopyOptions.KeepIdentity))
                    {
                        //bulkCopy.BatchSize = 50000;
                        bulkCopy.BulkCopyTimeout = 0;
                        bulkCopy.DestinationTableName = "[OrderItem]";
                        bulkCopy.WriteToServer(dt);
                    }

                }
            }
        }
    }
}
