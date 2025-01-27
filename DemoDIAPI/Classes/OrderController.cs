using System;
using System.Collections.Generic;
using System.Linq;
using SAPbobsCOM;
using DemoDIAPI.Entity;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;

namespace DemoDIAPI.Classes
{
    public static class OrderController
    {
        public static void ProcessOrders(SAPbobsCOM.Company company, List<Order> orders)
        {
            foreach (var order in orders)
            {
                try
                {
                    company.StartTransaction();

                    if (!RecordExists.IsInDatabase(company, "OCRD", "CardCode", order.CardCode))
                    {
                        ClientController.AddClient(company, order.CardCode, order.CardName);
                    }

                    foreach (var line in order.orderLines)
                    {
                        if (!RecordExists.IsInDatabase(company, "OITM", "ItemCode", line.ItemCode))
                        {
                            ItemController.AddItem(company, line.ItemCode, line.ItemName);
                        }
                    }
                    CreateOrder(company, order);

                    company.EndTransaction(BoWfTransOpt.wf_Commit);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    company.EndTransaction(BoWfTransOpt.wf_RollBack);
                }
            }
        }

        private static void CreateOrder(SAPbobsCOM.Company company, Order order)
        {
            var sapOrder = (Documents)company.GetBusinessObject(BoObjectTypes.oOrders);

            sapOrder.CardCode = order.CardCode;
            sapOrder.DocDate = order.DocDate;
            sapOrder.DocDueDate = order.DocDate.AddDays(7);
            sapOrder.Comments = order.Description;

            foreach (var line in order.orderLines)
            {
                sapOrder.Lines.ItemCode = line.ItemCode;
                sapOrder.Lines.Quantity = (double)line.Quantity;
                sapOrder.Lines.UnitPrice = (double)line.Price;
                sapOrder.Lines.DiscountPercent = (double)line.Discount;
                sapOrder.Lines.UoMEntry = 1;
                sapOrder.Lines.Add();
            }

            int result = sapOrder.Add();
            if (result == 0)
            {
                string docEntry = company.GetNewObjectKey();

                //TODO: Fer Queri que m'agafi el docEntry de la comanda i hem busqui el seu docNum


                Console.WriteLine($"Comanda Nº {docEntry} creada correctament.");
            }
            else
            {
                throw new Exception(company.GetLastErrorDescription());
            }

            Utilities.Release(sapOrder);
        }
    }
}