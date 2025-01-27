using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;

namespace DemoDIAPI.Classes
{
    public static class ItemController
    {
        public static void AddItem(SAPbobsCOM.Company company, string itemCode, string itemName)
        {
            if (RecordExists.IsInDatabase(company, "OITM", "ItemCode", itemCode))
            {
                Console.WriteLine($"L'article amb ItemCode '{itemCode}' ja existeix.");
            }
            else
            {
                var item = (Items)company.GetBusinessObject(BoObjectTypes.oItems);
                item.ItemCode = itemCode;
                item.ItemName = itemName;
                item.UoMGroupEntry = -1;
                item.DefaultWarehouse = "01";

                if (item.Add() == 0)
                {
                    Console.WriteLine($"Article {itemCode} creat correctament.");
                }
                else
                {
                    throw new Exception($"Error creant l'article: {company.GetLastErrorDescription()}");
                }

                Utilities.Release(item);
            }
        }
    }
}
