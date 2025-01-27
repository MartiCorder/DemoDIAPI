using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;

namespace DemoDIAPI.Classes
{
    public static class ClientController
    {
        public static void AddClient(SAPbobsCOM.Company company, string cardCode, string cardName)
        {
            if (RecordExists.IsInDatabase(company, "OCRD", "CardCode", cardCode))
            {
                Console.WriteLine($"El client amb CardCode '{cardCode}' ja existeix.");
            }
            else
            {
                var client = (BusinessPartners)company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                client.CardCode = cardCode;
                client.CardName = cardName;

                if (client.Add() == 0)
                {
                    Console.WriteLine($"Client {cardCode} creat correctament.");
                }
                else
                {
                    throw new Exception($"Error creant el client: {company.GetLastErrorDescription()}");
                }
                Utilities.Release(client);
            }
        }
    }
}
