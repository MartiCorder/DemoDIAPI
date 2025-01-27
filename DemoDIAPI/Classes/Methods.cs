using System;
using System.Collections.Generic;
using DemoDIAPI.Classes;
using DemoDIAPI;
using SAPbobsCOM;
/*public static class Methods
{
    public static void CrearProducte(SAPbobsCOM.Company company, string itemCode, string itemName)
    {
        Console.WriteLine("***Crear producte***");

        var item = (SAPbobsCOM.Items)company.GetBusinessObject(BoObjectTypes.oItems);

        item.ItemCode = itemCode;
        item.ItemName = itemName;

        Console.WriteLine($"Provant de crear producte: {itemCode}");

        if (item.Add() != 0)
            Console.WriteLine(company.GetLastErrorDescription());
        else
            Console.WriteLine("Article creat correctament!");

        Utilities.Release(item);
    }

    public static void CrearProductes(SAPbobsCOM.Company company, string excelPath)
    {
        Console.WriteLine("***Crear productes des de Excel***");
        var reader = new ExcelReader();
        var excelDataList = reader.ReadExcelFile(excelPath);

        if (excelDataList == null || excelDataList.Count == 0)
        {
            Console.WriteLine("No s'han trobat dades a l'Excel.");
            return;
        }

        foreach (var data in excelDataList)
        {
            var item = (SAPbobsCOM.Items)company.GetBusinessObject(BoObjectTypes.oItems);
            try
            {
                item.ItemCode = data.ItemCode;
                item.ItemName = data.ItemName;

                int result = item.Add();
                if (result != 0)
                {
                    Console.WriteLine($"Error creant {data.ItemCode}: {company.GetLastErrorDescription()}");
                }
                else
                {
                    Console.WriteLine($"Article {data.ItemCode} creat correctament!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Excepció en crear producte: {ex.Message}");
            }
        }
    }

    public static void ConsultarArticle(SAPbobsCOM.Company company, string itemCode)
    {
        Console.WriteLine("***Consultar article***");

        var query = $"""
            SELECT
                T0.ItemCode as 'CodiArticle',
                T0.ItemName,
                T1.ItmsGrpNam
            FROM OITM T0
            INNER JOIN OITB T1 ON T0.[ItmsGrpCod] = T1.[ItmsGrpCod]
            WHERE
                T0.ItemCode = '{itemCode}'
            """;

        var recordset = (SAPbobsCOM.Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);

        try
        {
            recordset.DoQuery(query);

            while (!recordset.EoF)
            {
                string codiArticle = recordset.Fields.Item("CodiArticle").Value.ToString();
                string itemName = recordset.Fields.Item("ItemName").Value.ToString();
                string itemGroupName = recordset.Fields.Item("ItmsGrpNam").Value.ToString();

                Console.WriteLine($"Hem recuperat l'article: {codiArticle}, amb nom {itemName} i grup {itemGroupName}");

                recordset.MoveNext();
            }
        }
        finally
        {
            Utilities.Release(recordset);
        }
    }

    public static List<Item> ConsultarArticlesAgrupats(SAPbobsCOM.Company company)
    {
        Console.WriteLine("***Consultar articles agrupats***");

        var items = new List<Item>();
        var query = @"
            SELECT 
                T0.ItemCode, 
                T0.ItemName, 
                T1.ItmsGrpNam 
            FROM OITM T0
            INNER JOIN OITB T1 ON T0.ItmsGrpCod = T1.ItmsGrpCod";

        var recordset = (SAPbobsCOM.Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);

        try
        {
            recordset.DoQuery(query);

            while (!recordset.EoF)
            {
                var item = new Item
                {
                    ItemCode = recordset.Fields.Item("ItemCode").Value.ToString(),
                    ItemName = recordset.Fields.Item("ItemName").Value.ToString(),
                    ItemGroup = recordset.Fields.Item("ItmsGrpNam").Value.ToString()
                };

                items.Add(item);
                recordset.MoveNext();
            }
        }
        finally
        {
            Utilities.Release(recordset);
        }

        return items;
    }

    public static List<Client> ConsultarClientsAmbArticles(SAPbobsCOM.Company company)
    {
        Console.WriteLine("***Consultar clients amb articles***");

        var clients = new List<Client>();

        string query = @"
        SELECT 
            C.CardCode, 
            C.CardName, 
            C.LicTradNum AS FederalTaxId,
            I.ItemCode,
            I.ItemName,
            G.ItmsGrpNam AS ItemGroup
        FROM OCRD C
        INNER JOIN INV1 L ON C.CardCode = L.CardCode
        INNER JOIN OITM I ON L.ItemCode = I.ItemCode
        INNER JOIN OITB G ON I.ItmsGrpCod = G.ItmsGrpCod
        ORDER BY C.CardCode";

        var recordset = (SAPbobsCOM.Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);

        try
        {
            recordset.DoQuery(query);

            while (!recordset.EoF)
            {
                string cardCode = recordset.Fields.Item("CardCode").Value.ToString();
                var client = clients.FirstOrDefault(c => c.CardCode == cardCode);

                if (client == null)
                {
                    client = new Client
                    {
                        CardCode = cardCode,
                        CardName = recordset.Fields.Item("CardName").Value.ToString(),
                        FederalTaxId = recordset.Fields.Item("FederalTaxId").Value.ToString()
                    };
                    clients.Add(client);
                }

                var item = new Item
                {
                    ItemCode = recordset.Fields.Item("ItemCode").Value.ToString(),
                    ItemName = recordset.Fields.Item("ItemName").Value.ToString(),
                    ItemGroup = recordset.Fields.Item("ItemGroup").Value.ToString()
                };

                recordset.MoveNext();
            }
        }
        finally
        {
            Utilities.Release(recordset);
        }
        return clients;
    }
}*/
