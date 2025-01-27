using DemoDIAPI;
using DocumentFormat.OpenXml.ExtendedProperties;
using SAPbobsCOM;
using System;
using System.Collections.Generic;

public static class RecordExists
{
    public static bool IsInDatabase(SAPbobsCOM.Company company, string tableName, string column, string data)
    {
        string query = $"""
            SELECT COUNT(*) as RecordCount
            FROM {tableName}
            WHERE {column} LIKE '%{data}%'
            """;

        var recordset = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
        try
        {
            recordset.DoQuery(query);
            recordset.MoveFirst();
            int count = Convert.ToInt32(recordset.Fields.Item("RecordCount").Value);
            return count > 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error checking database: {ex.Message}");
            return false;
        }
        finally
        {
            Utilities.Release(recordset);
        }
    }
}