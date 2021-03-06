﻿using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using VisDbDig.Model;

namespace VisDbDig.Sql
{
    // Export the schema of a SQL Server DB to JSON files
    // See readme.md
    class Program
    {
        // Usage
        //    VisDbDig.sql.exe [OLEDB Connection String] [Output Directory]
        // Example:
        //    VisDbDig.sql.exe "Provider=sqloledb;Data Source=.;Initial Catalog=WideWorldImporters;Integrated Security=SSPI;" "C:\temp\WideWorldImporters"
        //
        static void Main(string[] args)
        {
            string connStr = args[0];
            string outputPath = args[1];

            if (!Directory.Exists(outputPath))
                Directory.CreateDirectory(outputPath);

            using (var oledb = new OleDbConnection(connStr))
            {
                oledb.Open();

                // Read the table definitions.
                var tables = oledb.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, null)
                    .Select()
                    .Where(col => col["TABLE_SCHEMA"].ToString() != "sys")
                    .GroupBy(col => col["TABLE_NAME"].ToString())
                    .ToDictionary(
                        g => g.Key,
                        g => g.Select(col =>
                            new Field
                            {
                                Name = col["COLUMN_NAME"].ToString(),
                                DataType = OleDbDataType(col["DATA_TYPE"], col["CHARACTER_MAXIMUM_LENGTH"])
                            }
                        ).ToList()
                    );

                // Read the relationsips between tables.
                var relationships = oledb.GetOleDbSchemaTable(OleDbSchemaGuid.Foreign_Keys, null)
                    .Select()
                    .Select(fk => new Relationship
                    {
                        From = fk["FK_TABLE_NAME"].ToString(),
                        To = fk["PK_TABLE_NAME"].ToString()
                    })
                    .ToList();

                var tableNames = tables.Keys.ToList();

                // Write the files.
                File.WriteAllText(Path.Combine(outputPath, "types.json"), JsonConvert.SerializeObject(tables));
                File.WriteAllText(Path.Combine(outputPath, "relationships.json"), JsonConvert.SerializeObject(relationships));
                File.WriteAllText(Path.Combine(outputPath, "typenames.json"), JsonConvert.SerializeObject(tableNames));
            }
        }

        // Turn the OLE DB table defination into a SQL data type.
        private static string OleDbDataType(object oledbType, object maxLength)
        {
            string dataType = OleDbTypeNames[(int) oledbType];
            string len = maxLength?.ToString();

            if (string.IsNullOrEmpty(len))
                return dataType;

            if (len == "1073741823" || len == "2147483647")
                len = "MAX";

            return $"{dataType}[{len}]";
        }

        static Dictionary<int, string> OleDbTypeNames = new Dictionary<int, string>
        {
            { (int) OleDbType.BigInt, "bigint" },
            { (int) OleDbType.Binary, "binary" },
            { (int) OleDbType.Boolean, "bit" },
            { (int) OleDbType.BSTR, "nvarchar" },
            { (int) OleDbType.Char, "char" },
            { (int) OleDbType.Currency, "money" },
            { (int) OleDbType.Date, "datetime" },
            { (int) OleDbType.DBDate, "datetime" },
            { (int) OleDbType.DBTime, "datetime" },
            { (int) OleDbType.DBTimeStamp, "datetime" },
            { (int) OleDbType.Decimal, "decimal" },
            { (int) OleDbType.Double, "float" },
            { (int) OleDbType.Empty, "empty" },
            { (int) OleDbType.Error, "error" },
            { (int) OleDbType.Filetime, "filetime" },
            { (int) OleDbType.Guid, "uniqueidentifier" },
            { (int) OleDbType.IDispatch, "idispatch" },
            { (int) OleDbType.Integer, "identity" },
            { (int) OleDbType.IUnknown, "iunknown" },
            { (int) OleDbType.LongVarBinary, "image" },
            { (int) OleDbType.LongVarChar, "text" },
            { (int) OleDbType.LongVarWChar, "text" },
            { (int) OleDbType.Numeric, "decima" },
            { (int) OleDbType.PropVariant, "propvariant" },
            { (int) OleDbType.Single, "real" },
            { (int) OleDbType.SmallInt, "smallint" },
            { (int) OleDbType.TinyInt, "tinyint" },
            { (int) OleDbType.UnsignedBigInt, "unsignedbigint" },
            { (int) OleDbType.UnsignedInt, "unsignedint" },
            { (int) OleDbType.UnsignedSmallInt, "unsignedsmallint" },
            { (int) OleDbType.UnsignedTinyInt, "tinyint" },
            { (int) OleDbType.VarBinary, "varbinary" },
            { (int) OleDbType.VarChar, "varchar" },
            { (int) OleDbType.Variant, "sql_variant" },
            { (int) OleDbType.VarNumeric, "varnumeric" },
            { (int) OleDbType.VarWChar, "nvarchar" },
            { (int) OleDbType.WChar, "nchar" },
        };
    }
} 