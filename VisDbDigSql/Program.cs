using System;
using System.Data.OleDb;
using System.Linq;

namespace VisDbDig.Sql
{
    class Program
    {
        static void Main(string[] args)
        {
            string db = args[0];
            string connStr = $"Provider=sqloledb;Data Source=.;Initial Catalog={db};Integrated Security = SSPI;";

            using (var oledb = new OleDbConnection("Provider=sqloledb;" + connStr))
            {
                oledb.Open();

                var tables = oledb.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, null)
                    .Select()
                    .GroupBy(col => col["TABLE_NAME"].ToString())
                    .ToDictionary(
                        g => g.Key,
                        g => g.Select(col =>
                            new Field
                            {
                                InternalName = col["COLUMN_NAME"].ToString(),
                                DataType = col["DATA_TYPE"].ToString()
                            }
                        ).ToList()
                    );

                Console.WriteLine(JsonConvert.SerializeObject(tables));

                var relationships = oledb.GetOleDbSchemaTable(OleDbSchemaGuid.Foreign_Keys, null)
                    .Select()
                    .Select(fk => new Relationship
                    {
                        From = fk["FK_TABLE_NAME"].ToString(),
                        To = fk["PK_TABLE_NAME"].ToString()
                    })
                    .ToList();

                Console.WriteLine(JsonConvert.SerializeObject(relationships));
            }
        }
    }
}
