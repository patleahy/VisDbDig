using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using VisDbDig.Model;
using Vis = Microsoft.Office.Interop.Visio;

namespace VisDbDig
{
    // This command creates a Visio diagram using the output of VisDbDig.sql.exe or another tool which creates the same
    // file format. See readme.md
    class Program
    {
        static void Main(string[] args)
        {
            // Reading the schema definition from the input files.
            ReadArgs(args, out string tablesFilepath, out string relationshipsFilepath, out string filterFilepath);
            var tables = LoadTables(tablesFilepath);
            tables = FilterTables(tables, filterFilepath);
            var relationships = LoadRelationships(relationshipsFilepath);

            // Creating a document based on the Visio template VisDbDig.vstx.
            // That template contains a Master shape which was created to display the table definitions.
            MakeVisioDoc(out var visio, out var doc, out var page);
            visio.UndoEnabled = false;
            SetVisioDeferWork(visio, true);

            // Dropping one table shape on the page for each table definition. 
            // Its faster to drop all the shapes at once.
            // tableShapeIDs is the ID of each shape on the page.
            var tableShapeIDs = DropMany("table", tables.Count(), page);

            // Enter text into the each table shape to describe each table.
            // The method also remembers which Visio cell on each table shape the connectors should glue to. These
            // cells are returned in tableShapeGlueToCells
            SetupTableShapes(page, tables, tableShapeIDs, out var tableShapeGlueToCells);

            // For each relationsip figure out which where the beginning and end of each connector should glue to.
            var fromToPairs = GetRelationshipFromToCells(relationships, tableShapeGlueToCells);

            // Drop one Dynamic Connector shape on the page for each relatonship in the schema,
            var connectorShapeIDs = DropMany("Dynamic Connector", fromToPairs.Count, page);

            // Connect each connector between the correct two table shapes to represent each relationship.
            ConnectShapes(page, connectorShapeIDs, fromToPairs);

            // Have Visio layout the shapes on the page and resize the page to fit the layout.
            SetVisioDeferWork(visio, false);
            page.Layout();
            const int visFitPage = 1;
            visio.ActiveWindow.ViewFit = visFitPage;
            visio.UndoEnabled = true;
        }

        // Load the table definitions from the table file. 
        // The key the dictionary is the table names.
        // The value of the dictionary is the definitions of the fields in a table. 
        private static Dictionary<string,List<Field>> LoadTables(string filepath)
        {
            var json = File.ReadAllText(filepath);
            return JsonConvert.DeserializeObject<Dictionary<string, List<Field>>>(json);
        }

        // Load the relationships between tables from the relationsip file.
        private static List<Relationship> LoadRelationships(string filepath)
        {
            var json = File.ReadAllText(filepath);
            return JsonConvert.DeserializeObject<List<Relationship>>(json);
        }

        // The user can specify a file that contains a list of tables. If they do that then we can't filter the list 
        // of tables that were loaded by LoadTables using the list.
        private static Dictionary<string, List<Field>> FilterTables(Dictionary<string,List<Field>> tables, string filterFilepath)
        {
            if (string.IsNullOrEmpty(filterFilepath))
                return tables;
            
            var tablesToKeep = File.ReadAllLines(filterFilepath);
            var ret = new Dictionary<string,List<Field>>();
            foreach (var name in tablesToKeep)
            {
                ret[name] = tables[name];
            }
            return ret;
        }

        // Create Visio document based on the VisDbDig template. Return the document and  first page in the document.
        private static void MakeVisioDoc(out Vis.Application visio, out Vis.Document doc, out Vis.Page page)
        {
            string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "VisDbDig.vstx");

            visio = new Vis.Application();
            doc = visio.Documents.Add(templatePath);
            page = doc.Pages[1];
        }

        // Drop many instances of a shape master. The shapes will be dropped on the specificed shape.
        // The master must be in the document of the page specified.
        // Return the IDs of the new shapes.
        private static List<short> DropMany(string masterName, int count, Vis.Page page)
        {
            var master = page.Document.Masters.ItemU[masterName];
            var masters = Enumerable.Repeat(master, count).ToArray<object>();
            var xys = Enumerable.Repeat(0.0, count * 2).ToArray();
            page.DropMany(masters, xys, out var ids);
            return ids.OfType<short>().ToList();
        }

        // Set the text on the table shaopes based on the definition of the tables.
        // While we are iteratering accross the table shapes we keep track of the cells we can glue to. We return a 
        // dictionary which maps table names to the cell in the shape a connector can glue to.
        private static void SetupTableShapes(
            Vis.Page page, Dictionary<string,List<Field>> tables, List<short> tableShapeIDs, 
            out Dictionary<string, Vis.Cell> tableShapeGlueToCells)
        {
            tableShapeGlueToCells = new Dictionary<string, Vis.Cell>(tables.Count);
            int t = 0;
            foreach (var table in tables)
            {
                var shape = page.Shapes.ItemFromID[tableShapeIDs[t]];
                shape.Shapes.ItemU["Title"].Text = table.Key;
                shape.Shapes.ItemU["Fields"].Text = GetFieldsAsString(table.Value);
                tableShapeGlueToCells[table.Key] = shape.Cell(Vis.VisSectionIndices.visSectionObject, Vis.VisRowIndices.visRowXFormOut, Vis.VisCellIndices.visXFormPinX);
                Console.WriteLine($"T,{t}/{tables.Count},{table.Key}");
                t++;
            }
        }

        // Iterate accross all the relationships. Find the table shapes which correspond to the from and to tables of
        // the relationship. Because the user can filter which tables we draw the to or from table may not exist.
        // If both tables exist then add the to and from cell that the connector can glue to to a list we will return.
        private static List<Tuple<Vis.Cell, Vis.Cell>> GetRelationshipFromToCells(
            List<Relationship> relationships, 
            Dictionary<string, Vis.Cell> tableShapeGlueToCells)
        
        {
            var ret = new List<Tuple<Vis.Cell, Vis.Cell>>();
            foreach (var rel in relationships)
            {
                tableShapeGlueToCells.TryGetValue(rel.From, out var from);
                if (from == null)
                    continue;

                tableShapeGlueToCells.TryGetValue(rel.To, out var to);
                if (to == null)
                    continue;

                ret.Add(new Tuple<Vis.Cell, Vis.Cell>(from, to));
            }
            return ret;
        }

        // Use the connectors we added to the page to glue the tables together to make the relationsips on the page.
        static private void ConnectShapes(Vis.Page page, List<short> connectorShapeIDs, List<Tuple<Vis.Cell, Vis.Cell>> fromToPairs)
        {
            int c = 0;
            foreach(var fromTo in fromToPairs)
            {
                var connector = page.Shapes.ItemFromID[connectorShapeIDs[c]];
                connector.Cell(Vis.VisSectionIndices.visSectionObject, Vis.VisRowIndices.visRowXForm1D, Vis.VisCellIndices.vis1DBeginX).GlueTo(fromTo.Item1);
                connector.Cell(Vis.VisSectionIndices.visSectionObject, Vis.VisRowIndices.visRowXForm1D, Vis.VisCellIndices.vis1DEndX).GlueTo(fromTo.Item2);
                Console.WriteLine($"R,{c}/{fromToPairs.Count}");
                c++;
            }
        }

        // Get a field from the input data as a string.
        private static string GetFieldsAsString(IEnumerable<Field> fields)
        {
            var sb = new StringBuilder();
            string nl = "";
            foreach (var field in fields)
            {
                var set = field.OneToMany ? " *" : "";
                sb.Append($"{nl}{field.Name}\t: {field.DataType}{set}");
                nl = "\n";
            }
            return sb.ToString();
        }

        // Change Visio's settings to that creating the document is faster.
        private static void SetVisioDeferWork(Vis.Application visio, bool value)
        {
            short shortValue = (short) (value ? 0 : 1);
            visio.DeferRecalc = shortValue;
            visio.ScreenUpdating = shortValue;
            visio.LiveDynamics = value;
            visio.AutoLayout = value;
        }

        // Parse the command line arguments. See readme.md
        private static void ReadArgs(
            string[] args, 
            out string tablesFilepath, out string relationshipsFilepath, out string filterFilepath)
        {            
            if (args.Length == 1)
            {
                tablesFilepath = Path.Combine(args[0], "types.json");
                relationshipsFilepath = Path.Combine(args[0], "relationships.json");
                filterFilepath = null;
            }
            else
            {
                tablesFilepath = args[0];
                relationshipsFilepath = args[1];
                filterFilepath = args.Length > 2 ? args[2] : null;
            }
        }
    }

    // Useful extension methods.
    public static class Extensions
    {
        // Return a cell in a shape using its section, row, and cell index.
        public static Vis.Cell Cell(this Vis.Shape shape, Vis.VisSectionIndices section, Vis.VisRowIndices row, Vis.VisCellIndices cell)
        {
            return shape.CellsSRC[(short)section, (short)row, (short)cell];
        }
    }
}
