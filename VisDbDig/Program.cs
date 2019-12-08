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
    class Program
    {
        static void Main(string[] args)
        {
            ReadArgs(args, out string tablesFilepath, out string relationshipsFilepath, out string filterFilepath);

            var tables = LoadTables(tablesFilepath);
            tables = FilterTables(tables, filterFilepath);
            var relationships = LoadRelationships(relationshipsFilepath);

            MakeVisioPage(out var visio, out var doc, out var page);
            visio.UndoEnabled = false;
            SetVisioDeferWork(visio, true);

            var tableShapeIDs = DropMany("table", tables.Count(), page);
            SetupTableShapes(page, tables, tableShapeIDs, out var tableShapeGlueToCells);

            var fromToPairs = GetRelationshipFromToCells(relationships, tableShapeGlueToCells);

            var connectorShapeIDs = DropMany("Dynamic Connector", fromToPairs.Count, page);

            ConnectShapes(page, connectorShapeIDs, fromToPairs);

            SetVisioDeferWork(visio, false);
            page.Layout();
            const int visFitPage = 1;
            visio.ActiveWindow.ViewFit = visFitPage;
            visio.UndoEnabled = true;
        }

        private static Dictionary<string,List<Field>> LoadTables(string filepath)
        {
            var json = File.ReadAllText(filepath);
            return JsonConvert.DeserializeObject<Dictionary<string, List<Field>>>(json);
        }

        private static List<Relationship> LoadRelationships(string filepath)
        {
            var json = File.ReadAllText(filepath);
            return JsonConvert.DeserializeObject<List<Relationship>>(json);
        }

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

        private static void MakeVisioPage(out Vis.Application visio, out Vis.Document doc, out Vis.Page page)
        {
            string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "VisDbDig.vstx");

            visio = new Vis.Application();
            doc = visio.Documents.Add(templatePath);
            page = doc.Pages[1];
        }

        private static List<short> DropMany(string masterName, int count, Vis.Page page)
        {
            var master = page.Document.Masters.ItemU[masterName];
            var masters = Enumerable.Repeat(master, count).ToArray<object>();
            var xys = Enumerable.Repeat(0.0, count * 2).ToArray();
            page.DropMany(masters, xys, out var ids);
            return ids.OfType<short>().ToList();
        }

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

        private static void SetVisioDeferWork(Vis.Application visio, bool value)
        {
            short shortValue = (short) (value ? 0 : 1);
            visio.DeferRecalc = shortValue;
            visio.ScreenUpdating = shortValue;
            visio.LiveDynamics = value;
            visio.AutoLayout = value;
        }

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

    public static class Ext
    {
        public static Vis.Cell Cell(this Vis.Shape shape, Vis.VisSectionIndices section, Vis.VisRowIndices row, Vis.VisCellIndices cell)
        {
            return shape.CellsSRC[(short)section, (short)row, (short)cell];
        }
    }
}
