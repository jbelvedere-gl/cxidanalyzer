using ClosedXML.Excel;

namespace cxidanalyzer
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                var filePath = Environment.GetEnvironmentVariable("DYNATRACE_FILE_PATH");

                if (!string.IsNullOrEmpty(filePath))
                {
                    var items = LoadToCollection<FileStructure>(filePath);

                    var duplicateds = FindDuplicateds(items);

                    foreach (var item in duplicateds)
                    {
                        Console.WriteLine($"Connection ID: {item}");
                    }

                    Console.WriteLine("Process end.");
                }
            }
            catch
            {
                throw;
            }
        }

        private static List<T> LoadToCollection<T>(string filePath) where T : new()
        {
            var collection = new List<T>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);

                var rows = worksheet.RowsUsed().Skip(1);

                foreach (var row in rows)
                {
                    var item = new T();
                    var properties = typeof(T).GetProperties();

                    for (int i = 0; i < properties.Length; i++)
                    {
                        var cellValue = row.Cell(i + 1).Value;

                        properties[i].SetValue(item, cellValue.ToString());
                    }

                    collection.Add(item);
                }
            }

            return collection;
        }

        private static List<string> FindDuplicateds(IEnumerable<FileStructure> items)
        {
            var connections = new List<string>();

            foreach (var item in items)
            {

                if (!string.IsNullOrEmpty(item.Cx1))
                {
                    connections.Add(item.Cx1);
                }

                if (!string.IsNullOrEmpty(item.Cx2))
                {
                    connections.Add(item.Cx2);
                }

                if (!string.IsNullOrEmpty(item.Cx3))
                {
                    connections.Add(item.Cx3);
                }

                if (!string.IsNullOrEmpty(item.Cx4))
                {
                    connections.Add(item.Cx4);
                }

                if (!string.IsNullOrEmpty(item.Cx5))
                {
                    connections.Add(item.Cx5);
                }

                if (!string.IsNullOrEmpty(item.Cx6))
                {
                    connections.Add(item.Cx6);
                }

                if (!string.IsNullOrEmpty(item.Cx7))
                {
                    connections.Add(item.Cx7);
                }

                if (!string.IsNullOrEmpty(item.Cx8))
                {
                    connections.Add(item.Cx8);
                }

                if (!string.IsNullOrEmpty(item.Cx9))
                {
                    connections.Add(item.Cx9);
                }

                if (!string.IsNullOrEmpty(item.Cx10))
                {
                    connections.Add(item.Cx10);
                }

                if (!string.IsNullOrEmpty(item.Cx11))
                {
                    connections.Add(item.Cx11);
                }

                if (!string.IsNullOrEmpty(item.Cx12))
                {
                    connections.Add(item.Cx12);
                }

                if (!string.IsNullOrEmpty(item.Cx13))
                {
                    connections.Add(item.Cx13);
                }
            }

            var duplicateds = new List<string>();

            foreach (var cx in connections)
            {
                if (connections.Where(x => cx.Equals(x)).Count() > 1)
                {
                    if (!duplicateds.Any(x => x.Equals(cx)))
                    {
                        duplicateds.Add(cx);
                    }
                }
            }

            return duplicateds;
        }
    }
}
