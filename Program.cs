using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading.Tasks;
using System.Xml;
using OfficeOpenXml;
using System.IO;
using System.Globalization;
using System.Text;
using NetTopologySuite.Geometries;
using NetTopologySuite.IO;
using NetTopologySuite;
using System.Text.RegularExpressions;

class City
{
    public string Name { get; set; }
    public string Type { get; set; }
    public double Latitude { get; set; }
    public double Longitude { get; set; }
    public double Distance { get; set; }
    public int Population { get; set; }
    public Geometry Geometry => new Point(new Coordinate(Longitude, Latitude)) { SRID = 4326 };
}

class Program
{
    private static readonly HttpClient client = new HttpClient();
    private static Geometry russianBorder;

    static async Task Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // Загрузка геоданных России
        russianBorder = await GetRussianBorder();
        if (russianBorder == null)
        {
            Console.WriteLine("Не удалось загрузить границы России. Программа завершает работу.");
            return;
        }

        string inputPath = @"C:\Users\User\Desktop\grid.xlsx";
        string outputPath = @"C:\Users\User\Desktop\nearby_cities.xlsx";

        List<City> cities = ReadCitiesFromExcel(inputPath);
        List<City> filteredCities = FilterCities(cities);
        Console.WriteLine($"После фильтрации осталось городов: {filteredCities.Count}");

        List<List<City>> nearbyCities = await FindNearbyCitiesForAll(filteredCities, 100);
        WriteNearbyCitiesToExcel(filteredCities, nearbyCities, outputPath);

        Console.WriteLine("Обработка завершена. Результаты сохранены в " + outputPath);
    }

    static List<City> ReadCitiesFromExcel(string filePath)
    {
        var ci = new CultureInfo("ru-RU");
        ci.NumberFormat.NumberDecimalSeparator = ".";

        List<City> cities = new List<City>();

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension.Rows;

            for (int row = 2; row <= rowCount; row++)
            {
                cities.Add(new City
                {
                    Longitude = ParseDouble(worksheet.Cells[row, 1].Text, ci),
                    Latitude = ParseDouble(worksheet.Cells[row, 2].Text, ci),
                    Name = worksheet.Cells[row, 3].Text,
                    Population = ParseInt(worksheet.Cells[row, 4].Text)
                });
            }
        }

        return cities;
    }

    static List<City> FilterCities(List<City> cities)
    {
        var filtered = new List<City>();
        foreach (var city in cities)
        {
            if (russianBorder.Contains(city.Geometry))
            {
                filtered.Add(city);
            }
            else
            {
                Console.WriteLine($"Город {city.Name} не находится в пределах границ России и будет исключен.");
            }
        }
        return filtered;
    }

    static async Task<List<List<City>>> FindNearbyCitiesForAll(List<City> cities, int radiusKm)
    {
        List<List<City>> allNearbyCities = new List<List<City>>();
        foreach (var city in cities)
        {
            Console.WriteLine($"\nОбработка города: {city.Name}");
            var nearby = await FindCitiesInRadius(city.Latitude, city.Longitude, radiusKm);
            allNearbyCities.Add(nearby);
            await Task.Delay(1000);
        }
        return allNearbyCities;
    }

    static async Task<List<City>> FindCitiesInRadius(double lat, double lon, int radiusKm)
    {
        string overpassQuery = $@"
<osm-script>
  <query type=""node"">
    <around lat=""{lat}"" lon=""{lon}"" radius=""{radiusKm * 1000}""/>
    <has-kv k=""place"" regv=""city|town|village""/>
  </query>
  <print mode=""body""/>
  <recurse type=""down""/>
  <print mode=""skeleton"" order=""quadtile""/>
</osm-script>";

        try
        {
            using var content = new StringContent(overpassQuery, Encoding.UTF8, "application/xml");
            var response = await client.PostAsync("https://overpass-api.de/api/interpreter", content);
            var responseContent = await response.Content.ReadAsStringAsync();

            var cities = new List<City>();
            var xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(responseContent);

            foreach (XmlNode node in xmlDoc.SelectNodes("//node"))
            {
                try
                {
                    double elementLat = double.Parse(node.Attributes["lat"].Value, CultureInfo.InvariantCulture);
                    double elementLon = double.Parse(node.Attributes["lon"].Value, CultureInfo.InvariantCulture);

                    var point = new Point(new Coordinate(elementLon, elementLat)) { SRID = 4326 };
                    if (!russianBorder.Contains(point)) continue;

                    string cityName = node.SelectSingleNode("tag[@k='name']")?.Attributes["v"]?.Value ?? "Unknown";

                    // Проверка на русское название города
                    if (!IsRussianName(cityName)) continue;

                    var city = new City
                    {
                        Name = cityName,
                        Type = node.SelectSingleNode("tag[@k='place']")?.Attributes["v"]?.Value ?? "Unknown",
                        Latitude = elementLat,
                        Longitude = elementLon,
                        Distance = CalculateDistance(lat, lon, elementLat, elementLon),
                        Population = int.TryParse(
                            node.SelectSingleNode("tag[@k='population']")?.Attributes["v"]?.Value,
                            out int pop) ? pop : 0
                    };

                    cities.Add(city);

                    Console.WriteLine($"Найден: {city.Name.PadRight(20)} | " +
                                      $"Тип: {city.Type.PadRight(10)} | " +
                                      $"Население: {city.Population.ToString().PadRight(8)} | " +
                                      $"Расстояние: {Math.Round(city.Distance, 2)} км");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ошибка обработки узла: {ex.Message}");
                }
            }

            Console.WriteLine($"Всего найдено городов: {cities.Count}");
            return cities;
        }
        catch (Exception e)
        {
            Console.WriteLine($"Ошибка запроса: {e.Message}");
            return new List<City>();
        }
    }

    static async Task<Geometry> GetRussianBorder()
    {
        string geoJsonUrl = "https://raw.githubusercontent.com/johan/world.geo.json/master/countries/RUS.geo.json";

        try
        {
            var response = await client.GetStringAsync(geoJsonUrl);
            var reader = new GeoJsonReader();
            var featureCollection = reader.Read<NetTopologySuite.Features.FeatureCollection>(response);

            var geometry = featureCollection[0].Geometry;
            geometry.SRID = 4326;

            Console.WriteLine($"Границы России успешно загружены. Тип: {geometry.GeometryType}");
            return geometry;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка загрузки границ: {ex.Message}");
            return null;
        }
    }

    static void WriteNearbyCitiesToExcel(List<City> originalCities, List<List<City>> nearbyCities, string filePath)
    {
        using var package = new ExcelPackage();
        var worksheet = package.Workbook.Worksheets.Add("Результаты");

        string[] headers = {
            "Исх. долгота", "Исх. широта", "Исх. название", "Исх. население",
            "Ближ. название", "Ближ. долгота", "Ближ. широта", "Тип", "Население", "Расстояние (км)"
        };

        for (int i = 0; i < headers.Length; i++)
            worksheet.Cells[1, i + 1].Value = headers[i];

        int row = 2;
        for (int i = 0; i < originalCities.Count; i++)
        {
            var orig = originalCities[i];
            var nearby = nearbyCities[i];

            foreach (var city in nearby)
            {
                worksheet.Cells[row, 1].Value = orig.Longitude;
                worksheet.Cells[row, 2].Value = orig.Latitude;
                worksheet.Cells[row, 3].Value = orig.Name;
                worksheet.Cells[row, 4].Value = orig.Population;
                worksheet.Cells[row, 5].Value = city.Name;
                worksheet.Cells[row, 6].Value = city.Longitude;
                worksheet.Cells[row, 7].Value = city.Latitude;
                worksheet.Cells[row, 8].Value = city.Type;
                worksheet.Cells[row, 9].Value = city.Population;
                worksheet.Cells[row, 10].Value = Math.Round(city.Distance, 2);
                row++;
            }
        }

        var numberFormat = "0.000000";
        worksheet.Column(1).Style.Numberformat.Format = numberFormat;
        worksheet.Column(2).Style.Numberformat.Format = numberFormat;
        worksheet.Column(6).Style.Numberformat.Format = numberFormat;
        worksheet.Column(7).Style.Numberformat.Format = numberFormat;
        worksheet.Column(10).Style.Numberformat.Format = "0.00";

        worksheet.Cells.AutoFitColumns();
        package.SaveAs(new FileInfo(filePath));
    }

    static double CalculateDistance(double lat1, double lon1, double lat2, double lon2)
    {
        var d1 = lat1 * (Math.PI / 180.0);
        var num1 = lon1 * (Math.PI / 180.0);
        var d2 = lat2 * (Math.PI / 180.0);
        var num2 = lon2 * (Math.PI / 180.0) - num1;
        var d3 = Math.Pow(Math.Sin((d2 - d1) / 2.0), 2.0) +
                Math.Cos(d1) * Math.Cos(d2) * Math.Pow(Math.Sin(num2 / 2.0), 2.0);

        return 6376500.0 * (2.0 * Math.Atan2(Math.Sqrt(d3), Math.Sqrt(1.0 - d3))) / 1000;
    }

    static double ParseDouble(string value, CultureInfo ci) =>
        double.TryParse(value, NumberStyles.Any, ci, out double result) ? result : 0;

    static int ParseInt(string value) =>
        int.TryParse(value, out int result) ? result : 0;

    // Проверка на русское название города
    static bool IsRussianName(string name)
    {
        // Регулярное выражение для проверки кириллических символов, пробелов и дефисов
        return Regex.IsMatch(name, @"^[А-Яа-я\s-]+$");
    }
}