
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
using System.Linq;
using Newtonsoft.Json.Linq;
using NetTopologySuite.Geometries.Prepared;

class City
{
    public string Name { get; set; }
    public string Type { get; set; }
    public double Latitude { get; set; }
    public double Longitude { get; set; }
    public double Distance { get; set; }
    public int Population { get; set; }
    public Geometry Geometry => new Point(Longitude, Latitude) { SRID = 4326 };
}

class Program
{
    private static readonly HttpClient client = new HttpClient();
    private static Geometry russianBorder;
    private static IPreparedGeometry preparedRussianBorder;

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

        // Подготовка геометрии границы России
        preparedRussianBorder = PreparedGeometryFactory.Prepare(russianBorder);

        string inputPath = @"C:\Users\User\Desktop\grid.xlsx"; // Путь к основному файлу
        string outputPath = @"C:\Users\User\Desktop\nearby_cities.xlsx";
        //string excludedCitiesPath = @"C:\Users\User\Desktop\excluded_cities.xlsx"; //Больше не нужен

        List<City> cities = ReadCitiesFromExcel(inputPath);
        List<string> excludedCityNames = ReadExcludedCityNamesFromExcel(inputPath); // Используем тот же путь, что и для городов

        List<City> filteredCities = FilterCities(cities);
        Console.WriteLine($"После фильтрации по границам осталось городов: {filteredCities.Count}");

        List<List<City>> nearbyCities = await FindNearbyCitiesForAll(filteredCities, 100, excludedCityNames);
        WriteNearbyCitiesToExcel(filteredCities, nearbyCities, outputPath);

        Console.WriteLine("Обработка завершена. Результаты сохранены в " + outputPath);
    }

    static List<City> ReadCitiesFromExcel(string filePath)
    {
        var ci = new CultureInfo("ru-RU");
        ci.NumberFormat.NumberDecimalSeparator = ".";

        List<City> cities = new List<City>();

        try
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0]; // Первый лист для основных городов
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    try
                    {
                        cities.Add(new City
                        {
                            Longitude = ParseDouble(worksheet.Cells[row, 1].Text, ci),
                            Latitude = ParseDouble(worksheet.Cells[row, 2].Text, ci),
                            Name = worksheet.Cells[row, 3].Text,
                            Population = ParseInt(worksheet.Cells[row, 4].Text)
                        });
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Ошибка при чтении строки {row}: {ex.Message}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при чтении файла Excel: {ex.Message}");
        }

        return cities;
    }


    static List<string> ReadExcludedCityNamesFromExcel(string filePath)
    {
        List<string> excludedCities = new List<string>();

        try
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[1]; // Второй лист для исключенных городов
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 1; row <= rowCount; row++)
                {
                    string cityName = worksheet.Cells[row, 1].Text?.Trim();
                    if (!string.IsNullOrEmpty(cityName))
                    {
                        excludedCities.Add(cityName);
                    }
                }
            }
            Console.WriteLine($"Прочитано {excludedCities.Count} исключенных городов.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка чтения файла исключенных городов: {ex.Message}");
        }

        return excludedCities;
    }

    static double NormalizeLongitude(double lon)
    {
        while (lon > 180) lon -= 360;
        while (lon < -180) lon += 360;
        return lon;
    }

    static List<City> FilterCities(List<City> cities)
    {
        var filtered = new List<City>();
        var geometryFactory = NtsGeometryServices.Instance.CreateGeometryFactory(4326);

        foreach (var city in cities)
        {
            try
            {
                // Нормализация долготы
                double normalizedLongitude = NormalizeLongitude(city.Longitude);

                // Создаем геометрию города с нормализованной долготой
                Geometry cityGeometry = new Point(normalizedLongitude, city.Latitude) { SRID = 4326 };

                // Буферизация города (минимальный буфер для начала)
                double bufferSize = 0.0001;
                Geometry bufferedCity = cityGeometry.Buffer(bufferSize);
                bufferedCity.SRID = 4326;

                // Проверка пересечения с использованием PreparedGeometry
                bool intersects = preparedRussianBorder.Intersects(bufferedCity);

                // **Обходной путь для потенциального разрыва границы:**
                if (!intersects && Math.Abs(normalizedLongitude) > 170) // Если долгота близка к 180 меридиану
                {
                    // Создаем альтернативную геометрию, смещенную на 0.0001 градуса (примерно 10 метров)
                    Geometry alternativeCityGeometry = new Point(normalizedLongitude + 0.0001, city.Latitude) { SRID = 4326 };
                    Geometry alternativeBufferedCity = alternativeCityGeometry.Buffer(bufferSize);
                    alternativeBufferedCity.SRID = 4326;
                    intersects = preparedRussianBorder.Intersects(alternativeBufferedCity);

                    if (!intersects) //проверяем в другую сторону
                    {
                        alternativeCityGeometry = new Point(normalizedLongitude - 0.0001, city.Latitude) { SRID = 4326 };
                        alternativeBufferedCity = alternativeCityGeometry.Buffer(bufferSize);
                        alternativeBufferedCity.SRID = 4326;
                        intersects = preparedRussianBorder.Intersects(alternativeBufferedCity);
                    }
                }

                if (intersects)
                {
                    filtered.Add(city);
                    Console.WriteLine($"Город {city.Name} ({city.Longitude}, {city.Latitude}) добавлен.");
                }
                else
                {
                    Console.WriteLine($"Город {city.Name} ({city.Longitude}, {city.Latitude}) исключен.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при обработке города {city.Name}: {ex.Message}");
            }
        }

        return filtered;
    }

    static async Task<List<List<City>>> FindNearbyCitiesForAll(List<City> cities, int radiusKm, List<string> excludedCityNames)
    {
        List<List<City>> allNearbyCities = new List<List<City>>();
        foreach (var city in cities)
        {
            Console.WriteLine($"\nОбработка города: {city.Name}");
            var nearby = await FindCitiesInRadius(city.Latitude, city.Longitude, radiusKm, excludedCityNames, city); // Передаем исходный город!
            allNearbyCities.Add(nearby);
            await Task.Delay(1000);
        }
        return allNearbyCities;
    }

    static async Task<List<City>> FindCitiesInRadius(double lat, double lon, int radiusKm, List<string> excludedCityNames, City originalCity)
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
            response.EnsureSuccessStatusCode();

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

                    var point = new Point(elementLon, elementLat) { SRID = 4326 };
                    if (!russianBorder.Contains(point)) continue;

                    string cityName = node.SelectSingleNode("tag[@k='name']")?.Attributes["v"]?.Value ?? "Unknown";

                    if (!IsRussianName(cityName)) continue;

                    if (excludedCityNames.Contains(cityName))
                    {
                        Console.WriteLine($"Город {cityName} исключен из результатов.");
                        continue;
                    }

                    // Проверка на то, что это не тот же самый город
                    if (cityName == originalCity.Name)
                    {
                        Console.WriteLine($"Город {cityName} пропущен, так как это исходный город.");
                        continue;
                    }

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
        catch (HttpRequestException e)
        {
            Console.WriteLine($"Ошибка HTTP запроса: {e.Message}");
            return new List<City>();
        }
        catch (Exception e)
        {
            Console.WriteLine($"Ошибка запроса: {e.Message}");
            return new List<City>();
        }
    }

    static async Task<Geometry> GetRussianBorder()
    {
        Console.WriteLine("Получение границы России...");

        string geoJsonUrl = "https://raw.githubusercontent.com/johan/world.geo.json/master/countries/RUS.geo.json";

        try
        {
            using (var client = new HttpClient())
            {
                var response = await client.GetStringAsync(geoJsonUrl);
                Console.WriteLine("GeoJSON получен успешно.");

                if (string.IsNullOrEmpty(response))
                {
                    throw new Exception("Получен пустой ответ от сервера.");
                }

                var reader = new GeoJsonReader();
                var featureCollection = reader.Read<NetTopologySuite.Features.FeatureCollection>(response);

                if (featureCollection == null || featureCollection.Count == 0)
                {
                    throw new Exception("Не удалось прочитать FeatureCollection из GeoJSON.");
                }

                var feature = featureCollection[0];
                var geometry = feature.Geometry;

                if (geometry == null)
                {
                    throw new Exception("Геометрия в Feature отсутствует.");
                }

                Console.WriteLine($"Граница России получена успешно. Тип геометрии: {geometry.GeometryType}");
                return geometry;
            }
        }
        catch (Exception e)
        {
            Console.WriteLine($"Ошибка при получении границы России: {e.Message}");
            Console.WriteLine($"Stack Trace: {e.StackTrace}");
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

    static double ParseDouble(string value, CultureInfo ci)
    {
        if (string.IsNullOrEmpty(value)) return 0;
        return double.TryParse(value, NumberStyles.Any, ci, out double result) ? result : 0;
    }

    static int ParseInt(string value)
    {
        if (string.IsNullOrEmpty(value)) return 0;
        return int.TryParse(value, out int result) ? result : 0;
    }

    static bool IsRussianName(string name)
    {
        return Regex.IsMatch(name, @"^[А-Яа-я\s-]+$");
    }
}
