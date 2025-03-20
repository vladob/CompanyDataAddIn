using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

public class ClassificationManager
{
    private static readonly HttpClient httpClient = new HttpClient();
    // Adjust the base folder as needed, for example using Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
    private readonly string classificationsFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Classifications");

    public Dictionary<string, Dictionary<string, string>> Lookups = new Dictionary<string, Dictionary<string, string>>();
    public ClassificationManager()
    {
        if (!Directory.Exists(classificationsFolder))
        {
            Directory.CreateDirectory(classificationsFolder);
        }
    }

    /// <summary>
    /// Checks if the specified JSON file exists. If not, downloads from the API, saves it, and returns the data.
    /// </summary>
    /// <param name="fileName">File name (e.g., "LegalForms.json")</param>
    /// <param name="apiUrl">The URL for the API</param>
    /// <returns>A JSON token representing the classification data.</returns>
    public async Task<JToken> GetOrDownloadClassificationAsync(string fileName, string apiUrl)
    {
        string filePath = Path.Combine(classificationsFolder, fileName);
        if (!File.Exists(filePath))
        {
            // Download data from API
            string json = await httpClient.GetStringAsync(apiUrl);
            // Optionally, add error handling here
            File.WriteAllText(filePath, json);
            return JToken.Parse(json);
        }
        else
        {
            // Read from file and parse
            string json = File.ReadAllText(filePath);
            return JToken.Parse(json);
        }
    }

    /// <summary>
    /// For classifications where data comes from multiple API endpoints (e.g. Locations), this method downloads
    /// from all endpoints, merges the arrays, saves the merged data, and returns it.
    /// </summary>
    /// <param name="fileName">File name (e.g., "Locations.json")</param>
    /// <param name="apiUrls">List of API endpoints</param>
    /// <returns>A JSON token representing the merged classification data.</returns>
    public async Task<JToken> GetOrDownloadCombinedClassificationAsync(string fileName, List<string> apiUrls)
    {
        string filePath = Path.Combine(classificationsFolder, fileName);
        if (!File.Exists(filePath))
        {
            JArray combinedData = new JArray();
            foreach (var url in apiUrls)
            {
                string json = await httpClient.GetStringAsync(url);
                // Assuming each API returns an array of items.
                JArray items = JArray.Parse(json);
                combinedData.Merge(items);
            }
            // Save the merged data to file
            File.WriteAllText(filePath, combinedData.ToString());
            return combinedData;
        }
        else
        {
            string json = File.ReadAllText(filePath);
            return JToken.Parse(json);
        }
    }

    /// <summary>
    /// Initializes all classifications by checking for file presence and downloading any missing data.
    /// You can call this method once at Add-In startup.
    /// </summary>
    public async Task InitializeAllClassificationsAsync()
    {
        // Example: LegalForms classification from a single API call
        JToken legalFormsToken = await GetOrDownloadClassificationAsync("LegalForms.json", "https://www.registeruz.sk/cruz-public/api/pravne-formy");
        JToken organizationSizesToken = await GetOrDownloadClassificationAsync("OrganizationSizes.json", "https://www.registeruz.sk/cruz-public/api/velkosti-organizacie");
        JToken ownershipTypesToken = await GetOrDownloadClassificationAsync("OwnershipTypes.json", "https://www.registeruz.sk/cruz-public/api/druhy-vlastnictva");
        JToken skNaceToken = await GetOrDownloadClassificationAsync("SkNace.json", "https://www.registeruz.sk/cruz-public/api/sk-nace");
        JToken regionsToken = await GetOrDownloadClassificationAsync("Regions.json", "https://www.registeruz.sk/cruz-public/api/kraje");
        JToken distinctsToken = await GetOrDownloadClassificationAsync("Distinc.json", "https://www.registeruz.sk/cruz-public/api/okresy");
        JToken muncipalitiesToken = await GetOrDownloadClassificationAsync("Muncipalities.json", "https://www.registeruz.sk/cruz-public/api/sidla");

        Lookups.Add("LegalForms", ConvertClassificationData(legalFormsToken["klasifikacie"], "sk"));
        Lookups.Add("OrganizationSizes", ConvertClassificationData(organizationSizesToken["klasifikacie"], "sk"));
        Lookups.Add("OwnershipTypes", ConvertClassificationData(ownershipTypesToken["klasifikacie"], "sk"));
        Lookups.Add("SkNace", ConvertClassificationData(skNaceToken["klasifikacie"], "sk"));
        Lookups.Add("Regions", ConvertClassificationData(regionsToken["lokacie"], "sk"));
        Lookups.Add("Distincts", ConvertClassificationData(distinctsToken["lokacie"], "sk"));
        Lookups.Add("Muncipalities", ConvertClassificationData(muncipalitiesToken["lokacie"], "sk"));
    }

    public Dictionary<string, string> ConvertClassificationData(JToken classificationData, string language)
    {
        var dict = new Dictionary<string, string>();

        // Assume classificationData is a JArray of items.
        foreach (JToken item in classificationData)
        {
            string code = item["kod"]?.ToString();
            string title = item["nazov"]?[language]?.ToString();
            if (!string.IsNullOrEmpty(code) && title != null)
            {
                dict[code] = title;
            }
        }
        return dict;
    }

    // In-memory storage of the classifications for later use.
    public Dictionary<string, JToken> _classifications { get; private set; }
}
