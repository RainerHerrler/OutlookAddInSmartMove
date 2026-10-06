using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;

namespace SmartMove.OutlookAddIn
{
    internal static class SmartMoveConfiguration
    {
        private const string DefaultConfiguration = "{\r\n  \"ownAddresses\": []\r\n}\r\n";

        public static string ConfigurationPath => Path.Combine(
            StatisticsDatabase.DataDirectory,
            "settings.json");

        public static void EnsureFileExists()
        {
            string path = ConfigurationPath;
            if (File.Exists(path))
            {
                return;
            }

            Directory.CreateDirectory(StatisticsDatabase.DataDirectory);
            try
            {
                using (var stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.Read))
                using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
                {
                    writer.Write(DefaultConfiguration);
                }
            }
            catch (IOException) when (File.Exists(path))
            {
                // Another caller created the configuration at the same time.
            }
        }

        public static HashSet<string> LoadOwnAddresses()
        {
            EnsureFileExists();

            SettingsDocument document;
            try
            {
                string json = File.ReadAllText(ConfigurationPath, Encoding.UTF8);
                document = new JavaScriptSerializer().Deserialize<SettingsDocument>(json);
            }
            catch (Exception exception) when (
                exception is ArgumentException ||
                exception is InvalidOperationException)
            {
                throw new InvalidDataException(
                    "Die SmartMove-Konfiguration ist kein gültiges JSON: " + ConfigurationPath,
                    exception);
            }

            if (document == null)
            {
                throw new InvalidDataException(
                    "Die SmartMove-Konfiguration ist leer: " + ConfigurationPath);
            }

            var addresses = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string configuredAddress in document.ownAddresses ?? new string[0])
            {
                string address = AddressHeuristic.Normalize(configuredAddress);
                if (string.IsNullOrEmpty(address))
                {
                    continue;
                }

                int atIndex = address.IndexOf('@');
                if (atIndex <= 0 || atIndex == address.Length - 1 || address.IndexOfAny(new[] { ' ', '\t', '\r', '\n' }) >= 0)
                {
                    throw new InvalidDataException(
                        "Ungültige E-Mail-Adresse in der SmartMove-Konfiguration: " + configuredAddress);
                }

                addresses.Add(address);
            }

            return addresses;
        }

        private sealed class SettingsDocument
        {
            // Keep the JSON-facing property name identical to the documented file format.
            public string[] ownAddresses { get; set; }
        }
    }
}
