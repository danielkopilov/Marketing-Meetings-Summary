using System.IO;
using System.Text.Json;

namespace Marketing_Meetings_Summary;

/// <summary>
/// Persisted settings for the ATP Generator tab. Stores the ATP Creator
/// executable folder and the preferred ATP save location so they survive
/// across sessions until the user changes them manually.
/// </summary>
public class AtpSettings
{
    /// <summary>Folder that contains the ATP Creator .exe.</summary>
    public string ExeFolderPath { get; set; } = "";

    /// <summary>Preferred location where generated ATP files are saved.</summary>
    public string SavePath { get; set; } = "";

    private static string SettingsFilePath
    {
        get
        {
            string dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "MarketingMeetingsSummary");
            Directory.CreateDirectory(dir);
            return Path.Combine(dir, "atp_settings.json");
        }
    }

    public static AtpSettings Load()
    {
        try
        {
            if (File.Exists(SettingsFilePath))
            {
                string json = File.ReadAllText(SettingsFilePath);
                var loaded = JsonSerializer.Deserialize<AtpSettings>(json);
                if (loaded != null)
                    return loaded;
            }
        }
        catch
        {
            // Ignore corrupt/unreadable settings and fall back to defaults.
        }
        return new AtpSettings();
    }

    public void Save()
    {
        try
        {
            string json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(SettingsFilePath, json);
        }
        catch
        {
            // Persistence is best-effort; ignore write failures.
        }
    }
}
