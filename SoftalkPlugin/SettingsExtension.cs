using System;
using System.Collections.Generic;
using System.Text;
using Yomiage.SDK.Settings;

namespace SoftalkPlugin
{
    internal static class SettingsExtension
    {
        public static string GetExePath(this SettingsBase settings)
        {
            var key = "ExePath";
            if (settings.Strings?.TryGetSetting(key, out var value) == true)
            {
                return value.Value;
            }
            return string.Empty;
        }

        public static bool SetExePath(this SettingsBase settings, string value)
        {
            var key = "ExePath";
            if (settings.Strings?.ContainsKey(key) == true)
            {
                settings.Strings[key].Value = value;
                return true;
            }
            return false;
        }

        public static string GetVoiceName(this SettingsBase settings)
        {
            var key = "VoiceName";
            if (settings.Strings?.TryGetSetting(key, out var value) == true)
            {
                return value.Value;
            }
            return string.Empty;
        }

        public static string GetPresetName(this SettingsBase settings)
        {
            var key = "PresetName";
            if (settings.Strings?.TryGetSetting(key, out var value) == true)
            {
                return value.Value;
            }
            return string.Empty;
        }

        public static bool GetAccentEnable(this SettingsBase settings)
        {
            var key = "AccentEnable";
            if (settings.Bools?.TryGetSetting(key, out var value) == true)
            {
                return value.Value;
            }
            return false;
        }
    }
}
