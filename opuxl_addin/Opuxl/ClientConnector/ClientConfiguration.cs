using System;
using System.Configuration;

namespace ClientConnector
{
    class ClientConfiguration
    {
        private static String GetConfigurationSetting(String key, String defaultValue) {
            String value = ConfigurationManager.AppSettings[key];

            if (value == null) {
                value = defaultValue;
            }

            return value;
        }

        public static String Host { get { return GetConfigurationSetting("host", "localhost"); } }
        public static int Port { get { return int.Parse(GetConfigurationSetting("port", "61379")); } }
    }
}
