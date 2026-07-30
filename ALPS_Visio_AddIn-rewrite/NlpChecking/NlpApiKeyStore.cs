using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace ALPS_Visio_AddIn_rewrite.NlpChecking
{
    /// <summary>
    /// Stores the optional LLM API key encrypted for the current Windows user.
    /// </summary>
    internal sealed class NlpApiKeyStore
    {
        private readonly string filePath;

        public NlpApiKeyStore()
        {
            filePath = Path.Combine(
                Environment.GetFolderPath(
                    Environment.SpecialFolder.LocalApplicationData),
                "ALPS Visio Add-In",
                "nlp-api-key.dat");
        }

        public string Load()
        {
            try
            {
                if (!File.Exists(filePath))
                    return null;

                byte[] protectedBytes = File.ReadAllBytes(filePath);
                byte[] clearBytes = ProtectedData.Unprotect(
                    protectedBytes, null, DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(clearBytes);
            }
            catch (IOException)
            {
                return null;
            }
            catch (CryptographicException)
            {
                return null;
            }
        }

        public void Save(string apiKey)
        {
            apiKey = apiKey?.Trim();
            if (string.IsNullOrEmpty(apiKey))
            {
                Clear();
                return;
            }

            string directory = Path.GetDirectoryName(filePath);
            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            byte[] clearBytes = Encoding.UTF8.GetBytes(apiKey);
            byte[] protectedBytes = ProtectedData.Protect(
                clearBytes, null, DataProtectionScope.CurrentUser);
            File.WriteAllBytes(filePath, protectedBytes);
        }

        public void Clear()
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }
}
