using System;
using System.IO;
using System.Collections.Generic;

namespace iLogicExternal
{
    /// <summary>
    /// Utility class for file operations to avoid namespace conflicts
    /// </summary>
    public static class FileUtils
    {
        /// <summary>
        /// Checks if the specified .ilogicignore file exists in the given folder
        /// </summary>
        /// <param name="folderPath">The folder to check</param>
        /// <param name="fileName">The ignore file name (default: .ilogicignore)</param>
        /// <returns>True if the file exists, false otherwise</returns>
        public static bool HasIgnoreFile(string folderPath, string fileName = ".ilogicignore")
        {
            if (string.IsNullOrEmpty(folderPath))
                return false;

            string ignoreFilePath = Path.Combine(folderPath, fileName);
            return File.Exists(ignoreFilePath);
        }

        /// <summary>
        /// Parses an ignore file and extracts patterns and directives
        /// </summary>
        /// <param name="filePath">Path to the ignore file</param>
        /// <param name="shouldTransfer">Output parameter indicating if transfer is enabled</param>
        /// <returns>List of patterns found in the file</returns>
        public static List<string> ParseIgnoreFile(string filePath, out bool shouldTransfer)
        {
            List<string> patterns = new List<string>();
            shouldTransfer = true;

            try
            {
                string[] lines = File.ReadAllLines(filePath);

                foreach (string line in lines)
                {
                    string trimmedLine = line.Trim();

                    // Skip empty lines and comments
                    if (string.IsNullOrWhiteSpace(trimmedLine) || trimmedLine.StartsWith("#"))
                        continue;

                    // Check for special directive
                    if (trimmedLine.Equals("@disable-transfer", StringComparison.OrdinalIgnoreCase))
                    {
                        shouldTransfer = false;
                        continue;
                    }

                    // Add pattern to the list
                    patterns.Add(trimmedLine);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error parsing ignore file: {ex.Message}");
            }

            return patterns;
        }

        /// <summary>
        /// Ensures a directory exists, creating it if necessary
        /// </summary>
        /// <param name="path">The directory path</param>
        public static void EnsureDirectoryExists(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }
    }
}