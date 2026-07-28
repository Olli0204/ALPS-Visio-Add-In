using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace ALPS_Visio_AddIn_rewrite.VisioInfrastructure
{
    /// <summary>
    /// Locates the newest installed SID and SBD stencil files without depending
    /// directly on the VSTO application singleton.
    /// </summary>
    internal sealed class StencilFileLocator
    {
        private const string SidPrefix = "Abstract PASS SID Visio Shapes";
        private const string SbdPrefix = "Abstract PASS SBD Visio Shapes";
        private const string VersionPattern = "\\s(v)(\\d*)\\.(\\d*)\\.(\\d*)(\\S)*";
        private const string Ending = ".vssm";

        private readonly Func<string> shapesPathProvider;
        private string sidName;
        private string sbdName;

        public StencilFileLocator(Func<string> shapesPathProvider)
        {
            this.shapesPathProvider = shapesPathProvider
                ?? throw new ArgumentNullException(nameof(shapesPathProvider));
        }

        public string GetSidName()
        {
            if (sidName == null)
                sidName = FindNewestStencil(SidPrefix);

            return sidName;
        }

        public string GetSbdName()
        {
            if (sbdName == null)
                sbdName = FindNewestStencil(SbdPrefix);

            return sbdName;
        }

        private string FindNewestStencil(string prefix)
        {
            DirectoryInfo shapesDirectory = new DirectoryInfo(shapesPathProvider());
            string namePattern = prefix + VersionPattern + Ending;
            IList<FileInfo> possibleFiles = FindMatchingFiles(shapesDirectory, namePattern);

            if (possibleFiles.Count == 0)
                return GetMissingStencilName(prefix);

            FileInfo newestFile = possibleFiles[0];
            foreach (FileInfo candidate in possibleFiles)
            {
                string newestName = GetNewestName(newestFile.Name, candidate.Name);
                if (!newestFile.Name.Equals(newestName))
                    newestFile = candidate;
            }

            return newestFile.Name.StartsWith("~$$")
                ? GetMissingStencilName(prefix)
                : newestFile.Name;
        }

        private static string GetNewestName(string firstName, string secondName)
        {
            if (firstName.Equals(secondName))
                return firstName;

            string firstVersion = Regex.Match(firstName, VersionPattern).ToString();
            string secondVersion = Regex.Match(secondName, VersionPattern).ToString();
            firstVersion = firstVersion.Replace(Ending, "");
            secondVersion = secondVersion.Replace(Ending, "");

            const string endPattern = "[^\\.\\d]$";
            string firstEnding = Regex.Match(firstVersion, endPattern).ToString();
            string secondEnding = Regex.Match(secondVersion, endPattern).ToString();
            firstVersion = Regex.Replace(firstVersion, endPattern, "");
            secondVersion = Regex.Replace(secondVersion, endPattern, "");

            Regex numbers = new Regex(@"\d+");
            MatchCollection firstNumbers = numbers.Matches(firstVersion);
            MatchCollection secondNumbers = numbers.Matches(secondVersion);

            for (int index = 0; index < firstNumbers.Count; index++)
            {
                if (secondNumbers.Count <= index)
                    return firstName;

                int first = int.Parse(firstNumbers[index].Value);
                int second = int.Parse(secondNumbers[index].Value);

                if (first > second)
                    return firstName;

                if (first < second)
                    return secondName;
            }

            if (secondNumbers.Count > firstNumbers.Count)
                return secondName;

            if (firstEnding.Length > secondEnding.Length)
                return firstName;

            if (firstEnding.Length < secondEnding.Length)
                return secondName;

            return firstName;
        }

        private static IList<FileInfo> FindMatchingFiles(DirectoryInfo root,
            string namePattern)
        {
            FileInfo[] files;
            try
            {
                files = root.GetFiles("*.*");
            }
            catch (UnauthorizedAccessException exception)
            {
                Console.WriteLine(exception.Message);
                return new List<FileInfo>();
            }
            catch (DirectoryNotFoundException exception)
            {
                Console.WriteLine(exception.Message);
                return new List<FileInfo>();
            }

            Regex nameRegex = new Regex(namePattern);
            List<FileInfo> matchingFiles = new List<FileInfo>();
            foreach (FileInfo file in files)
            {
                if (nameRegex.IsMatch(file.Name))
                    matchingFiles.Add(file);
            }

            foreach (DirectoryInfo directory in root.GetDirectories())
            {
                matchingFiles.AddRange(FindMatchingFiles(directory, namePattern));
            }

            return matchingFiles;
        }

        private static string GetMissingStencilName(string prefix)
        {
            return prefix + " v.x.x.x.x" + Ending;
        }
    }
}
