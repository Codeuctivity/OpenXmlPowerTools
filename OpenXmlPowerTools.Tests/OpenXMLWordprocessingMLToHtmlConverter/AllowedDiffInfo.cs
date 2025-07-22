namespace Codeuctivity.Tests.OpenXMLWordProcessingMLToHtmlConverter
{
    internal class AllowedDiffInfo
    {
        public bool DiffFileExists;
        public string NewDiffImageFileName;
        public string[] ExistingDiffImageFilename;

        public AllowedDiffInfo(bool diffFileExists, string newDiffImageFileName, string[] matchingFiles)
        {
            DiffFileExists = diffFileExists;
            NewDiffImageFileName = newDiffImageFileName;
            ExistingDiffImageFilename = matchingFiles;
        }
    }
}