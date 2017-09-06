using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;
using ResMerger;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Xml.Linq;

namespace ResMergerCustomTool
{
    [Guid("86549FFF-7343-4AAB-847F-A0320B6718F2")]
    [CustomTool("ResMergerGenerator", "Merges the used resource dictionaries in one file")]
    public class ResMergerGenerator : IVsSingleFileGenerator
    {
        internal static Guid CSharpCategoryGuid = new Guid("FAE04EC1-301F-11D3-BF4B-00C04F79EFBC");
        private const string VisualStudioVersion = "14.0";

        public int DefaultExtension(out string pbstrDefaultExtension)
        {
            pbstrDefaultExtension = ".xaml";
            return pbstrDefaultExtension.Length;
        }

        public int Generate(string wszInputFilePath, string bstrInputFileContents, string wszDefaultNamespace, IntPtr[] rgbOutputFileContents, out uint pcbOutput, IVsGeneratorProgress pGenerateProgress)
        {
            //GetProjectPath
            string projectPath = GetProjectPath(wszInputFilePath);
            //GetProjectName
            string projectName = GetProjectName(projectPath);
            //GetRelativeSourcePathOfFile
            string relativeSourceFilePath = GetRelativeSourceFilePath(projectPath, wszInputFilePath);

            XDocument mergedDocument = ResourceMerger.CreateOutputDocument(projectPath, projectName, relativeSourceFilePath);
            using (MemoryStream ms = new MemoryStream())
            {
                mergedDocument.Save(ms);
                byte[] bytes = ms.ToArray();
                int length = bytes.Length;
                rgbOutputFileContents[0] = Marshal.AllocCoTaskMem(length);
                Marshal.Copy(bytes, 0, rgbOutputFileContents[0], length);
                pcbOutput = (uint)length;
                return VSConstants.S_OK;
            }
        }

        private string GetRelativeSourceFilePath(string projectPath, string wszInputFilePath)
        {
            return wszInputFilePath.Replace(@"\", @"/") .Replace(projectPath, string.Empty);
        }

        private string GetProjectName(string projectPath)
        {
            DirectoryInfo di = new DirectoryInfo(projectPath);
            FileInfo[] fi = di.GetFiles("*.csproj");
            return fi[0].Name;
        }

        private string GetProjectPath(string wszInputFilePath)
        {
            string directory = Path.GetDirectoryName(wszInputFilePath);
            DirectoryInfo di = new DirectoryInfo(directory);

            FileInfo[] fi = di.GetFiles("*.csproj");
            if (fi.Length > 0)
            {
                return fi[0].DirectoryName;
            }

            di = di.Parent;
            fi = di.GetFiles("*.csproj");
            if (fi.Length > 0)
            {
                return fi[0].DirectoryName;
            }

            return string.Empty;
        }


        internal static GuidAttribute GetGuidAttribute(Type t)
        {
            return (GuidAttribute)GetAttribute(t, typeof(GuidAttribute));
        }

        internal static CustomToolAttribute GetCustomToolAttribute(Type t)
        {
            return (CustomToolAttribute)GetAttribute(t, typeof(CustomToolAttribute));
        }

        internal static Attribute GetAttribute(Type t, Type attributeType)
        {
            object[] attributes = t.GetCustomAttributes(attributeType, /* inherit */ true);
            if (attributes.Length == 0)
                throw new Exception(
                    String.Format("Class '{0}' does not provide a '{1}' attribute.",
                        t.FullName, attributeType.FullName));
            return (Attribute)attributes[0];
        }

        internal static string GetKeyName(Guid categoryGuid, string toolName)
        {
            return
                String.Format("SOFTWARE\\Microsoft\\VisualStudio\\" + VisualStudioVersion +
                              "\\Generators\\{{{0}}}\\{1}\\", categoryGuid, toolName);
        }

        [ComRegisterFunction]
        public static void RegisterClass(Type t)
        {
            GuidAttribute guidAttribute = GetGuidAttribute(t);
            CustomToolAttribute customToolAttribute = GetCustomToolAttribute(t);
            using (RegistryKey key = Registry.LocalMachine.CreateSubKey(GetKeyName(CSharpCategoryGuid, customToolAttribute.Name)))
            {
                key.SetValue("", customToolAttribute.Description);
                key.SetValue("CLSID", "{" + guidAttribute.Value + "}");
                key.SetValue("GeneratesDesignTimeSource", 1);
            }
        }

        [ComUnregisterFunction]
        public static void UnregisterClass(Type t)
        {
            CustomToolAttribute customToolAttribute = GetCustomToolAttribute(t);
            Registry.LocalMachine.DeleteSubKey(GetKeyName(CSharpCategoryGuid, customToolAttribute.Name), false);
        }
    }
}
