using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using ResMerger;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ResMergerCustomTool
{
    public class ResMergerGenerator : IVsSingleFileGenerator
    {
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
            throw new NotImplementedException();
        }

        private string GetProjectName(string projectPath)
        {
            throw new NotImplementedException();
        }

        private string GetProjectPath(string wszInputFilePath)
        {
            throw new NotImplementedException();
        }
    }
}
