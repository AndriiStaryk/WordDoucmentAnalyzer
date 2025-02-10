using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace DocumentAnalyzer;
internal class FileManager
{
    private static readonly Guid FOLDERID_Downloads = new Guid("374DE290-123F-4565-9164-39C4925E467B");

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern int SHGetKnownFolderPath(Guid rfid, uint dwFlags, IntPtr hToken, out IntPtr ppszPath);

    public static string GetDownloadsFolderPath()
    {
        IntPtr ppszPath = IntPtr.Zero;
        try
        {
            int hr = SHGetKnownFolderPath(FOLDERID_Downloads, 0, IntPtr.Zero, out ppszPath);
            if (hr != 0)
            {
                throw new System.ComponentModel.Win32Exception(hr);
            }

            string path = Marshal.PtrToStringUni(ppszPath) ?? string.Empty;

            if (string.IsNullOrEmpty(path))
            {
                throw new System.ComponentModel.Win32Exception(hr);
            }

            return path;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error retrieving Downloads folder: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return string.Empty;
        }
        finally
        {
            if (ppszPath != IntPtr.Zero)
            {
                Marshal.FreeCoTaskMem(ppszPath); 
            }
        }
    }
}