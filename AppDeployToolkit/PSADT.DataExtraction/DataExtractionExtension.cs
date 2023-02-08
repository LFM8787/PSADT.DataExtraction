// Date Modified: 26/01/2021
// Version Number: 3.8.4

using System;
using System.Text;
using System.Drawing;
using System.Runtime.InteropServices;

namespace PSADT
{
	public class DataExtraction
	{
		[System.Flags]
		enum LoadLibraryFlags : uint
		{
			DONT_RESOLVE_DLL_REFERENCES         = 0x00000001,
			LOAD_IGNORE_CODE_AUTHZ_LEVEL        = 0x00000010,
			LOAD_LIBRARY_AS_DATAFILE            = 0x00000002,
			LOAD_LIBRARY_AS_DATAFILE_EXCLUSIVE  = 0x00000040,
			LOAD_LIBRARY_AS_IMAGE_RESOURCE      = 0x00000020,
			LOAD_LIBRARY_SEARCH_APPLICATION_DIR = 0x00000200,
			LOAD_LIBRARY_SEARCH_DEFAULT_DIRS    = 0x00001000,
			LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR    = 0x00000100,
			LOAD_LIBRARY_SEARCH_SYSTEM32        = 0x00000800,
			LOAD_LIBRARY_SEARCH_USER_DIRS       = 0x00000400,
			LOAD_WITH_ALTERED_SEARCH_PATH       = 0x00000008
		}

		[DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = false)]
		static extern IntPtr LoadLibraryEx(string lpFileName, IntPtr hFile, LoadLibraryFlags dwFlags);

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)]
		static extern int LoadString(IntPtr hInstance, uint uID, StringBuilder lpBuffer, int nBufferMax);

		// Get specific string from specific resource file
		public static string ExtractStringFromFile(string file, int number)
		{
			StringBuilder sb = new StringBuilder(2048);
			uint u = Convert.ToUInt32(number);
			try
			{
				IntPtr lib = LoadLibraryEx(file, IntPtr.Zero, (LoadLibraryFlags.LOAD_LIBRARY_AS_IMAGE_RESOURCE | LoadLibraryFlags.LOAD_LIBRARY_AS_DATAFILE));
				LoadString(lib, u, sb, sb.Capacity);

				return sb.ToString();
			}
			catch
			{
				return null;
			}
		}

		[DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
		private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);

		// Get specific icon index inside a resource file
		public static Icon ExtractIcon(string file, int number, bool largeIcon)
		{
			IntPtr large;
			IntPtr small;
			ExtractIconEx(file, number, out large, out small, 1);
			try
			{
				return Icon.FromHandle(largeIcon ? large : small);
			}
			catch
			{
				return null;
			}
		}
	}
}