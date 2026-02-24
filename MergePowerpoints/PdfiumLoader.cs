using System;
using System.IO;
using System.Runtime.InteropServices;

namespace MergePowerpoints;

internal static class PdfiumLoader
{
    private static bool _initialized;

    static PdfiumLoader()
    {
        try
        {
            // Register resolver for the assembly that performs DllImports to "pdfium"
            // PDFiumSharp v1 uses the namespace PDFiumSharp but the DllImport target name is "pdfium"
            NativeLibrary.SetDllImportResolver(typeof(PDFiumSharp.PdfDocument).Assembly, ResolvePdfium);
            _initialized = true;
        }
        catch
        {
            _initialized = false;
        }
    }

    public static void EnsureLoaded()
    {
        // Intentionally no-op: triggering the static ctor is enough to register the resolver
        if (!_initialized)
        {
            // Best-effort manual load in case resolver registration failed
            TryLoadManually();
        }
    }

    private static IntPtr ResolvePdfium(string libraryName, System.Reflection.Assembly assembly, DllImportSearchPath? searchPath)
    {
        if (!string.Equals(libraryName, "pdfium", StringComparison.OrdinalIgnoreCase))
            return IntPtr.Zero;

        // Probe likely locations and file names
        var handle = TryLoadFromKnownLocations();
        return handle;
    }

    private static IntPtr TryLoadFromKnownLocations()
    {
        string baseDir = AppContext.BaseDirectory;
        bool is64 = Environment.Is64BitProcess;

        // 1) Architecture-suffixed files placed at output root (observed in current bin): pdfium_x64.dll / pdfium_x86.dll
        string archFile = Path.Combine(baseDir, is64 ? "pdfium_x64.dll" : "pdfium_x86.dll");
        if (File.Exists(archFile) && NativeLibrary.TryLoad(archFile, out var handle))
            return handle;

        // 2) Standard name at output root
        string stdRoot = Path.Combine(baseDir, "pdfium.dll");
        if (File.Exists(stdRoot) && NativeLibrary.TryLoad(stdRoot, out handle))
            return handle;

        // 3) Runtimes subfolder used by native packages
        string rid = is64 ? "win-x64" : "win-x86";
        string runtimePath = Path.Combine(baseDir, "runtimes", rid, "native", "pdfium.dll");
        if (File.Exists(runtimePath) && NativeLibrary.TryLoad(runtimePath, out handle))
            return handle;

        // 4) As a last resort, let default resolver try (returns IntPtr.Zero to continue default resolution)
        return IntPtr.Zero;
    }

    private static void TryLoadManually()
    {
        // Attempt an eager load to surface any errors early
        var handle = TryLoadFromKnownLocations();
        if (handle != IntPtr.Zero)
        {
            // Loaded successfully
            return;
        }

        // If we couldn't load, do nothing: consuming code will still fail with a clear DllNotFoundException
    }
}
