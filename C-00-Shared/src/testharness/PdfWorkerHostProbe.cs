using System;
using System.Reflection;

// No COM activation: verify the gate used before opening an Office source.
internal static class PdfWorkerHostProbe
{
    private static int Main(string[] args)
    {
        try
        {
            var worker = Assembly.LoadFrom(System.IO.Path.GetFullPath(args[0]));
            var gate = worker.GetType("PdfWorker", true).GetMethod("CanOpenOfficeSource", BindingFlags.NonPublic | BindingFlags.Static);
            object[][] cases = {
                new object[] { true, (uint)0, false, 0, true, "WPS Writer without an application HWND and no documents" },
                new object[] { true, (uint)0, false, 1, false, "Unidentified Writer containing a user document" },
                new object[] { true, (uint)123, false, 0, false, "Writer identified as an existing process" },
                new object[] { false, (uint)0, false, 0, false, "Unidentified spreadsheet host" },
                new object[] { false, (uint)123, true, 0, true, "New empty spreadsheet process" },
                new object[] { true, (uint)123, true, 2, false, "New process that already contains documents" }
            };
            foreach (object[] test in cases)
            {
                bool actual = (bool)gate.Invoke(null, new object[] { test[0], test[1], test[2], test[3] });
                if (actual != (bool)test[4]) throw new Exception("Host gate mismatch: " + test[5]);
            }
            Console.WriteLine("PASS: " + cases.Length + " Office host ownership scenarios; no Office activation.");
            return 0;
        }
        catch (Exception error) { Console.WriteLine(error); return 1; }
    }
}
