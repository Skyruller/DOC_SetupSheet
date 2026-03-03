using System.Runtime.InteropServices;

namespace SetupSheet
{
    internal static class ComSafe
    {
        public static void Release(object o)
        {
            if (o == null) return;
            try { Marshal.FinalReleaseComObject(o); } catch { }
        }

        public static void Cleanup()
        {
            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();
            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();
        }
    }
}
