using System;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows.Forms;

internal static class Program
{
    [STAThread]
    static int Main(string[] args)
    {
        var assembly = Assembly.GetExecutingAssembly();
        AppDomain.CurrentDomain.AssemblyResolve += (sender, e) =>
        {
            var name = new AssemblyName(e.Name).Name;
            using (var resource = assembly.GetManifestResourceStream("bundle." + name + ".dll"))
            {
                if (resource == null) return null;
                using (var copy = new MemoryStream())
                {
                    resource.CopyTo(copy);
                    return Assembly.Load(copy.ToArray());
                }
            }
        };
        return Run(args);
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    static int Run(string[] args)
    {
        if (Array.IndexOf(args, "--self-test") >= 0) return Checks.Run();
        using (var instance = new Mutex(true, @"Local\OutlookEnvelope.1", out bool first))
        using (var show = new EventWaitHandle(false, EventResetMode.AutoReset, @"Local\OutlookEnvelope.Settings.1"))
        {
            if (!first) { show.Set(); return 0; }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            using (var window = new EnvelopeWindow(show)) Application.Run(window);
        }
        return 0;
    }
}
