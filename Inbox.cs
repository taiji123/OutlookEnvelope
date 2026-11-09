using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Automation;

internal static class Inbox
{
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr hwnd);
    [DllImport("user32.dll")] static extern bool IsIconic(IntPtr hwnd);
    [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr hwnd, int cmd);
    static readonly Regex inbox = new Regex(@"^Inbox(?:$|[\s,(]+(?:\d+\b|folder\b))", RegexOptions.IgnoreCase);

    public static async Task Open(string account)
    {
        // Activate the packaged new Outlook explicitly even if classic Outlook is the default.
        Process.Start("explorer.exe", @"shell:AppsFolder\Microsoft.OutlookForWindows_8wekyb3d8bbwe!Microsoft.OutlookforWindows");
        await Task.Run(async () =>
        {
            bool mailSelected = false;
            for (int i = 0; i < 28; i++)
            {
                await Task.Delay(500);
                foreach (var p in Process.GetProcessesByName("olk"))
                using (p)
                {
                    var handle = p.MainWindowHandle;
                    if (handle == IntPtr.Zero) continue;
                    try
                    {
                        if (IsIconic(handle)) ShowWindow(handle, 9);
                        var root = AutomationElement.FromHandle(handle);
                        var folders = root.FindAll(TreeScope.Descendants, OfType(ControlType.TreeItem))
                            .Cast<AutomationElement>().Where(e => inbox.IsMatch(e.Current.Name ?? "")).ToList();
                        var selected = folders.FirstOrDefault(e => UnderAccount(e, account));
                        if (selected == null)
                        {
                            var direct = folders.Where(e => !UnderAccount(e, "Favorites")).ToList();
                            if (direct.Count == 1) selected = direct[0];
                            else if (folders.Count == 1) selected = folders[0];
                        }
                        if (selected != null && Activate(selected)) { SetForegroundWindow(handle); return; }
                        if (!mailSelected)
                        {
                            var mail = root.FindFirst(TreeScope.Descendants, new AndCondition(
                                new PropertyCondition(AutomationElement.NameProperty, "Mail"),
                                new OrCondition(OfType(ControlType.Button), OfType(ControlType.TabItem), OfType(ControlType.ListItem))));
                            if (mail != null) mailSelected = Activate(mail);
                        }
                        if (!string.IsNullOrEmpty(account))
                        {
                            var items = root.FindAll(TreeScope.Descendants, OfType(ControlType.TreeItem));
                            foreach (AutomationElement node in items)
                            {
                                if ((node.Current.Name ?? "").IndexOf(account, StringComparison.OrdinalIgnoreCase) < 0) continue;
                                if (node.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out var pattern))
                                {
                                    var expandable = (ExpandCollapsePattern)pattern;
                                    if (expandable.Current.ExpandCollapseState == ExpandCollapseState.Collapsed) expandable.Expand();
                                }
                            }
                        }
                    }
                    catch (ElementNotAvailableException) { }
                    catch (InvalidOperationException) { }
                }
            }
            throw new InvalidOperationException("Outlook opened, but the work Inbox could not be selected. Expand your work account in the Mail folder pane and try again. The folder must be named Inbox.");
        });
    }
    static Condition OfType(ControlType type) => new PropertyCondition(AutomationElement.ControlTypeProperty, type);
    static bool UnderAccount(AutomationElement item, string account)
    {
        if (string.IsNullOrEmpty(account)) return false;
        for (var node = TreeWalker.ControlViewWalker.GetParent(item); node != null; node = TreeWalker.ControlViewWalker.GetParent(node))
        {
            if (node.Current.ControlType == ControlType.Window || node.Current.ControlType == ControlType.Document) break;
            if ((node.Current.Name ?? "").IndexOf(account, StringComparison.OrdinalIgnoreCase) >= 0) return true;
        }
        return false;
    }
    static bool Activate(AutomationElement item)
    {
        if (item.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var select))
        { ((SelectionItemPattern)select).Select(); return true; }
        if (item.TryGetCurrentPattern(InvokePattern.Pattern, out var invoke))
        { ((InvokePattern)invoke).Invoke(); return true; }
        return false;
    }
}
