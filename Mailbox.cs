using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

internal sealed class Change
{
    public string Id;
    public bool Removed, Read;
    public DateTimeOffset Received, Modified;
}

internal sealed class EnvelopeState
{
    readonly HashSet<string> seen = new HashSet<string>(StringComparer.Ordinal);
    readonly HashSet<string> pending = new HashSet<string>(StringComparer.Ordinal);
    readonly DateTimeOffset started;
    public EnvelopeState(DateTimeOffset started) { this.started = started; }
    public bool Visible => pending.Count > 0;

    public void Apply(IEnumerable<Change> changes)
    {
        // Delta pages are not chronological. Apply deliveries and read transitions by time.
        foreach (var change in changes.OrderBy(c => c.Removed ? DateTimeOffset.MaxValue :
            c.Read ? c.Modified : c.Received).ThenBy(c => c.Read))
        {
            if (change.Removed) { pending.Remove(change.Id); seen.Remove(change.Id); continue; }
            bool first = seen.Add(change.Id);
            if (change.Read && (pending.Contains(change.Id) || (first && change.Received >= started))) pending.Clear();
            if (first && !change.Read && change.Received >= started) pending.Add(change.Id);
        }
    }
}

internal sealed class ResyncRequired : Exception { }
internal sealed class MailError : Exception
{
    public readonly int RetrySeconds;
    public readonly HttpStatusCode Code;
    public MailError(HttpStatusCode code, int retry) : base("Microsoft 365 returned " + (int)code + " (" + code + ").")
    { RetrySeconds = retry; Code = code; }
}

internal sealed class Mailbox
{
    static readonly HttpClient sharedClient = new HttpClient(new HttpClientHandler { AllowAutoRedirect = false });
    readonly HttpClient http;
    public Mailbox(HttpClient client = null) { http = client ?? sharedClient; }
    const string FirstPage = "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages/delta?$select=id,isRead,receivedDateTime,lastModifiedDateTime";
    readonly JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = 16 * 1024 * 1024 };
    string cursor;
    public EnvelopeState State { get; private set; } = new EnvelopeState(DateTimeOffset.UtcNow);
    public async Task Poll(string token, CancellationToken cancel)
    {
        string page = cursor ?? FirstPage;
        var batch = new Dictionary<string, Change>(StringComparer.Ordinal);
        while (true)
        {
            var uri = new Uri(page);
            if (uri.Scheme != "https" || uri.Host != "graph.microsoft.com" || !uri.IsDefaultPort)
                throw new InvalidOperationException("Unexpected Microsoft Graph page address.");
            using (var request = new HttpRequestMessage(HttpMethod.Get, uri))
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                request.Headers.Add("Prefer", "odata.maxpagesize=500");
                using (var response = await http.SendAsync(request, cancel))
                {
                    if (response.StatusCode == HttpStatusCode.Gone) throw new ResyncRequired();
                    if (!response.IsSuccessStatusCode)
                    {
                        var wait = response.Headers.RetryAfter;
                        var delay = wait?.Delta ?? (wait?.Date - DateTimeOffset.UtcNow);
                        throw new MailError(response.StatusCode, (int)Math.Ceiling(delay?.TotalSeconds ?? 0));
                    }
                    var result = json.Deserialize<Dictionary<string, object>>(await response.Content.ReadAsStringAsync());
                    foreach (Dictionary<string, object> item in (System.Collections.IEnumerable)result["value"])
                    {
                        if (!item.ContainsKey("id")) continue;
                        var change = new Change { Id = (string)item["id"], Removed = item.ContainsKey("@removed") };
                        if (!change.Removed)
                        {
                            if (!item.ContainsKey("isRead")) continue;
                            change.Read = (bool)item["isRead"];
                            if (item.ContainsKey("receivedDateTime"))
                                change.Received = DateTimeOffset.Parse((string)item["receivedDateTime"], CultureInfo.InvariantCulture);
                            else if (batch.TryGetValue(change.Id, out var prior))
                                change.Received = prior.Received;
                            change.Modified = item.ContainsKey("lastModifiedDateTime") ?
                                DateTimeOffset.Parse((string)item["lastModifiedDateTime"], CultureInfo.InvariantCulture) : DateTimeOffset.UtcNow;
                        }
                        batch[change.Id] = change;
                    }
                    if (result.TryGetValue("@odata.nextLink", out var next)) { page = (string)next; continue; }
                    var nextCursor = (string)result["@odata.deltaLink"];
                    cancel.ThrowIfCancellationRequested();
                    State.Apply(batch.Values);
                    cursor = nextCursor;
                    return;
                }
            }
        }
    }
}
