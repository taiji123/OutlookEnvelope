using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Identity.Client;

internal static class Checks
{
    sealed class Replay : HttpMessageHandler
    {
        public readonly Queue<string> Pages = new Queue<string>();
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancel)
        {
            if (request.Headers.Authorization.Parameter != "offline-test") throw new Exception("Missing authorization");
            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(Pages.Dequeue()) });
        }
    }
    static void Expect(bool ok, string message) { if (!ok) throw new Exception(message); }
    public static int Run()
    {
        try { Verify().GetAwaiter().GetResult(); return 0; }
        catch (Exception e) { Console.Error.WriteLine(e); return 1; }
    }
    static async Task Verify()
    {
        var session = new SignIn(Guid.NewGuid(), Guid.NewGuid());
        try { await session.Get(false, false, CancellationToken.None); throw new Exception("Expected sign-in request."); }
        catch (MsalUiRequiredException) { }
        using (var icon = EnvelopeWindow.DrawEnvelope()) Expect(icon.Width == 32, "Envelope artwork");
        var replay = new Replay();
        using (var http = new HttpClient(replay))
        {
            var box = new Mailbox(http);
            var old = DateTimeOffset.UtcNow.AddDays(-1);
            var now = DateTimeOffset.UtcNow.AddSeconds(1);
            Func<string, bool, DateTimeOffset, string> item = (id, read, received) =>
                "{\"id\":\"" + id + "\",\"isRead\":" + (read ? "true" : "false") + ",\"receivedDateTime\":\"" + received.ToString("o") + "\"}";
            Func<string, bool, string> page = (values, more) =>
                "{\"value\":[" + values + "],\"@odata." + (more ? "nextLink" : "deltaLink") + "\":\"https://graph.microsoft.com/v1.0/test\"}";
            Func<Task> poll = () => box.Poll("offline-test", CancellationToken.None);
            replay.Pages.Enqueue(page(item("old", false, old), false));
            await poll(); Expect(!box.State.Visible, "Existing unread mail lit the envelope.");
            replay.Pages.Enqueue(page(item("a", false, now), true));
            replay.Pages.Enqueue(page(item("b", false, now), false));
            await poll(); Expect(box.State.Visible, "Paged arrivals missing.");
            replay.Pages.Enqueue(page("{\"id\":\"a\",\"isRead\":true}", false));
            await poll(); Expect(!box.State.Visible, "Reading one of two should clear.");
            replay.Pages.Enqueue(page(item("a", false, now) + "," + item("b", false, now), false));
            await poll(); Expect(!box.State.Visible, "Old unread mail relit the envelope.");
            replay.Pages.Enqueue(page(item("c", false, now.AddMinutes(1)), false));
            await poll(); Expect(box.State.Visible, "Next arrival did not relight.");
            replay.Pages.Enqueue(page(item("b", true, now), false));
            await poll(); Expect(box.State.Visible, "Reading an earlier acknowledged message cleared a newer delivery.");
            replay.Pages.Enqueue(page("{\"id\":\"c\",\"@removed\":{\"reason\":\"deleted\"}}", false));
            await poll(); Expect(!box.State.Visible, "Removed delivery left envelope visible.");
        }
    }
}
