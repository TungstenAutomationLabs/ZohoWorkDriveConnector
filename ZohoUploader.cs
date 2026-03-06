
// =============================================================================
// ZohoUploader.cs
// TotalAgility (KTA Cloud) → Zoho WorkDrive Integration
//
// USER GUIDE:
// ───────────
// In TotalAgility Script Task, select:
//   Class  : ZohoUploader
//   Method : UploadFolderDocuments
//
// INPUT parameters:
//   sessionId           — KTA session ID (SPP_SYSTEM_SESSION_ID)
//   ktaUrl              — https://<tenant>.totalagility.com/TotalAgility/Agility.Server.Web
//   ktaFolderId         — KTA source folder ID
//   zohoAccountsUrl     — https://accounts.zoho.com
//   zohoUploadUrl       — https://workdrive.zoho.com/api/v1/upload
//   zohoMetadataUrl     — https://www.zohoapis.com/workdrive/api/v1/custommetadata
//   zohoClientId        — Zoho API Console → Client ID
//   zohoClientSecret    — Zoho API Console → Client Secret
//   zohoFolderId        — Zoho WorkDrive target folder ID
//   zohoScope           — WorkDrive.files.CREATE,WorkDrive.DataTemplates.CREATE
//   zohoDataTemplateId  — Zoho data template ID (from WorkDrive Data Templates)
//   zohoFieldMapping    — Semicolon-separated list of  KtaFieldName:ZohoCustomFieldId
//                         e.g. "InvoiceNumber:abc-001;InvoiceDate:abc-002;Total:abc-003"
//                         KtaFieldName must exactly match the field name in KTA.
//                         If a KTA field name is not mapped it is silently skipped.
//
// OUTPUT parameters:
//   Result              — "Upload complete. Success: X, Failed: Y"
//   DetailedLog         — Per-document upload status and field values
//
// NuGet dependency: Newtonsoft.Json
// =============================================================================

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace TotalAgilityZohoIntegration
{
    public class ZohoUploader
    {
        private static readonly HttpClient _httpClient = new HttpClient();

        private string _cachedZohoToken = null;
        private DateTime _zohoTokenExpiry = DateTime.MinValue;
        private string _cachedClientId = null;
        private readonly object _tokenLock = new object();

        // =====================================================================
        // PUBLIC — Only method visible in TotalAgility
        // =====================================================================
        public void UploadFolderDocuments(
string sessionId,
string ktaUrl,
string ktaFolderId,
string zohoAccountsUrl,
string zohoUploadUrl,
string zohoMetadataUrl,
string zohoClientId,
string zohoClientSecret,
string zohoFolderId,
string zohoScope,
string zohoDataTemplateId,
string zohoFieldMapping,
out string Result,
out string DetailedLog)
        {
            Result = "";
            DetailedLog = "";
            var log = new StringBuilder();

            try
            {
                var kta = new KtaConfig
                {
                    SessionId = sessionId,
                    Url = ktaUrl,
                    FolderId = ktaFolderId
                };

                var zoho = new ZohoConfig
                {
                    AccountsUrl = zohoAccountsUrl,
                    UploadUrl = zohoUploadUrl,
                    MetadataUrl = zohoMetadataUrl,
                    ClientId = zohoClientId,
                    ClientSecret = zohoClientSecret,
                    FolderId = zohoFolderId,
                    Scope = zohoScope,
                    DataTemplateId = zohoDataTemplateId,
                    FieldMapping = ParseFieldMapping(zohoFieldMapping)
                };

                kta.Validate();
                zoho.Validate();

                var results = RunUploadAsync(kta, zoho, log)
                .GetAwaiter().GetResult();

                int success = 0, failed = 0;

                foreach (var r in results)
                {
                    if (r == null) continue;

                    if (r.Success)
                    {
                        success++;
                        log.AppendLine($"[OK]   Document         : {r.DocumentName}");
                        log.AppendLine($"       PDF Size        : {r.PdfSizeBytes:N0} bytes");
                        log.AppendLine($"       PDF Resource ID : {r.ZohoPdfFileId}");
                        log.AppendLine($"       Metadata Update : {(r.MetadataUpdated ? "Success" : "Skipped/Not configured")}");
                    }
                    else
                    {
                        failed++;
                        log.AppendLine($"[FAIL] Document        : {r.DocumentName}");
                        log.AppendLine($"       Error           : {r.ErrorMessage}");
                    }

                    if (r.Fields != null && r.Fields.Count > 0)
                    {
                        log.AppendLine("       Fields:");
                        foreach (var kv in r.Fields)
                            log.AppendLine($"         {kv.Key} = {kv.Value}");
                    }

                    log.AppendLine();
                }

                Result = $"Upload complete. Success: {success}, Failed: {failed}";
                DetailedLog = log.ToString();
            }
            catch (Exception ex)
            {
                Result = "Error: " + ex.Message;
                DetailedLog = log.ToString() + Environment.NewLine + ex.ToString();
            }
        }

        // =====================================================================
        // ALL METHODS AND CLASSES BELOW ARE PRIVATE
        // =====================================================================

        private async Task<List<UploadItem>> RunUploadAsync(
KtaConfig kta,
ZohoConfig zoho,
StringBuilder log)
        {
            var results = new List<UploadItem>();
            string ktaUrl = kta.Url.TrimEnd('/');

            List<FolderDocument> docs =
            GetFolderDocuments(kta.SessionId, kta.FolderId, ktaUrl);

            if (docs == null || docs.Count == 0)
                throw new Exception("No documents found in KTA folder.");

            log.AppendLine($"Found {docs.Count} document(s) in folder.");
            log.AppendLine();

            string zohoToken = await GetZohoToken(zoho);

            foreach (var doc in docs)
            {
                var item = new UploadItem
                {
                    DocumentName = doc.FileName ?? doc.Id
                };

                try
                {
                    DocumentData data = GetDocumentData(doc.Id, ktaUrl, kta.SessionId);

                    if (string.IsNullOrEmpty(data.FileName))
                        data.FileName = doc.FileName ?? $"Document_{doc.Id}";

                    item.DocumentName = data.FileName;
                    item.Fields = data.Fields ?? new Dictionary<string, string>();

                    byte[] pdfBytes = GetDocumentAsPdf(doc.Id, ktaUrl, kta.SessionId);

                    if (pdfBytes == null || pdfBytes.Length == 0)
                        throw new Exception("KTA returned empty PDF content.");

                    item.PdfSizeBytes = pdfBytes.Length;

                    string pdfName = EnsurePdfExtension(data.FileName);

                    // ── 1. Upload PDF ──────────────────────────────────────────
                    item.ZohoPdfFileId = await UploadToZoho(
pdfBytes, pdfName, "application/pdf", zohoToken, zoho);

                    // ── 2. Update Zoho custom metadata on the uploaded PDF ─────
                    if (!string.IsNullOrWhiteSpace(zoho.DataTemplateId)
&& !string.IsNullOrWhiteSpace(item.ZohoPdfFileId)
&& zoho.FieldMapping != null
&& zoho.FieldMapping.Count > 0)
                    {
                        await UpdateCustomMetadata(
                        item.ZohoPdfFileId,
                        zoho.DataTemplateId,
                        data.Fields,
                        zoho.FieldMapping,
                        zohoToken,
                        zoho.MetadataUrl);

                        item.MetadataUpdated = true;
                    }

                    item.Success = true;
                }
                catch (Exception ex)
                {
                    item.Success = false;
                    item.ErrorMessage = ex.Message ?? "Unknown error.";
                }

                results.Add(item);
            }

            return results;
        }

        // =====================================================================
        // Custom metadata update — /workdrive/api/v1/custommetadata (JSON:API)
        // =====================================================================

        /// <summary>
                    /// Updates Zoho WorkDrive custom metadata fields for a document.
                    /// </summary>
                    /// <param name="resourceId">The resource_id returned by the upload call.</param>
                    /// <param name="dataTemplateId">Zoho data template ID.</param>
                    /// <param name="ktaFields">KTA field name → value dictionary from the document.</param>
                    /// <param name="fieldMapping">KTA field name → Zoho custom_field_id mapping.</param>
                    /// <param name="accessToken">Valid Zoho OAuth2 access token.</param>
                    /// <param name="metadataUrl">Zoho custom-metadata endpoint URL.</param>
        private async Task UpdateCustomMetadata(
string resourceId,
string dataTemplateId,
Dictionary<string, string> ktaFields,
Dictionary<string, string> fieldMapping,
string accessToken,
string metadataUrl)
        {
            // Build the custom_data array — include mapped fields (empty values allowed to clear).
            var customData = new List<object>();

            foreach (var mapping in fieldMapping)
            {
                string ktaFieldName = mapping.Key;   // e.g. "InvoiceNumber"
                string zohoFieldId = mapping.Value;  // e.g. "abc-MjcxODY1..."

                if (string.IsNullOrWhiteSpace(zohoFieldId)) continue;

                string value = "";
                ktaFields?.TryGetValue(ktaFieldName, out value);

                customData.Add(new
                {
                    custom_field_id = zohoFieldId,
                    value = value ?? ""
                });
            }

            if (customData.Count == 0) return;

            // Assemble the JSON:API request body
            var requestBody = new
            {
                data = new[]
{
new
{
attributes = new
{
resource_id = resourceId,
data_template_id = dataTemplateId,
custom_data = customData
},
type = "custommetadata"
}
}
            };

            string json = JsonConvert.SerializeObject(
            requestBody,
            Formatting.None,
            new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });

            var request = new HttpRequestMessage(HttpMethod.Post, metadataUrl.TrimEnd('/'));
            request.Headers.Authorization = new AuthenticationHeaderValue("Zoho-oauthtoken", accessToken);
            request.Headers.Accept.Clear();
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/vnd.api+json"));
            request.Content = new StringContent(json, Encoding.UTF8, "application/vnd.api+json");

            HttpResponseMessage response = await _httpClient.SendAsync(request);
            string body = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
                throw new Exception(
                $"Zoho custom metadata update failed for resource '{resourceId}' " +
                $"({(int)response.StatusCode}): {body}");
        }

        // =====================================================================
        // Field-mapping helper
        // =====================================================================

        /// <summary>
                    /// Parses "KtaField1:ZohoFieldId1;KtaField2:ZohoFieldId2"
                    /// into a Dictionary<KtaFieldName, ZohoCustomFieldId>.
                    /// Entries that do not contain exactly one ':' are silently skipped.
                    /// </summary>
        private Dictionary<string, string> ParseFieldMapping(string raw)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            if (string.IsNullOrWhiteSpace(raw)) return map;

            foreach (string entry in raw.Split(new[] { ';' },
             StringSplitOptions.RemoveEmptyEntries))
            {
                int colon = entry.IndexOf(':');
                if (colon <= 0 || colon == entry.Length - 1) continue;

                string ktaName = entry.Substring(0, colon).Trim();
                string zohoId = entry.Substring(colon + 1).Trim();

                if (!string.IsNullOrEmpty(ktaName) && !string.IsNullOrEmpty(zohoId))
                    map[ktaName] = zohoId;
            }

            return map;
        }

        // =====================================================================
        // KTA helpers
        // =====================================================================

        private List<FolderDocument> GetFolderDocuments(
string sessionId, string folderId, string ktaUrl)
        {
            string url = ktaUrl + "/CaptureDocumentService.svc/json/GetFolder";
            string body =
            "{\"sessionId\":\"" + EscapeJson(sessionId) + "\"," +
            "\"folderId\":\"" + EscapeJson(folderId) + "\"}";

            return ParseFolderResponse(PostJson(url, body, "GetFolder"));
        }

        private DocumentData GetDocumentData(
        string documentId, string ktaUrl, string sessionId)
        {
            string url = ktaUrl + "/CaptureDocumentService.svc/json/GetDocument";
            string body =
            "{\"sessionId\":\"" + EscapeJson(sessionId) + "\"," +
            "\"reportingData\":{\"Station\":\"\",\"MarkCompleted\":false}," +
            "\"documentId\":\"" + EscapeJson(documentId) + "\"}";

            return ParseDocumentResponse(PostJson(url, body, "GetDocument"));
        }

        private byte[] GetDocumentAsPdf(
        string documentId, string ktaUrl, string sessionId)
        {
            string url = ktaUrl + "/CaptureDocumentService.svc/json/GetDocumentFile2";
            string body =
            "{\"sessionId\":\"" + EscapeJson(sessionId) + "\"," +
            "\"reportingData\":{\"Station\":\"\",\"MarkCompleted\":false}," +
            "\"documentId\":\"" + EscapeJson(documentId) + "\"," +
            "\"documentFileOptions\":{\"FileType\":\"PDF\",\"IncludeAnnotations\":0}}";

            HttpWebRequest request = CreateWebRequest(url);
            using (var writer = new StreamWriter(request.GetRequestStream()))
            {
                writer.Write(body);
                writer.Flush();
            }

            try
            {
                using (HttpWebResponse response =
                  (HttpWebResponse)request.GetResponse())
                using (Stream stream = response.GetResponseStream()
                  ?? throw new Exception("PDF response stream is null."))
                using (MemoryStream ms = new MemoryStream())
                {
                    byte[] buffer = new byte[8192];
                    int count;
                    while ((count = stream.Read(buffer, 0, buffer.Length)) > 0)
                        ms.Write(buffer, 0, count);
                    return ms.ToArray();
                }
            }
            catch (WebException ex)
            {
                HandleWebException(ex, "GetDocumentFile2(PDF)");
                throw;
            }
        }

        // =====================================================================
        // Zoho auth + upload helpers
        // =====================================================================

        private async Task<string> GetZohoToken(ZohoConfig zoho)
        {
            lock (_tokenLock)
            {
                if (_cachedClientId == zoho.ClientId
                && !string.IsNullOrEmpty(_cachedZohoToken)
                && DateTime.UtcNow < _zohoTokenExpiry)
                    return _cachedZohoToken;
            }

            var content = new FormUrlEncodedContent(new[]
            {
new KeyValuePair<string, string>("client_id",   zoho.ClientId),
new KeyValuePair<string, string>("client_secret", zoho.ClientSecret),
new KeyValuePair<string, string>("grant_type",  "client_credentials"),
new KeyValuePair<string, string>("scope",     zoho.Scope)
});

            var response = await _httpClient.PostAsync(
            zoho.AccountsUrl.TrimEnd('/') + "/oauth/v2/token", content);

            string body = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
                throw new Exception("Zoho token request failed: " + body);

            JObject json = JObject.Parse(body);
            string token = json["access_token"]?.ToString();

            if (string.IsNullOrEmpty(token))
                throw new Exception("access_token missing in Zoho response: " + body);

            lock (_tokenLock)
            {
                _cachedZohoToken = token;
                _cachedClientId = zoho.ClientId;
                int exp = json["expires_in"]?.ToObject<int>() ?? 3600;
                _zohoTokenExpiry = DateTime.UtcNow.AddSeconds(exp - 300);
            }

            return token;
        }

        private async Task<string> UploadToZoho(
        byte[] bytes,
        string fileName,
        string mimeType,
        string accessToken,
        ZohoConfig zoho)
        {
            string url =
            zoho.UploadUrl.TrimEnd('/') +
            $"?filename={Uri.EscapeDataString(fileName)}" +
            $"&parent_id={Uri.EscapeDataString(zoho.FolderId)}" +
            $"&override-name-exist=true";

            using (var form = new MultipartFormDataContent())
            {
                var content = new ByteArrayContent(bytes);
                content.Headers.ContentType =
                MediaTypeHeaderValue.Parse(mimeType ?? "application/octet-stream");
                form.Add(content, "content", fileName);

                var request = new HttpRequestMessage(HttpMethod.Post, url);
                request.Headers.Authorization =
                new AuthenticationHeaderValue("Zoho-oauthtoken", accessToken);
                request.Content = form;

                var response = await _httpClient.SendAsync(request);
                string body = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                    throw new Exception(
                    $"Zoho upload failed for '{fileName}' " +
                    $"({(int)response.StatusCode}): {body}");

                try
                {
                    JObject json = JObject.Parse(body);
                    return json["data"]?[0]?["attributes"]?["resource_id"]
                      ?.ToString() ?? "";
                }
                catch { return ""; }
            }
        }

        // =====================================================================
        // KTA response parsers
        // =====================================================================

        private List<FolderDocument> ParseFolderResponse(string response)
        {
            var list = new List<FolderDocument>();
            if (string.IsNullOrEmpty(response)) return list;

            JObject json;
            try { json = JObject.Parse(response); }
            catch { return list; }

            JToken docs =
            json["GetFolderResult"]?["Documents"]
            ?? json["Documents"]
            ?? json["d"]?["Documents"];

            if (docs == null) return list;

            foreach (JToken d in docs)
            {
                if (d == null) continue;
                string id = d["Id"]?.ToString();
                if (string.IsNullOrEmpty(id)) continue;

                list.Add(new FolderDocument
                {
                    Id = id,
                    FileName = d["FileName"]?.ToString() ?? $"Document_{id}.pdf"
                });
            }

            return list;
        }

        private DocumentData ParseDocumentResponse(string response)
        {
            var doc = new DocumentData();
            if (string.IsNullOrEmpty(response)) return doc;

            JObject json;
            try { json = JObject.Parse(response); }
            catch { return doc; }

            JToken root = json["d"];
            if (root == null) return doc;

            doc.Id = SafeStr(root["Id"]);
            doc.FileName = SafeStr(root["FileName"]);
            doc.DocumentName = SafeStr(root["Name"]);
            doc.NumberOfPages = root["NumberOfPages"]?.ToObject<int>() ?? 0;
            doc.IsValid = root["Valid"]?.ToObject<bool>() ?? false;
            doc.IsVerified = root["Verified"]?.ToObject<bool>() ?? false;
            doc.CreatedAt = ParseKtaDate(root["CreatedAt"]?.ToString()) ?? "";
            doc.DocumentTypeName = SafeStr(root["DocumentType"]?["Identity"]?["Name"]);

            JToken fields = root["Fields"];
            if (fields == null) return doc;

            foreach (JToken f in fields)
            {
                if (f == null) continue;
                string name = f["Name"]?.ToString();
                if (string.IsNullOrEmpty(name)) continue;

                int type = f["FieldType"]?.ToObject<int>() ?? 0;

                if (type == 5)
                    doc.TableFields[name] = ExtractTableField(f);
                else
                    doc.Fields[name] = ExtractScalarField(f, type) ?? "";
            }

            return doc;
        }

        private string ExtractScalarField(JToken field, int fieldType)
        {
            if (field == null) return "";

            if (fieldType == 2)
            {
                JToken d = field["DoubleValue"];
                if (d != null && d.Type != JTokenType.Null) return d.ToString();
            }

            if (fieldType == 3)
            {
                string parsed = ParseKtaDate(field["DateValue"]?.ToString());
                if (!string.IsNullOrEmpty(parsed)) return parsed;
            }

            JToken val = field["Value"];
            if (val != null
            && val.Type != JTokenType.Null
            && val.Type != JTokenType.Object)
                return val.ToString();

            JToken dbl = field["DoubleValue"];
            if (dbl != null && dbl.Type != JTokenType.Null)
                return dbl.ToString();

            return "";
        }

        private List<Dictionary<string, string>> ExtractTableField(JToken field)
        {
            var rows = new List<Dictionary<string, string>>();
            if (field == null) return rows;

            JToken rowsToken = field["Table"]?["Rows"];
            if (rowsToken == null) return rows;

            foreach (JToken row in rowsToken)
            {
                if (row == null) continue;
                var rowDict = new Dictionary<string, string>();
                JToken cells = row["Cells"];

                if (cells != null)
                {
                    foreach (JToken cell in cells)
                    {
                        if (cell == null) continue;
                        int col = cell["Column"]?.ToObject<int>() ?? 0;
                        string colKey = $"Col_{col}";
                        string val = "";

                        JToken v = cell["Value"];
                        if (v != null && v.Type != JTokenType.Null
                        && v.Type != JTokenType.Object)
                        {
                            val = v.ToString();
                        }
                        else
                        {
                            JToken d = cell["DoubleValue"];
                            if (d != null && d.Type != JTokenType.Null)
                                val = d.ToString();
                            else
                                val = ParseKtaDate(cell["DateValue"]?.ToString()) ?? "";
                        }

                        rowDict[colKey] = val;
                    }
                }

                rows.Add(rowDict);
            }

            return rows;
        }

        // =====================================================================
        // Low-level HTTP / utility helpers
        // =====================================================================

        private string ParseKtaDate(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return null;
            int s = raw.IndexOf('(');
            if (s < 0) return null;

            int p = raw.IndexOf('+', s + 1);
            int m = raw.IndexOf('-', s + 1);
            int e = (p > 0 && (m < 0 || p < m)) ? p : m;
            if (e < 0) return null;

            string tick = raw.Substring(s + 1, e - s - 1);
            if (!long.TryParse(tick, out long ms) || ms == 0) return null;

            try
            {
                return DateTimeOffset.FromUnixTimeMilliseconds(ms)
                .UtcDateTime.ToString("yyyy-MM-dd");
            }
            catch { return null; }
        }

        private string PostJson(string url, string body, string op)
        {
            HttpWebRequest req = CreateWebRequest(url);
            using (var w = new StreamWriter(req.GetRequestStream()))
            {
                w.Write(body);
                w.Flush();
            }

            try
            {
                using (var resp = (HttpWebResponse)req.GetResponse())
                using (var stream = resp.GetResponseStream()
                  ?? throw new Exception($"{op} stream is null."))
                using (var reader = new StreamReader(stream))
                    return reader.ReadToEnd() ?? "";
            }
            catch (WebException ex)
            {
                HandleWebException(ex, op);
                throw;
            }
        }

        private HttpWebRequest CreateWebRequest(string url)
        {
            var req = (HttpWebRequest)WebRequest.Create(url);
            req.ContentType = "application/json";
            req.Method = "POST";
            req.Proxy = null;
            return req;
        }

        private void HandleWebException(WebException ex, string op)
        {
            if (ex?.Response is HttpWebResponse r)
            {
                string err = "";
                try
                {
                    using (var s = r.GetResponseStream())
                    using (var rd = new StreamReader(s))
                        err = rd.ReadToEnd();
                }
                catch { }

                switch (r.StatusCode)
                {
                    case HttpStatusCode.Unauthorized:
                        throw new Exception(
                        $"401 Unauthorized — {op}: session expired or invalid credentials.");
                    case HttpStatusCode.Forbidden:
                        throw new Exception(
                        $"403 Forbidden — {op}: insufficient permissions.");
                    case HttpStatusCode.NotFound:
                        throw new Exception(
                        $"404 Not Found — {op}: check URL and IDs.");
                    default:
                        throw new Exception(
                        $"{op} failed ({(int)r.StatusCode}): {err}");
                }
            }
            throw new Exception($"{op} network error: {ex?.Message}");
        }

        private string SafeStr(JToken t)
        => (t == null || t.Type == JTokenType.Null) ? "" : t.ToString() ?? "";

        private string EnsurePdfExtension(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "Document.pdf";
            string ext = Path.GetExtension(name);
            if (string.IsNullOrEmpty(ext)) return name + ".pdf";
            if (!ext.Equals(".pdf", StringComparison.OrdinalIgnoreCase))
                return Path.GetFileNameWithoutExtension(name) + ".pdf";
            return name;
        }

        private string EscapeJson(string v)
        {
            if (string.IsNullOrEmpty(v)) return "";
            return v.Replace("\\", "\\\\").Replace("\"", "\\\"")
            .Replace("\n", "\\n").Replace("\r", "\\r")
            .Replace("\t", "\\t");
        }

        // ── Private nested classes — completely hidden ─────────────────────────

        private class KtaConfig
        {
            public string SessionId { get; set; }
            public string Url { get; set; }
            public string FolderId { get; set; }

            public void Validate()
            {
                if (string.IsNullOrWhiteSpace(SessionId))
                    throw new ArgumentException("sessionId is required.");
                if (string.IsNullOrWhiteSpace(Url))
                    throw new ArgumentException("ktaUrl is required.");
                if (string.IsNullOrWhiteSpace(FolderId))
                    throw new ArgumentException("ktaFolderId is required.");
            }
        }

        private class ZohoConfig
        {
            public string AccountsUrl { get; set; }
            public string UploadUrl { get; set; }
            public string MetadataUrl { get; set; }
            public string ClientId { get; set; }
            public string ClientSecret { get; set; }
            public string FolderId { get; set; }
            public string Scope { get; set; }
            public string DataTemplateId { get; set; }
            public Dictionary<string, string> FieldMapping { get; set; }

            public void Validate()
            {
                if (string.IsNullOrWhiteSpace(AccountsUrl))
                    throw new ArgumentException("zohoAccountsUrl is required.");
                if (string.IsNullOrWhiteSpace(UploadUrl))
                    throw new ArgumentException("zohoUploadUrl is required.");
                if (string.IsNullOrWhiteSpace(ClientId))
                    throw new ArgumentException("zohoClientId is required.");
                if (string.IsNullOrWhiteSpace(ClientSecret))
                    throw new ArgumentException("zohoClientSecret is required.");
                if (string.IsNullOrWhiteSpace(FolderId))
                    throw new ArgumentException("zohoFolderId is required.");
                if (string.IsNullOrWhiteSpace(Scope))
                    throw new ArgumentException("zohoScope is required.");

                // MetadataUrl and DataTemplateId are optional;
                // metadata update is simply skipped when absent.
            }
        }

        private class FolderDocument
        {
            public string Id { get; set; } = "";
            public string FileName { get; set; } = "";
        }

        private class DocumentData
        {
            public string Id { get; set; } = "";
            public string FileName { get; set; } = "";
            public string DocumentName { get; set; } = "";
            public string DocumentTypeName { get; set; } = "";
            public int NumberOfPages { get; set; } = 0;
            public bool IsValid { get; set; } = false;
            public bool IsVerified { get; set; } = false;
            public string CreatedAt { get; set; } = "";

            public Dictionary<string, string> Fields { get; set; }
            = new Dictionary<string, string>();

            public Dictionary<string, List<Dictionary<string, string>>> TableFields
            { get; set; }
            = new Dictionary<string, List<Dictionary<string, string>>>();
        }

        private class UploadItem
        {
            public bool Success { get; set; } = false;
            public bool MetadataUpdated { get; set; } = false;
            public string DocumentName { get; set; } = "";
            public string ZohoPdfFileId { get; set; } = "";
            public long PdfSizeBytes { get; set; } = 0;
            public string ErrorMessage { get; set; } = "";

            public Dictionary<string, string> Fields { get; set; }
            = new Dictionary<string, string>();
        }
    }
}