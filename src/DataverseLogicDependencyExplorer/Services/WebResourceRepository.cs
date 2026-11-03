// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace DataverseLogicDependencyExplorer.Services
{
    public class WebResourceContent
    {
        public Guid Id { get; set; }
        public string Name { get; set; }
        public string DisplayName { get; set; }
        public string Text { get; set; }
        public string EncodingName { get; set; }
        /// <summary>Set when the content could not be decoded or was empty.</summary>
        public string Problem { get; set; }
        public bool? IsManaged { get; set; }
    }

    /// <summary>A web resource without its content: cheap to list, used to plan the download.</summary>
    public class WebResourceHeader
    {
        public Guid Id { get; set; }
        public string Name { get; set; }
        public bool? IsManaged { get; set; }
    }

    /// <summary>
    /// Retrieves script web resources (type 3). Listing is separated from downloading so content can be fetched in parallel chunks
    /// and analysed while it arrives. Only scripts that other scanners need again (libraries registered on forms) are retained;
    /// everything else is released after analysis to keep memory bounded. Thread-safe.
    /// </summary>
    public class WebResourceRepository
    {
        private const int ScriptType = 3;
        private const int NameChunk = 25;

        private readonly IOrganizationService svc;
        private readonly object sync = new object();
        private readonly Dictionary<string, WebResourceContent> byName = new Dictionary<string, WebResourceContent>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> queriedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public WebResourceRepository(IOrganizationService svc) { this.svc = svc; }

        public int RetainedCount { get { lock (sync) return byName.Count; } }

        /// <summary>Lists every script web resource (id, name, managed flag) without downloading content.</summary>
        public List<WebResourceHeader> ListScriptHeaders(CancellationToken ct, bool skipManaged, List<string> warnings)
        {
            Func<bool, QueryExpression> build = useState =>
            {
                var q = new QueryExpression("webresource") { ColumnSet = new ColumnSet("webresourceid", "name", "ismanaged") };
                q.Criteria.AddCondition("webresourcetype", ConditionOperator.Equal, ScriptType);
                if (skipManaged) q.Criteria.AddCondition("ismanaged", ConditionOperator.Equal, false);
                if (useState) q.Criteria.AddCondition("componentstate", ConditionOperator.Equal, 0); // published layer only
                q.AddOrder("name", OrderType.Ascending);
                return q;
            };

            List<Entity> rows;
            try { rows = DataverseHelper.RetrieveAll(svc, build(true), ct, 5000).ToList(); }
            catch (FaultException<OrganizationServiceFault>)
            {
                if (warnings != null) warnings.Add("The web resource listing with the published-state filter failed; it was retried without it (unpublished duplicates are possible).");
                rows = DataverseHelper.RetrieveAll(svc, build(false), ct, 5000).ToList();
            }
            return rows.Select(e => new WebResourceHeader
            {
                Id = e.Id,
                Name = DataverseHelper.Str(e, "name"),
                IsManaged = DataverseHelper.Bool(e, "ismanaged")
            }).Where(h => !string.IsNullOrEmpty(h.Name)).ToList();
        }

        /// <summary>Downloads and decodes the content of the given web resources using the supplied connection (one per worker thread).</summary>
        public List<WebResourceContent> DownloadByIds(IOrganizationService connection, IList<Guid> ids, CancellationToken ct)
        {
            var result = new List<WebResourceContent>();
            if (ids == null || ids.Count == 0) return result;

            var q = new QueryExpression("webresource") { ColumnSet = new ColumnSet("webresourceid", "name", "displayname", "content", "ismanaged") };
            q.Criteria.AddCondition("webresourceid", ConditionOperator.In, ids.Cast<object>().ToArray());
            foreach (Entity e in DataverseHelper.RetrieveAll(connection, q, ct, Math.Max(ids.Count, 1)))
            {
                WebResourceContent w = ToContent(e);
                if (w != null) result.Add(w);
            }
            return result;
        }

        /// <summary>Keeps a script for later scanners (e.g. libraries registered on forms).</summary>
        public void Retain(WebResourceContent w)
        {
            if (w == null || string.IsNullOrEmpty(w.Name)) return;
            lock (sync) byName[w.Name] = w;
        }

        /// <summary>Loads the named scripts that are not retained yet. Names not found stay absent (see <see cref="TryGet"/>).</summary>
        public void EnsureNames(IEnumerable<string> names, CancellationToken ct, List<string> warnings)
        {
            List<string> missing;
            lock (sync)
                missing = names.Where(n => !string.IsNullOrWhiteSpace(n) && !byName.ContainsKey(n) && !queriedNames.Contains(n))
                               .Distinct(StringComparer.OrdinalIgnoreCase).ToList();

            foreach (List<string> chunk in DataverseHelper.Chunk(missing, NameChunk))
            {
                foreach (Entity e in RetrieveByName(ct, chunk, warnings)) Retain(ToContent(e));
                lock (sync) foreach (string n in chunk) queriedNames.Add(n);
            }
        }

        public WebResourceContent TryGet(string name)
        {
            lock (sync)
            {
                WebResourceContent w;
                return name != null && byName.TryGetValue(name, out w) ? w : null;
            }
        }

        private IEnumerable<Entity> RetrieveByName(CancellationToken ct, List<string> names, List<string> warnings)
        {
            Func<bool, QueryExpression> build = useState =>
            {
                var q = new QueryExpression("webresource") { ColumnSet = new ColumnSet("webresourceid", "name", "displayname", "content", "ismanaged") };
                q.Criteria.AddCondition("webresourcetype", ConditionOperator.Equal, ScriptType);
                q.Criteria.AddCondition("name", ConditionOperator.In, names.Cast<object>().ToArray());
                if (useState) q.Criteria.AddCondition("componentstate", ConditionOperator.Equal, 0);
                q.AddOrder("name", OrderType.Ascending);
                return q;
            };

            try { return DataverseHelper.RetrieveAll(svc, build(true), ct, 50).ToList(); }
            catch (FaultException<OrganizationServiceFault>)
            {
                if (warnings != null) warnings.Add("The web resource query with the published-state filter failed; it was retried without it (unpublished duplicates are possible).");
                return DataverseHelper.RetrieveAll(svc, build(false), ct, 50).ToList();
            }
        }

        public static WebResourceContent ToContent(Entity e)
        {
            string name = DataverseHelper.Str(e, "name");
            if (string.IsNullOrEmpty(name)) return null;
            var w = new WebResourceContent
            {
                Id = e.Id,
                Name = name,
                DisplayName = DataverseHelper.Str(e, "displayname"),
                IsManaged = DataverseHelper.Bool(e, "ismanaged")
            };

            string base64 = DataverseHelper.Str(e, "content");
            if (string.IsNullOrEmpty(base64)) w.Problem = "The web resource has no content.";
            else
            {
                try
                {
                    string encoding;
                    w.Text = DecodeText(Convert.FromBase64String(base64), out encoding);
                    w.EncodingName = encoding;
                }
                catch (FormatException) { w.Problem = "The web resource content is not valid base64."; }
                catch (Exception ex) { w.Problem = "The web resource content could not be decoded: " + ex.Message; }
            }
            return w;
        }

        /// <summary>Decodes script bytes honouring byte-order marks (UTF-8, UTF-16 LE/BE, UTF-32) and falling back to Windows-1252.</summary>
        public static string DecodeText(byte[] bytes, out string encodingName)
        {
            if (bytes == null || bytes.Length == 0) { encodingName = "empty"; return string.Empty; }

            if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
            { encodingName = "UTF-8 (BOM)"; return Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3); }
            if (bytes.Length >= 4 && bytes[0] == 0xFF && bytes[1] == 0xFE && bytes[2] == 0 && bytes[3] == 0)
            { encodingName = "UTF-32 LE"; return new UTF32Encoding(false, true).GetString(bytes, 4, bytes.Length - 4); }
            if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
            { encodingName = "UTF-16 LE"; return Encoding.Unicode.GetString(bytes, 2, bytes.Length - 2); }
            if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
            { encodingName = "UTF-16 BE"; return Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2); }

            try
            {
                string text = new UTF8Encoding(false, true).GetString(bytes);
                encodingName = "UTF-8";
                return text;
            }
            catch (DecoderFallbackException)
            {
                encodingName = "Windows-1252 (fallback)";
                return Encoding.GetEncoding(1252).GetString(bytes);
            }
        }
    }
}
