// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Threading;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace DataverseLogicDependencyExplorer.Services
{
    /// <summary>Paging, throttling-aware retries, permission detection and safe attribute readers shared by all scanners.</summary>
    public static class DataverseHelper
    {
        public const int DefaultPageSize = 50;

        // Dataverse service-protection limit error codes (requests / execution time / concurrent requests).
        private static readonly int[] ThrottleCodes = { -2147015902, -2147015903, -2147015898 };
        // SecLib::AccessCheckEx failed - the user lacks a required privilege.
        private const int PrivilegeDeniedCode = -2147220960;

        /// <summary>Lazily pages through every record of a query (never loads an unbounded single page).</summary>
        public static IEnumerable<Entity> RetrieveAll(IOrganizationService svc, QueryExpression query, CancellationToken ct, int pageSize = DefaultPageSize)
        {
            query.PageInfo = new PagingInfo { Count = pageSize, PageNumber = 1, PagingCookie = null };
            while (true)
            {
                ct.ThrowIfCancellationRequested();
                EntityCollection page = Retry(() => svc.RetrieveMultiple(query), ct);
                foreach (Entity e in page.Entities) yield return e;
                if (!page.MoreRecords) yield break;
                query.PageInfo.PageNumber++;
                query.PageInfo.PagingCookie = page.PagingCookie;
            }
        }

        /// <summary>Retries on Dataverse throttling (honouring Retry-After) and once on a timeout.</summary>
        public static T Retry<T>(Func<T> action, CancellationToken ct)
        {
            for (int attempt = 0; ; attempt++)
            {
                try { return action(); }
                catch (FaultException<OrganizationServiceFault> ex) when (attempt < 4 && ThrottleCodes.Contains(ex.Detail.ErrorCode))
                {
                    TimeSpan wait = TimeSpan.FromSeconds(5 * (attempt + 1));
                    object retryAfter = null;
                    if (ex.Detail.ErrorDetails != null && ex.Detail.ErrorDetails.TryGetValue("Retry-After", out retryAfter) && retryAfter is TimeSpan)
                        wait = (TimeSpan)retryAfter;
                    if (ct.WaitHandle.WaitOne(wait)) ct.ThrowIfCancellationRequested();
                }
                catch (TimeoutException) when (attempt < 1)
                {
                    ct.ThrowIfCancellationRequested();
                }
            }
        }

        /// <summary>True when the failure means the connected user is not allowed to read the data.</summary>
        public static bool IsPermissionDenied(Exception ex)
        {
            for (Exception e = ex; e != null; e = e.InnerException)
            {
                var fault = e as FaultException<OrganizationServiceFault>;
                if (fault != null && fault.Detail != null && fault.Detail.ErrorCode == PrivilegeDeniedCode) return true;
                if (e is UnauthorizedAccessException) return true;

                string m = e.Message ?? string.Empty;
                if (m.IndexOf("prvRead", StringComparison.OrdinalIgnoreCase) >= 0
                    || m.IndexOf("AccessCheckEx", StringComparison.OrdinalIgnoreCase) >= 0
                    || m.IndexOf("SecLib", StringComparison.OrdinalIgnoreCase) >= 0
                    || m.IndexOf("access is denied", StringComparison.OrdinalIgnoreCase) >= 0
                    || m.IndexOf("does not have", StringComparison.OrdinalIgnoreCase) >= 0 && m.IndexOf("privilege", StringComparison.OrdinalIgnoreCase) >= 0
                    || m.IndexOf("Forbidden", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            return false;
        }

        /// <summary>True when the failure indicates an attribute or entity that does not exist in this environment.</summary>
        public static bool IsMissingMetadata(Exception ex)
        {
            for (Exception e = ex; e != null; e = e.InnerException)
            {
                var fault = e as FaultException<OrganizationServiceFault>;
                if (fault != null && fault.Detail != null && (fault.Detail.ErrorCode == -2147217149 || fault.Detail.ErrorCode == -2147220969))
                    return true;
                string m = e.Message ?? string.Empty;
                if (m.IndexOf("does not exist", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    m.IndexOf("Could not find", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    m.IndexOf("doesn't contain attribute", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            return false;
        }

        public static string Describe(Exception ex)
        {
            if (ex == null) return string.Empty;
            var fault = ex as FaultException<OrganizationServiceFault>;
            if (fault != null && fault.Detail != null && !string.IsNullOrEmpty(fault.Detail.Message))
                return fault.Detail.Message;
            return ex.Message;
        }

        // ---------------------------------------------------------------- attribute readers

        public static string Str(Entity e, string attribute)
        {
            if (e == null || !e.Contains(attribute)) return null;
            object v = e[attribute];
            var aliased = v as AliasedValue;
            if (aliased != null) v = aliased.Value;
            var er = v as EntityReference;
            if (er != null) return er.Name ?? er.Id.ToString();
            var os = v as OptionSetValue;
            if (os != null) return os.Value.ToString();
            return v == null ? null : v.ToString();
        }

        public static int? Opt(Entity e, string attribute)
        {
            if (e == null || !e.Contains(attribute)) return null;
            object v = e[attribute];
            var aliased = v as AliasedValue;
            if (aliased != null) v = aliased.Value;
            var os = v as OptionSetValue;
            if (os != null) return os.Value;
            if (v is int) return (int)v;
            return null;
        }

        public static bool? Bool(Entity e, string attribute)
        {
            if (e == null || !e.Contains(attribute)) return null;
            object v = e[attribute];
            var aliased = v as AliasedValue;
            if (aliased != null) v = aliased.Value;
            if (v is bool) return (bool)v;
            return null;
        }

        public static EntityReference Ref(Entity e, string attribute)
        {
            if (e == null || !e.Contains(attribute)) return null;
            object v = e[attribute];
            var aliased = v as AliasedValue;
            if (aliased != null) v = aliased.Value;
            return v as EntityReference;
        }

        // ---------------------------------------------------------------- misc

        public static IEnumerable<List<T>> Chunk<T>(IEnumerable<T> source, int size)
        {
            var current = new List<T>(size);
            foreach (T item in source)
            {
                current.Add(item);
                if (current.Count == size) { yield return current; current = new List<T>(size); }
            }
            if (current.Count > 0) yield return current;
        }

        public static string Truncate(string text, int max)
        {
            if (string.IsNullOrEmpty(text) || text.Length <= max) return text;
            return text.Substring(0, max) + "... [" + (text.Length - max) + " more characters]";
        }

        /// <summary>Collapses whitespace so multi-line snippets fit one grid line.</summary>
        public static string OneLine(string text)
        {
            return string.IsNullOrEmpty(text) ? string.Empty : System.Text.RegularExpressions.Regex.Replace(text, @"\s+", " ").Trim();
        }

        public static string StateText(int? state)
        {
            if (!state.HasValue) return null;
            return state.Value == 0 ? "Draft/Active" : state.Value == 1 ? "Activated/Inactive" : state.Value.ToString();
        }
    }
}
