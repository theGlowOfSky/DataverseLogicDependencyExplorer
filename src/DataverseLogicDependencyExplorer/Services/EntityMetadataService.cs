// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Query;
using DataverseLogicDependencyExplorer.Models;

namespace DataverseLogicDependencyExplorer.Services
{
    public enum TableKindFilter { All, Custom, Standard }

    /// <summary>
    /// Loads table and column metadata with supported metadata requests (late-bound, no generated proxy classes) and caches it
    /// for the session. Call <see cref="Invalidate"/> (or pass forceRefresh) to reload.
    /// </summary>
    public class EntityMetadataService
    {
        private readonly object sync = new object();
        private List<TableInfo> tables;
        private readonly Dictionary<string, List<ColumnInfo>> columns = new Dictionary<string, List<ColumnInfo>>(StringComparer.OrdinalIgnoreCase);

        public bool HasTables { get { lock (sync) { return tables != null; } } }
        public IReadOnlyList<TableInfo> CachedTables { get { lock (sync) { return tables ?? new List<TableInfo>(); } } }

        public void Invalidate()
        {
            lock (sync)
            {
                tables = null;
                columns.Clear();
            }
        }

        public TableInfo FindTable(string logicalName)
        {
            lock (sync)
            {
                if (tables == null || string.IsNullOrEmpty(logicalName)) return null;
                return tables.FirstOrDefault(t => string.Equals(t.LogicalName, logicalName, StringComparison.OrdinalIgnoreCase));
            }
        }

        /// <summary>Loads (or returns the cached) list of tables that are valid analysis targets.</summary>
        public List<TableInfo> LoadTables(IOrganizationService svc, bool forceRefresh, CancellationToken ct)
        {
            lock (sync) { if (!forceRefresh && tables != null) return tables; }

            var result = new List<TableInfo>();
            IEnumerable<EntityMetadata> all;
            try
            {
                var query = new EntityQueryExpression { Properties = new MetadataPropertiesExpression { AllProperties = false } };
                query.Properties.PropertyNames.AddRange(new[]
                {
                    "LogicalName", "DisplayName", "IsCustomEntity", "IsIntersect", "IsPrivate", "IsLogicalEntity", "EntitySetName"
                });
                var resp = (RetrieveMetadataChangesResponse)DataverseHelper.Retry(
                    () => svc.Execute(new RetrieveMetadataChangesRequest { Query = query }), ct);
                all = resp.EntityMetadata;
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceWarning("RetrieveMetadataChanges failed (" + ex.Message + "); falling back to RetrieveAllEntities.");
                var resp = (RetrieveAllEntitiesResponse)DataverseHelper.Retry(
                    () => svc.Execute(new RetrieveAllEntitiesRequest { EntityFilters = EntityFilters.Entity, RetrieveAsIfPublished = false }), ct);
                all = resp.EntityMetadata;
            }

            foreach (EntityMetadata em in all)
            {
                ct.ThrowIfCancellationRequested();
                TableInfo t = MapTable(em.LogicalName, LabelOf(em.DisplayName), em.IsCustomEntity, em.IsIntersect, em.IsPrivate, em.IsLogicalEntity, em.EntitySetName);
                if (t != null) result.Add(t);
            }

            result.Sort((a, b) => string.Compare(a.LogicalName, b.LogicalName, StringComparison.OrdinalIgnoreCase));
            lock (sync) { tables = result; }
            return result;
        }

        /// <summary>Loads (or returns cached) columns of one table. Throws when the table metadata cannot be read.</summary>
        public List<ColumnInfo> LoadColumns(IOrganizationService svc, string tableLogicalName, bool forceRefresh, CancellationToken ct)
        {
            List<ColumnInfo> cached;
            lock (sync) { if (!forceRefresh && columns.TryGetValue(tableLogicalName, out cached)) return cached; }

            var resp = (RetrieveEntityResponse)DataverseHelper.Retry(() => svc.Execute(new RetrieveEntityRequest
            {
                LogicalName = tableLogicalName,
                EntityFilters = EntityFilters.Attributes,
                RetrieveAsIfPublished = false
            }), ct);

            var list = new List<ColumnInfo>();
            if (resp.EntityMetadata != null && resp.EntityMetadata.Attributes != null)
            {
                foreach (AttributeMetadata a in resp.EntityMetadata.Attributes)
                {
                    string type = a.AttributeTypeName != null ? a.AttributeTypeName.Value : (a.AttributeType.HasValue ? a.AttributeType.Value.ToString() : null);
                    ColumnInfo c = MapColumn(a.LogicalName, LabelOf(a.DisplayName), type, a.IsCustomAttribute, a.AttributeOf);
                    if (c != null) list.Add(c);
                }
            }
            list.Sort((x, y) => string.Compare(x.LogicalName, y.LogicalName, StringComparison.OrdinalIgnoreCase));
            lock (sync) { columns[tableLogicalName] = list; }
            return list;
        }

        // ------------------------------------------------------------------ pure mapping/filter logic (unit tested)

        /// <summary>Maps raw metadata values to a selectable table, or null when the table is not a valid analysis target.</summary>
        public static TableInfo MapTable(string logicalName, string displayName, bool? isCustom, bool? isIntersect, bool? isPrivate, bool? isLogical, string entitySetName)
        {
            if (string.IsNullOrWhiteSpace(logicalName)) return null;
            if (isPrivate == true || isLogical == true) return null;   // internal / logical tables cannot hold customisations
            if (isIntersect == true) return null;                      // N:N relationship tables have no forms, rules or flows of their own
            return new TableInfo
            {
                LogicalName = logicalName,
                DisplayName = string.IsNullOrWhiteSpace(displayName) ? null : displayName,
                IsCustom = isCustom == true,
                IsIntersect = false,
                EntitySetName = entitySetName
            };
        }

        /// <summary>Maps raw attribute metadata to a selectable column; shadow columns (AttributeOf set, e.g. "xxxname") are excluded.</summary>
        public static ColumnInfo MapColumn(string logicalName, string displayName, string attributeType, bool? isCustom, string attributeOf)
        {
            if (string.IsNullOrWhiteSpace(logicalName)) return null;
            if (!string.IsNullOrEmpty(attributeOf)) return null;
            return new ColumnInfo
            {
                LogicalName = logicalName,
                DisplayName = string.IsNullOrWhiteSpace(displayName) ? null : displayName,
                AttributeType = attributeType,
                IsCustom = isCustom == true
            };
        }

        public static List<TableInfo> FilterTables(IEnumerable<TableInfo> source, string text, TableKindFilter kind)
        {
            string f = (text ?? string.Empty).Trim();
            return source.Where(t =>
                    (kind == TableKindFilter.All || (kind == TableKindFilter.Custom ? t.IsCustom : !t.IsCustom)) &&
                    (f.Length == 0
                     || (t.LogicalName ?? string.Empty).IndexOf(f, StringComparison.OrdinalIgnoreCase) >= 0
                     || (t.DisplayName ?? string.Empty).IndexOf(f, StringComparison.OrdinalIgnoreCase) >= 0))
                .ToList();
        }

        public static List<ColumnInfo> FilterColumns(IEnumerable<ColumnInfo> source, string text)
        {
            string f = (text ?? string.Empty).Trim();
            return source.Where(c => f.Length == 0
                     || (c.LogicalName ?? string.Empty).IndexOf(f, StringComparison.OrdinalIgnoreCase) >= 0
                     || (c.DisplayName ?? string.Empty).IndexOf(f, StringComparison.OrdinalIgnoreCase) >= 0)
                .ToList();
        }

        private static string LabelOf(Label label)
        {
            if (label == null) return null;
            if (label.UserLocalizedLabel != null && !string.IsNullOrWhiteSpace(label.UserLocalizedLabel.Label)) return label.UserLocalizedLabel.Label;
            LocalizedLabel first = label.LocalizedLabels == null ? null : label.LocalizedLabels.FirstOrDefault();
            return first == null ? null : first.Label;
        }
    }
}
