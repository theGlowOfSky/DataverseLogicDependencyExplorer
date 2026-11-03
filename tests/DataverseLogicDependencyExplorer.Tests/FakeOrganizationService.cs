// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace DataverseLogicDependencyExplorer.Tests
{
    /// <summary>
    /// Isolated stand-in for the Dataverse connection: returns canned rows per table and records the queries it receives so tests
    /// can assert that targeted queries (table filters, paging) were issued. No live environment is contacted.
    /// </summary>
    public class FakeOrganizationService : IOrganizationService
    {
        public readonly Dictionary<string, Func<QueryExpression, IEnumerable<Entity>>> Handlers =
            new Dictionary<string, Func<QueryExpression, IEnumerable<Entity>>>(StringComparer.OrdinalIgnoreCase);
        public readonly List<QueryExpression> Queries = new List<QueryExpression>();

        public FakeOrganizationService On(string entity, params Entity[] rows)
        {
            Handlers[entity] = q => rows;
            return this;
        }

        public FakeOrganizationService OnThrow(string entity, Exception ex)
        {
            Handlers[entity] = q => { throw ex; };
            return this;
        }

        /// <summary>Artificial latency per query, to make parallel workers overlap in tests.</summary>
        public int DelayMs;

        /// <summary>
        /// Web resources: honours the conditions the plugin sends (webresourceid In, name In, ismanaged = false) and only returns the
        /// heavy 'content' column when it was requested, like Dataverse would.
        /// </summary>
        public FakeOrganizationService OnWebResources(params Entity[] rows)
        {
            Handlers["webresource"] = q =>
            {
                IEnumerable<Entity> r = rows;
                foreach (ConditionExpression c in q.Criteria.Conditions)
                {
                    if (c.AttributeName == "webresourceid" && c.Operator == ConditionOperator.In)
                    { var ids = new HashSet<Guid>(c.Values.Cast<Guid>()); r = r.Where(e => ids.Contains(e.Id)); }
                    else if (c.AttributeName == "name" && c.Operator == ConditionOperator.In)
                    { var names = new HashSet<string>(c.Values.Cast<string>(), StringComparer.OrdinalIgnoreCase); r = r.Where(e => names.Contains((string)e["name"])); }
                    else if (c.AttributeName == "ismanaged" && c.Operator == ConditionOperator.Equal)
                    { bool want = (bool)c.Values[0]; r = r.Where(e => e.Contains("ismanaged") && (bool)e["ismanaged"] == want); }
                }
                bool withContent = q.ColumnSet.Columns.Contains("content");
                return r.Select(e =>
                {
                    var copy = new Entity(e.LogicalName, e.Id);
                    foreach (var kv in e.Attributes) if (withContent || kv.Key != "content") copy[kv.Key] = kv.Value;
                    return copy;
                }).ToList();
            };
            return this;
        }

        public EntityCollection RetrieveMultiple(QueryBase query)
        {
            var qe = (QueryExpression)query;
            lock (Queries) Queries.Add(qe);
            if (DelayMs > 0) System.Threading.Thread.Sleep(DelayMs);
            Func<QueryExpression, IEnumerable<Entity>> h;
            var result = new EntityCollection();
            if (Handlers.TryGetValue(qe.EntityName, out h)) result.Entities.AddRange(h(qe));
            result.MoreRecords = false;
            return result;
        }

        public static FaultException<OrganizationServiceFault> Fault(int code, string message)
        {
            return new FaultException<OrganizationServiceFault>(new OrganizationServiceFault { ErrorCode = code, Message = message }, message);
        }

        public static FaultException<OrganizationServiceFault> PermissionFault()
        {
            return Fault(-2147220960, "SecLib::AccessCheckEx failed. User is missing prvReadWorkflow privilege.");
        }

        public static Entity Row(string logicalName, Guid? id = null, params object[] pairs)
        {
            var e = new Entity(logicalName, id ?? Guid.NewGuid());
            for (int i = 0; i + 1 < pairs.Length; i += 2) e[(string)pairs[i]] = pairs[i + 1];
            return e;
        }

        public Guid Create(Entity entity) { throw new NotSupportedException("Read-only fake."); }
        public Entity Retrieve(string entityName, Guid id, ColumnSet columnSet) { throw new NotSupportedException(); }
        public void Update(Entity entity) { throw new NotSupportedException("Read-only fake."); }
        public void Delete(string entityName, Guid id) { throw new NotSupportedException("Read-only fake."); }
        public OrganizationResponse Execute(OrganizationRequest request) { throw new NotSupportedException(); }
        public void Associate(string entityName, Guid entityId, Relationship relationship, EntityReferenceCollection relatedEntities) { throw new NotSupportedException(); }
        public void Disassociate(string entityName, Guid entityId, Relationship relationship, EntityReferenceCollection relatedEntities) { throw new NotSupportedException(); }
    }
}
