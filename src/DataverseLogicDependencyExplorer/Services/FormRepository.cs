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
    public class FormRecord
    {
        public Guid Id { get; set; }
        public string Name { get; set; }
        public string Table { get; set; }
        public int? Type { get; set; }
        public string Xml { get; set; }
        public bool? IsManaged { get; set; }
        public int? ActivationState { get; set; }
        private FormDefinition definition;
        public FormDefinition Definition { get { return definition ?? (definition = FormXmlReader.Parse(Xml)); } }

        public string StateText
        {
            get
            {
                var parts = new List<string>();
                if (ActivationState.HasValue) parts.Add(ActivationState.Value == 1 ? "Active" : "Inactive");
                if (IsManaged.HasValue) parts.Add(IsManaged.Value ? "Managed" : "Unmanaged");
                return parts.Count == 0 ? null : string.Join(", ", parts);
            }
        }
    }

    /// <summary>Retrieves the forms of one table once per scan and shares them between the form and JavaScript scanners.</summary>
    public class FormRepository
    {
        // 2 Main, 5 Mobile-Express, 6 Quick View, 7 Quick Create, 11 Card, 12 Main-Interactive. Dashboards (0) have no table.
        private static readonly object[] FormTypes = { 2, 5, 6, 7, 11, 12 };

        private readonly IOrganizationService svc;
        private readonly Dictionary<string, List<FormRecord>> byTable = new Dictionary<string, List<FormRecord>>(StringComparer.OrdinalIgnoreCase);
        public List<string> Warnings { get; } = new List<string>();

        public FormRepository(IOrganizationService svc) { this.svc = svc; }

        public List<FormRecord> GetForms(string table, CancellationToken ct)
        {
            List<FormRecord> cached;
            if (byTable.TryGetValue(table, out cached)) return cached;

            Func<bool, QueryExpression> build = useState =>
            {
                var q = new QueryExpression("systemform")
                {
                    ColumnSet = new ColumnSet("formid", "name", "objecttypecode", "type", "formxml", "ismanaged", "formactivationstate")
                };
                q.Criteria.AddCondition("objecttypecode", ConditionOperator.Equal, table);
                q.Criteria.AddCondition("type", ConditionOperator.In, FormTypes);
                if (useState) q.Criteria.AddCondition("componentstate", ConditionOperator.Equal, 0);
                q.AddOrder("name", OrderType.Ascending);
                return q;
            };

            List<Entity> rows;
            try { rows = DataverseHelper.RetrieveAll(svc, build(true), ct, 50).ToList(); }
            catch (FaultException<OrganizationServiceFault>)
            {
                Warnings.Add("The form query with the published-state filter failed; it was retried without it (unpublished duplicates are possible).");
                rows = DataverseHelper.RetrieveAll(svc, build(false), ct, 50).ToList();
            }

            var list = rows.Select(e => new FormRecord
            {
                Id = e.Id,
                Name = DataverseHelper.Str(e, "name"),
                Table = table,
                Type = DataverseHelper.Opt(e, "type"),
                Xml = DataverseHelper.Str(e, "formxml"),
                IsManaged = DataverseHelper.Bool(e, "ismanaged"),
                ActivationState = DataverseHelper.Opt(e, "formactivationstate")
            }).ToList();
            byTable[table] = list;
            return list;
        }
    }
}
