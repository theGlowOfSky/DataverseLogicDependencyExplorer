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
using Microsoft.Xrm.Sdk.Query;
using DataverseLogicDependencyExplorer.Models;

namespace DataverseLogicDependencyExplorer.Services
{
    /// <summary>
    /// Finds the (non-default) solutions a component belongs to. Each solution component type is loaded once, on demand,
    /// and only when a finding of that type exists.
    /// </summary>
    public class SolutionLookup
    {
        public const int WorkflowComponent = 29, WebResourceComponent = 61, FormComponent = 60, PluginStepComponent = 92;

        private static readonly object[] IgnoredSolutions = { "Default", "Active", "Basic" };

        private readonly IOrganizationService svc;
        private readonly Dictionary<int, Dictionary<Guid, List<string>>> byType = new Dictionary<int, Dictionary<Guid, List<string>>>();
        private readonly HashSet<int> failedTypes = new HashSet<int>();
        public List<string> Warnings { get; } = new List<string>();

        public SolutionLookup(IOrganizationService svc) { this.svc = svc; }

        public static int? ComponentTypeCode(ComponentType t)
        {
            switch (t)
            {
                case ComponentType.BusinessRule:
                case ComponentType.ClassicWorkflow:
                case ComponentType.CloudFlow: return WorkflowComponent;
                case ComponentType.JavaScriptWebResource: return WebResourceComponent;
                case ComponentType.Form:
                case ComponentType.FormEventRegistration: return FormComponent;
                case ComponentType.PluginStep: return PluginStepComponent;
                default: return null;
            }
        }

        public string Find(int componentType, Guid id, CancellationToken ct)
        {
            Dictionary<Guid, List<string>> map = Load(componentType, ct);
            List<string> names;
            return map != null && map.TryGetValue(id, out names) ? string.Join(", ", names.Distinct().OrderBy(n => n)) : null;
        }

        private Dictionary<Guid, List<string>> Load(int type, CancellationToken ct)
        {
            Dictionary<Guid, List<string>> cached;
            if (byType.TryGetValue(type, out cached)) return cached;
            if (failedTypes.Contains(type)) return null;

            var map = new Dictionary<Guid, List<string>>();
            try
            {
                var q = new QueryExpression("solutioncomponent") { ColumnSet = new ColumnSet("objectid", "componenttype") };
                q.Criteria.AddCondition("componenttype", ConditionOperator.Equal, type);
                LinkEntity link = q.AddLink("solution", "solutionid", "solutionid", JoinOperator.Inner);
                link.EntityAlias = "sol";
                link.Columns = new ColumnSet("friendlyname", "uniquename");
                link.LinkCriteria.AddCondition("uniquename", ConditionOperator.NotIn, IgnoredSolutions);

                foreach (Entity e in DataverseHelper.RetrieveAll(svc, q, ct, 500))
                {
                    Guid? objectId = null;
                    if (e.Contains("objectid")) objectId = e["objectid"] as Guid?;
                    if (!objectId.HasValue) continue;
                    string name = DataverseHelper.Str(e, "sol.friendlyname") ?? DataverseHelper.Str(e, "sol.uniquename");
                    if (string.IsNullOrEmpty(name)) continue;
                    List<string> list;
                    if (!map.TryGetValue(objectId.Value, out list)) map[objectId.Value] = list = new List<string>();
                    list.Add(name);
                }
                byType[type] = map;
                return map;
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                failedTypes.Add(type);
                Warnings.Add("Solution information for component type " + type + " could not be read (" + DataverseHelper.Describe(ex) + "). The Solution column is blank for those findings.");
                return null;
            }
        }

        /// <summary>Fills in SolutionName for findings that do not already have one.</summary>
        public void Enrich(IEnumerable<DependencyResult> results, CancellationToken ct)
        {
            foreach (DependencyResult r in results)
            {
                if (!string.IsNullOrEmpty(r.SolutionName)) continue;
                int? code = ComponentTypeCode(r.ComponentType);
                Guid id;
                if (!code.HasValue || !Guid.TryParse(r.ComponentId, out id)) continue;
                r.SolutionName = Find(code.Value, id, ct);
            }
        }
    }
}
