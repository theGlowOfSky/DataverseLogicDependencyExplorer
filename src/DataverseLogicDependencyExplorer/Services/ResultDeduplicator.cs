// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Linq;
using DataverseLogicDependencyExplorer.Models;

namespace DataverseLogicDependencyExplorer.Services
{
    /// <summary>Removes findings that the same component produced through overlapping analysis paths.</summary>
    public static class ResultDeduplicator
    {
        /// <summary>Keeps one finding per duplicate key, preferring the strongest confidence (then the first found).</summary>
        public static List<DependencyResult> Deduplicate(IEnumerable<DependencyResult> results)
        {
            var best = new Dictionary<string, DependencyResult>();
            var order = new List<string>();
            foreach (DependencyResult r in results)
            {
                string key = r.DuplicateKey();
                DependencyResult existing;
                if (!best.TryGetValue(key, out existing)) { best[key] = r; order.Add(key); }
                else if (r.Confidence < existing.Confidence) best[key] = r;
            }
            return order.Select(k => best[k]).ToList();
        }

        public static List<DependencyResult> Sort(IEnumerable<DependencyResult> results)
        {
            return results.OrderBy(r => r.ComponentType)
                          .ThenBy(r => r.Confidence)
                          .ThenBy(r => r.ComponentName, StringComparer.OrdinalIgnoreCase)
                          .ThenBy(r => r.LineNumber ?? 0)
                          .ToList();
        }
    }
}
