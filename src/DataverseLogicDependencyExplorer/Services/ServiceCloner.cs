// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Reflection;
using Microsoft.Xrm.Sdk;

namespace DataverseLogicDependencyExplorer.Services
{
    /// <summary>
    /// Creates extra connections for parallel downloads. Dataverse clients (ServiceClient / CrmServiceClient) expose a public
    /// parameterless Clone() that reuses the existing authentication without prompting; it is called by reflection so the plugin
    /// does not need a compile-time dependency on a specific client type. When cloning is not available the original connection
    /// is used by a single download worker (analysis still runs in parallel with it).
    /// </summary>
    public static class ServiceCloner
    {
        public static IList<IOrganizationService> CreateWorkers(IOrganizationService service, int count)
        {
            var workers = new List<IOrganizationService>();
            if (service != null && count > 1)
            {
                MethodInfo clone = service.GetType().GetMethod("Clone", Type.EmptyTypes);
                if (clone != null && typeof(IOrganizationService).IsAssignableFrom(clone.ReturnType))
                {
                    for (int i = 0; i < count; i++)
                    {
                        try
                        {
                            var c = clone.Invoke(service, null) as IOrganizationService;
                            if (c == null) break;
                            workers.Add(c);
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Trace.TraceWarning("Connection clone failed (" + ex.Message + "); using fewer parallel downloads.");
                            break;
                        }
                    }
                }
            }
            if (workers.Count == 0) workers.Add(service);
            return workers;
        }

        /// <summary>Disposes the cloned connections (never the original).</summary>
        public static void Release(IOrganizationService original, IList<IOrganizationService> workers)
        {
            if (workers == null) return;
            foreach (IOrganizationService w in workers)
            {
                if (ReferenceEquals(w, original)) continue;
                var d = w as IDisposable;
                if (d != null) { try { d.Dispose(); } catch (Exception) { } }
            }
        }
    }
}
