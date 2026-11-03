// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
namespace DataverseLogicDependencyExplorer.Services
{
    /// <summary>Texts of the confirmation shown before a potentially long scan (kept out of the UI so they can be tested).</summary>
    public static class ScanWarnings
    {
        /// <summary>
        /// The one warning of the tool: shown only when the full "all script web resources" search is about to start.
        /// When managed web resources are already being skipped, skipping them is not suggested again.
        /// </summary>
        public static string AllWebResources(bool skipManagedAlreadyTicked)
        {
            string hint = skipManagedAlreadyTicked
                ? "To finish faster you can cancel and choose 'Only libraries registered on the selected table's forms'."
                : "To finish faster you can cancel and either tick 'Skip managed web resources' or choose 'Only libraries registered on the selected table's forms'.";

            return "You selected 'All script web resources'.\n\n" +
                   "This searches EVERY JavaScript web resource in the environment, so it can take several minutes in a large environment. " +
                   "Progress is shown while it runs, and findings appear in the Impact Report tab as soon as they are found.\n\n" +
                   hint + "\n\n" +
                   "Continue with the full search?";
        }
    }
}
