// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System.ComponentModel.Composition;
using DataverseLogicDependencyExplorer.Resources;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;

namespace DataverseLogicDependencyExplorer
{
    /// <summary>
    /// MEF entry point discovered by XrmToolBox (the assembly must be in the Plugins folder). Registration follows the pattern of
    /// the working reference plugins: [Export(typeof(IXrmToolBoxPlugin))] on a class deriving from PluginBase that returns the
    /// user control. The plugin has no third-party runtime dependencies, so no assembly resolver is needed.
    /// </summary>
    [Export(typeof(IXrmToolBoxPlugin)),
     ExportMetadata("Name", "Dataverse Logic Dependency Explorer"),
     ExportMetadata("Description", "Finds where a Dataverse table or column is referenced by business rules, JavaScript, forms, plugin registrations, workflows and cloud flows, and exports an impact report to Excel. Maintainer: Gagan Deep Singh (linkedin.com/in/glowofsky)."),
     ExportMetadata("SmallImageBase64", IconData.Small32),
     ExportMetadata("BigImageBase64", IconData.Big80),
     ExportMetadata("BackgroundColor", "White"),
     ExportMetadata("PrimaryFontColor", "#0F6CBD"),
     ExportMetadata("SecondaryFontColor", "DimGray")]
    public class Plugin : PluginBase
    {
        public override IXrmToolBoxPluginControl GetControl()
        {
            return new MainControl();
        }
    }
}
