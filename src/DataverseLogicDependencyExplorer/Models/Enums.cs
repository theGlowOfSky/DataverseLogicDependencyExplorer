// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;

namespace DataverseLogicDependencyExplorer.Models
{
    /// <summary>What the user wants analysed.</summary>
    public enum SearchMode
    {
        /// <summary>The selected table and its associated components.</summary>
        Table,
        /// <summary>The selected column and components that reference it.</summary>
        Column
    }

    /// <summary>The component categories a user can include in a scan (one scanner per category).</summary>
    public enum ScannerKind
    {
        BusinessRules,
        JavaScript,
        Forms,
        Plugins,
        ClassicWorkflows,
        CloudFlows
    }

    /// <summary>The kind of component a single finding belongs to.</summary>
    public enum ComponentType
    {
        BusinessRule,
        JavaScriptWebResource,
        Form,
        FormEventRegistration,
        PluginStep,
        ClassicWorkflow,
        CloudFlow
    }

    /// <summary>
    /// How strong the evidence for a finding is. The classification is chosen from the evidence a scanner actually
    /// inspected; a name appearing in some text is never labelled "confirmed".
    /// </summary>
    public enum ConfidenceLevel
    {
        /// <summary>Dataverse metadata itself identifies the association (registration, filtering attribute, primary table).</summary>
        ConfirmedMetadataReference = 0,
        /// <summary>A relevant table or column is identified inside a parsed component definition (XML/JSON structure).</summary>
        StrongDefinitionReference = 1,
        /// <summary>A source-text match was found, but a true runtime dependency is not established.</summary>
        TextMatchReviewRequired = 2,
        /// <summary>A pattern indicates a possible reference that cannot be resolved statically.</summary>
        PossibleOrDynamic = 3
    }

    /// <summary>Outcome of one component-category scan. "NotInspected" must never be read as "no dependencies".</summary>
    public enum ScanStatus
    {
        NotInspected = 0,
        Completed = 1,
        CompletedWithWarnings = 2,
        PartiallySupported = 3,
        PermissionDenied = 4,
        Failed = 5,
        Cancelled = 6
    }

    public static class EnumText
    {
        public static string Display(this ConfidenceLevel c)
        {
            switch (c)
            {
                case ConfidenceLevel.ConfirmedMetadataReference: return "Confirmed metadata reference";
                case ConfidenceLevel.StrongDefinitionReference: return "Strong definition reference";
                case ConfidenceLevel.TextMatchReviewRequired: return "Text match - review required";
                case ConfidenceLevel.PossibleOrDynamic: return "Possible or dynamic dependency";
                default: return c.ToString();
            }
        }

        public static string Explain(this ConfidenceLevel c)
        {
            switch (c)
            {
                case ConfidenceLevel.ConfirmedMetadataReference:
                    return "The actual Dataverse metadata identifies the association, registration, filtering attribute or relationship.";
                case ConfidenceLevel.StrongDefinitionReference:
                    return "A relevant table or column was identified inside a parsed component definition (XML or JSON structure).";
                case ConfidenceLevel.TextMatchReviewRequired:
                    return "A source-text match was found, but a real runtime dependency is not established. Review it manually.";
                case ConfidenceLevel.PossibleOrDynamic:
                    return "A pattern suggests a possible reference that cannot be resolved by static analysis.";
                default: return string.Empty;
            }
        }

        public static string Display(this ScanStatus s)
        {
            switch (s)
            {
                case ScanStatus.NotInspected: return "Not inspected";
                case ScanStatus.Completed: return "Completed";
                case ScanStatus.CompletedWithWarnings: return "Completed with warnings";
                case ScanStatus.PartiallySupported: return "Partially supported";
                case ScanStatus.PermissionDenied: return "Permission denied";
                case ScanStatus.Failed: return "Failed";
                case ScanStatus.Cancelled: return "Cancelled";
                default: return s.ToString();
            }
        }

        public static string Display(this ComponentType t)
        {
            switch (t)
            {
                case ComponentType.BusinessRule: return "Business Rule";
                case ComponentType.JavaScriptWebResource: return "JavaScript Web Resource";
                case ComponentType.Form: return "Form";
                case ComponentType.FormEventRegistration: return "Form Event Registration";
                case ComponentType.PluginStep: return "Plugin Step";
                case ComponentType.ClassicWorkflow: return "Classic Workflow / Process";
                case ComponentType.CloudFlow: return "Cloud Flow";
                default: return t.ToString();
            }
        }

        public static string Display(this ScannerKind k)
        {
            switch (k)
            {
                case ScannerKind.BusinessRules: return "Business Rules";
                case ScannerKind.JavaScript: return "JavaScript Web Resources";
                case ScannerKind.Forms: return "Forms and Form Event Registrations";
                case ScannerKind.Plugins: return "Plugin Registrations";
                case ScannerKind.ClassicWorkflows: return "Classic Workflows and Processes";
                case ScannerKind.CloudFlows: return "Power Automate Cloud Flows";
                default: return k.ToString();
            }
        }

        public static string Display(this SearchMode m)
        {
            return m == SearchMode.Table ? "Table (and its associated components)" : "Column (and components that reference it)";
        }
    }
}
