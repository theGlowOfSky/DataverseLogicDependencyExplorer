// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Linq;

namespace DataverseLogicDependencyExplorer.Models
{
    /// <summary>How much of the JavaScript web resource library is downloaded and analysed.</summary>
    public enum JavaScriptScope
    {
        /// <summary>Every script web resource in the environment (most complete, largest download).</summary>
        AllScriptWebResources,
        /// <summary>Only the libraries registered on forms of the selected table (targeted, much faster).</summary>
        RegisteredOnSelectedTableForms
    }

    /// <summary>Everything a scan needs to know. Built by the UI and consumed by the analysis service.</summary>
    public class ScanRequest
    {
        public SearchMode Mode { get; set; }
        public string TableLogicalName { get; set; }
        public string TableDisplayName { get; set; }
        /// <summary>Plural entity set name used in Web API URLs (e.g. "accounts"); optional.</summary>
        public string TableEntitySetName { get; set; }
        public string ColumnLogicalName { get; set; }
        public string ColumnDisplayName { get; set; }

        public HashSet<ScannerKind> Scopes { get; set; } = new HashSet<ScannerKind>();

        public JavaScriptScope JavaScriptScope { get; set; } = JavaScriptScope.AllScriptWebResources;
        /// <summary>Skip managed (solution-installed) script web resources; faster, and managed code is usually not yours to change.</summary>
        public bool SkipManagedWebResources { get; set; }
        /// <summary>Include matches that only occur inside JavaScript comments (off by default; they are counted either way).</summary>
        public bool IncludeCommentMatches { get; set; }
        /// <summary>Also scan classic process definitions of other tables for cross-table references (downloads every process definition).</summary>
        public bool DeepProcessScan { get; set; }

        public string EnvironmentName { get; set; }
        public string EnvironmentUrl { get; set; }
        /// <summary>The environment identifier supplied by the connection, when available.</summary>
        public string EnvironmentId { get; set; }

        /// <summary>The value being searched for in definitions: the column in column mode, otherwise the table.</summary>
        public string SearchValue
        {
            get { return Mode == SearchMode.Column ? ColumnLogicalName : TableLogicalName; }
        }

        public static IList<ScannerKind> AllScopes
        {
            get { return Enum.GetValues(typeof(ScannerKind)).Cast<ScannerKind>().ToList(); }
        }

        /// <summary>
        /// Validates the combination of inputs. Returns null when valid, otherwise an actionable message for the user.
        /// </summary>
        public static string Validate(SearchMode mode, string table, string column, int scopeCount)
        {
            bool hasTable = !string.IsNullOrWhiteSpace(table);
            bool hasColumn = !string.IsNullOrWhiteSpace(column);

            if (mode == SearchMode.Column)
            {
                if (hasColumn && !hasTable)
                    return "A column was chosen without its parent table. Select the parent table first, then pick the column.";
                if (!hasTable)
                    return "Column mode needs a table and a column. Select a table, then select one of its columns.";
                if (!hasColumn)
                    return "Column mode needs a column. Select a column of '" + table + "' or switch the search mode to 'Table'.";
            }
            else
            {
                if (!hasTable)
                    return "Select a table to analyse.";
            }

            if (scopeCount <= 0)
                return "Select at least one component scope to scan.";

            return null;
        }

        public string Validate()
        {
            return Validate(Mode, TableLogicalName, ColumnLogicalName, Scopes == null ? 0 : Scopes.Count);
        }
    }

    public class TableInfo
    {
        public string LogicalName { get; set; }
        public string DisplayName { get; set; }
        public bool IsCustom { get; set; }
        public bool IsIntersect { get; set; }
        public string EntitySetName { get; set; }

        public string Title
        {
            get { return string.IsNullOrEmpty(DisplayName) || DisplayName == LogicalName ? LogicalName : DisplayName + " (" + LogicalName + ")"; }
        }
    }

    public class ColumnInfo
    {
        public string LogicalName { get; set; }
        public string DisplayName { get; set; }
        public string AttributeType { get; set; }
        public bool IsCustom { get; set; }
        public string Title
        {
            get { return string.IsNullOrEmpty(DisplayName) || DisplayName == LogicalName ? LogicalName : DisplayName + " (" + LogicalName + ")"; }
        }
    }
}
