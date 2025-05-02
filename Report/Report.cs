using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Report.DTO.VisualDto;
using TabularEditor.TOMWrapper;
using TabularEditor.Scripting;

namespace Report.DTO
{
    public class PagesDto
    {
        [Newtonsoft.Json.JsonProperty("$schema")]
        public string Schema { get; set; }

        [Newtonsoft.Json.JsonProperty("pageOrder")]
        public List<string> PageOrder { get; set; }

        [Newtonsoft.Json.JsonProperty("activePageName")]
        public string ActivePageName { get; set; }
    }

    public class PageDto
    {
        [Newtonsoft.Json.JsonProperty("$schema")]
        public string Schema { get; set; }

        [Newtonsoft.Json.JsonProperty("name")]
        public string Name { get; set; }

        [Newtonsoft.Json.JsonProperty("displayName")]
        public string DisplayName { get; set; }

        [Newtonsoft.Json.JsonProperty("displayOption")]
        public string DisplayOption { get; set; } // Could create enum if you want stricter typing

        [Newtonsoft.Json.JsonProperty("height")]
        public double? Height { get; set; }

        [Newtonsoft.Json.JsonProperty("width")]
        public double? Width { get; set; }
    }


    public class VisualDto
    {
        public class Root
        {
            [JsonProperty("$schema")]
            public string Schema { get; set; }

            [JsonProperty("name")]
            public string Name { get; set; }

            [JsonProperty("position")]
            public Position Position { get; set; }

            [JsonProperty("visual")]
            public Visual Visual { get; set; }
        }

        public class Position
        {
            [JsonProperty("x")]
            public double X { get; set; }

            [JsonProperty("y")]
            public double Y { get; set; }

            [JsonProperty("z")]
            public int Z { get; set; }

            [JsonProperty("height")]
            public double Height { get; set; }

            [JsonProperty("width")]
            public double Width { get; set; }

            [JsonProperty("tabOrder")]
            public int TabOrder { get; set; }
        }

        public class Visual
        {
            [JsonProperty("visualType")]
            public string VisualType { get; set; }

            [JsonProperty("query")]
            public Query Query { get; set; }

            [JsonProperty("objects")]
            public Objects Objects { get; set; }

            [JsonProperty("drillFilterOtherVisuals")]
            public bool DrillFilterOtherVisuals { get; set; }
        }

        public class Query
        {
            [JsonProperty("queryState")]
            public QueryState QueryState { get; set; }

            [JsonProperty("sortDefinition")]
            public SortDefinition SortDefinition { get; set; }
        }

        public class QueryState
        {
            [JsonProperty("Y")]
            public ProjectionsSet Y { get; set; }

            [JsonProperty("Values")]
            public ProjectionsSet Values { get; set; }

            [JsonProperty("Category")]
            public ProjectionsSet Category { get; set; }

            [JsonProperty("Series")]
            public ProjectionsSet Series { get; set; }
        }

        public class ProjectionsSet
        {
            [JsonProperty("projections")]
            public List<Projection> Projections { get; set; }
        }

        public class Projection
        {
            [JsonProperty("field")]
            public Field Field { get; set; }

            [JsonProperty("queryRef")]
            public string QueryRef { get; set; }

            [JsonProperty("nativeQueryRef")]
            public string NativeQueryRef { get; set; }

            [JsonProperty("active")]
            public bool? Active { get; set; }

            [JsonProperty("hidden")]
            public bool? Hidden { get; set; }
        }

        public class Field
        {
            [JsonProperty("Aggregation")]
            public Aggregation Aggregation { get; set; }

            [JsonProperty("NativeVisualCalculation")]
            public NativeVisualCalculation NativeVisualCalculation { get; set; }

            [JsonProperty("Measure")]
            public MeasureObject Measure { get; set; }

            [JsonProperty("Column")]
            public ColumnField Column { get; set; }
        }

        public class Aggregation
        {
            [JsonProperty("Expression")]
            public Expression Expression { get; set; }

            [JsonProperty("Function")]
            public int Function { get; set; }
        }

        public class NativeVisualCalculation
        {
            [JsonProperty("Language")]
            public string Language { get; set; }

            [JsonProperty("Expression")]
            public string Expression { get; set; }

            [JsonProperty("Name")]
            public string Name { get; set; }
        }

        public class MeasureObject
        {
            [JsonProperty("Expression")]
            public Expression Expression { get; set; }

            [JsonProperty("Property")]
            public string Property { get; set; }
        }

        public class ColumnField
        {
            [JsonProperty("Expression")]
            public Expression Expression { get; set; }

            [JsonProperty("Property")]
            public string Property { get; set; }
        }

        public class Expression
        {
            [JsonProperty("Column")]
            public ColumnExpression Column { get; set; }

            [JsonProperty("SourceRef")]
            public SourceRef SourceRef { get; set; }
        }

        public class ColumnExpression
        {
            [JsonProperty("Expression")]
            public SourceRef Expression { get; set; }

            [JsonProperty("Property")]
            public string Property { get; set; }
        }

        public class SourceRef
        {
            [JsonProperty("Entity")]
            public string Entity { get; set; }

            [JsonProperty("Source")]
            public string Source { get; set; }
        }

        public class SortDefinition
        {
            [JsonProperty("sort")]
            public List<Sort> Sort { get; set; }

            [JsonProperty("isDefaultSort")]
            public bool IsDefaultSort { get; set; }
        }

        public class Sort
        {
            [JsonProperty("field")]
            public Field Field { get; set; }

            [JsonProperty("direction")]
            public string Direction { get; set; }
        }

        public class Objects
        {
            [JsonProperty("valueAxis")]
            public List<ObjectProperties> ValueAxis { get; set; }

            [JsonProperty("general")]
            public List<ObjectProperties> General { get; set; }

            [JsonProperty("data")]
            public List<ObjectProperties> Data { get; set; }

            [JsonProperty("title")]
            public List<ObjectProperties> Title { get; set; }

            [JsonProperty("legend")]
            public List<ObjectProperties> Legend { get; set; }

            [JsonProperty("labels")]
            public List<ObjectProperties> Labels { get; set; }
        }

        public class ObjectProperties
        {
            [JsonProperty("properties")]
            public Dictionary<string, object> Properties { get; set; }
        }
    }
    public class VisualExtended
    {
        public VisualDto.Root Visual { get; set; }
        public string VisualFilePath { get; set; }

        public VisualDto.Root Content => Visual;

        // Get Projections from a selected container: Values, Y, Series, etc.
        public IEnumerable<VisualDto.Projection> GetProjections(Func<VisualDto.QueryState, VisualDto.ProjectionsSet> selector)
        {
            return selector?.Invoke(Visual?.Visual?.Query?.QueryState)?.Projections
                   ?? Enumerable.Empty<VisualDto.Projection>();
        }

        // Check if this visual contains a given field key like 'Table'[Field]
        public bool ContainsField(string fieldKey, Func<VisualDto.QueryState, VisualDto.ProjectionsSet> selector)
        {
            return GetProjections(selector).Any(p => p.ToFieldKey() == fieldKey);
        }

        public IEnumerable<string> GetAllReferencedMeasures()
        {
            return GetAllProjections()
                .Select(p => p.Field?.Measure)
                .Where(m => m?.Expression?.SourceRef?.Entity != null && m.Property != null)
                .Select(m => $"'{m.Expression.SourceRef.Entity}'[{m.Property}]")
                .Distinct();
        }

        public IEnumerable<string> GetAllReferencedColumns()
        {
            return GetAllProjections()
                .Select(p => p.Field?.Column)
                .Where(c => c?.Expression?.SourceRef?.Entity != null && c.Property != null)
                .Select(c => $"'{c.Expression.SourceRef.Entity}'[{c.Property}]")
                .Distinct();
        }

        // Helper to gather all projections from all known containers
        private IEnumerable<VisualDto.Projection> GetAllProjections()
        {
            var queryState = Visual?.Visual?.Query?.QueryState;
            if (queryState == null) return Enumerable.Empty<VisualDto.Projection>();

            return new[]
            {
                queryState.Values?.Projections,
                queryState.Y?.Projections,
                queryState.Series?.Projections,
                queryState.Category?.Projections
            }
            .Where(list => list != null)
            .SelectMany(list => list);
        }

        public void ReplaceField(string originalFieldKey, Measure newMeasure)
        {
            foreach (var proj in GetAllProjections())
            {
                if (proj.ToFieldKey() == originalFieldKey)
                {
                    proj.ReplaceField(newMeasure);
                }
            }
        }

        public void ReplaceField(string originalFieldKey, Column newColumn)
        {
            foreach (var proj in GetAllProjections())
            {
                if (proj.ToFieldKey() == originalFieldKey)
                {
                    proj.ReplaceField(newColumn);
                }
            }
        }

    }

    public static class ProjectionExtensions
    {
        public static string ToFieldKey(this VisualDto.Projection pr)
        {
            var measure = pr.Field?.Measure;
            var column = pr.Field?.Column;

            if (measure?.Expression?.SourceRef?.Entity != null && measure.Property != null)
            {
                return $"'{measure.Expression.SourceRef.Entity}'[{measure.Property}]";
            }
            else if (column?.Expression?.SourceRef?.Entity != null && column.Property != null)
            {
                return $"'{column.Expression.SourceRef.Entity}'[{column.Property}]";
            }

            return null;
        }

        public static bool MatchesFieldKey(this VisualDto.Projection pr, string fieldKey)
        {
            return pr.ToFieldKey() == fieldKey;
        }

        public static void ReplaceField(this VisualDto.Projection projection, Measure newMeasure)
        {
            if (projection?.Field == null || newMeasure == null) return;

            projection.Field.Measure = new VisualDto.MeasureObject
            {
                Property = newMeasure.Name,
                Expression = new VisualDto.Expression
                {
                    SourceRef = new VisualDto.SourceRef
                    {
                        Entity = newMeasure.Table.Name
                    }
                }
            };
            projection.QueryRef = newMeasure.Table.Name + "." + newMeasure.Name;
            projection.NativeQueryRef = newMeasure.Name;

            // Clear column to avoid ambiguity
            projection.Field.Column = null;
        }

        public static void ReplaceField(this VisualDto.Projection projection, Column newColumn)
        {
            if (projection?.Field == null || newColumn == null) return;

            projection.Field.Column = new VisualDto.ColumnField
            {
                Property = newColumn.Name,
                Expression = new VisualDto.Expression
                {
                    SourceRef = new VisualDto.SourceRef
                    {
                        Entity = newColumn.Table.Name
                    }
                }
            };
            projection.QueryRef = newColumn.Table.Name + "." + newColumn.Name;
            projection.NativeQueryRef = newColumn.Name;

            // Clear measure to avoid ambiguity
            projection.Field.Measure = null;
        }

    }

    public class PageExtended
    {
        public PageDto Page { get; set; }
        public IList<VisualExtended> Visuals { get; set; }
        public string PageFilePath { get; set; }

        public PageExtended()
        {
            Visuals = new List<VisualExtended>();
        }
    }

    public class ReportExtended
    {
        public IList<PageExtended> Pages { get; set; }
        public string PagesFilePath { get; set; }

        public PagesDto PagesConfig { get; set; }

        public ReportExtended()
        {
            Pages = new List<PageExtended>();



        }
    }

}
