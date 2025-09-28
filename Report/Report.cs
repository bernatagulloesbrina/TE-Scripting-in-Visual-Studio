using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TabularEditor.TOMWrapper;
using TabularEditor.Scripting;
using Newtonsoft.Json.Linq;
using static Report.DTO.VisualDto;

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


    public partial class VisualDto
    {
        public class Root
        {
            [JsonProperty("$schema")] public string Schema { get; set; }
            [JsonProperty("name")] public string Name { get; set; }
            [JsonProperty("position")] public Position Position { get; set; }
            [JsonProperty("visual")] public Visual Visual { get; set; }
            

            [JsonProperty("visualGroup")] public VisualGroup VisualGroup { get; set; }
            [JsonProperty("parentGroupName")] public string ParentGroupName { get; set; }
            [JsonProperty("filterConfig")] public FilterConfig FilterConfig { get; set; }
            [JsonProperty("isHidden")] public bool IsHidden { get; set; }

            [JsonExtensionData]
            
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }


        public class VisualContainerObjects
        {
            [JsonProperty("general")]
            public List<VisualContainerObject> General { get; set; }

            // Add other known properties as needed, e.g.:
            [JsonProperty("title")]
            public List<VisualContainerObject> Title { get; set; }

            [JsonProperty("subTitle")]
            public List<VisualContainerObject> SubTitle { get; set; }

            // This will capture any additional properties not explicitly defined above
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class VisualContainerObject
        {
            [JsonProperty("properties")]
            public Dictionary<string, VisualContainerProperty> Properties { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class VisualContainerProperty
        {
            [JsonProperty("expr")]
            public VisualExpr Expr { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class VisualExpr
        {
            [JsonProperty("Literal")]
            public VisualLiteral Literal { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class VisualLiteral
        {
            [JsonProperty("Value")]
            public string Value { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class VisualGroup
        {
            [JsonProperty("displayName")] public string DisplayName { get; set; }
            [JsonProperty("groupMode")] public string GroupMode { get; set; }
        }

        public class Position
        {
            [JsonProperty("x")] public double X { get; set; }
            [JsonProperty("y")] public double Y { get; set; }
            [JsonProperty("z")] public int Z { get; set; }
            [JsonProperty("height")] public double Height { get; set; }
            [JsonProperty("width")] public double Width { get; set; }

            [JsonProperty("tabOrder", NullValueHandling = NullValueHandling.Ignore)]
            public int? TabOrder { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class Visual
        {
            [JsonProperty("visualType")] public string VisualType { get; set; }
            [JsonProperty("query")] public Query Query { get; set; }
            [JsonProperty("objects")] public Objects Objects { get; set; }
            [JsonProperty("visualContainerObjects")]
            public VisualContainerObjects VisualContainerObjects { get; set; }
            [JsonProperty("drillFilterOtherVisuals")] public bool DrillFilterOtherVisuals { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class Query
        {
            [JsonProperty("queryState")] public QueryState QueryState { get; set; }
            [JsonProperty("sortDefinition")] public SortDefinition SortDefinition { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class  QueryState
        {
            [JsonProperty("Rows", Order = 1)] public VisualDto.ProjectionsSet Rows { get; set; }
            [JsonProperty("Category", Order = 2)] public VisualDto.ProjectionsSet Category { get; set; }
            [JsonProperty("Y", Order = 3)] public VisualDto.ProjectionsSet Y { get; set; }
            [JsonProperty("Y2", Order = 4)] public VisualDto.ProjectionsSet Y2 { get; set; }
            [JsonProperty("Values", Order = 5)] public VisualDto.ProjectionsSet Values { get; set; }
            
            [JsonProperty("Series", Order = 6)] public VisualDto.ProjectionsSet Series { get; set; }
            [JsonProperty("Data", Order = 7)] public VisualDto.ProjectionsSet Data { get; set; }

            
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class ProjectionsSet
        {
            [JsonProperty("projections")] public List<VisualDto.Projection> Projections { get; set; }
            [JsonProperty("fieldParameters")] public List<VisualDto.FieldParameter> FieldParameters { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class FieldParameter
        {
            [JsonProperty("parameterExpr")]
            public Field ParameterExpr { get; set; }

            [JsonProperty("index")]
            public int Index { get; set; }

            [JsonProperty("length")]
            public int Length { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class Projection
        {
            [JsonProperty("field")] public VisualDto.Field Field { get; set; }
            [JsonProperty("queryRef")] public string QueryRef { get; set; }
            [JsonProperty("nativeQueryRef")] public string NativeQueryRef { get; set; }
            [JsonProperty("active")] public bool? Active { get; set; }
            [JsonProperty("hidden")] public bool? Hidden { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class Field
        {
            [JsonProperty("Aggregation")] public VisualDto.Aggregation Aggregation { get; set; }
            [JsonProperty("NativeVisualCalculation")] public NativeVisualCalculation NativeVisualCalculation { get; set; }
            [JsonProperty("Measure")] public VisualDto.MeasureObject Measure { get; set; }
            [JsonProperty("Column")] public VisualDto.ColumnField Column { get; set; }

            [JsonExtensionData] public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class Aggregation
        {
            [JsonProperty("Expression")] public VisualDto.Expression Expression { get; set; }
            [JsonProperty("Function")] public int Function { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class NativeVisualCalculation
        {
            [JsonProperty("Language")] public string Language { get; set; }
            [JsonProperty("Expression")] public string Expression { get; set; }
            [JsonProperty("Name")] public string Name { get; set; }

            [JsonProperty("DataType")] public string DataType { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class MeasureObject
        {
            [JsonProperty("Expression")] public VisualDto.Expression Expression { get; set; }
            [JsonProperty("Property")] public string Property { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class ColumnField
        {
            [JsonProperty("Expression")] public VisualDto.Expression Expression { get; set; }
            [JsonProperty("Property")] public string Property { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class Expression
        {
            [JsonProperty("Column")] public ColumnExpression Column { get; set; }
            [JsonProperty("SourceRef")] public VisualDto.SourceRef SourceRef { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class ColumnExpression
        {
            [JsonProperty("Expression")] public VisualDto.SourceRef Expression { get; set; }
            [JsonProperty("Property")] public string Property { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class SourceRef
        {
            [JsonProperty("Schema")] public string Schema { get; set; }
            [JsonProperty("Entity")] public string Entity { get; set; }
            [JsonProperty("Source")] public string Source { get; set; }

            
        }

        public class SortDefinition
        {
            [JsonProperty("sort")] public List<VisualDto.Sort> Sort { get; set; }
            [JsonProperty("isDefaultSort")] public bool IsDefaultSort { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class Sort
        {
            [JsonProperty("field")] public VisualDto.Field Field { get; set; }
            [JsonProperty("direction")] public string Direction { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class Objects
        {
            [JsonProperty("valueAxis")] public List<VisualDto.ObjectProperties> ValueAxis { get; set; }
            [JsonProperty("general")] public List<VisualDto.ObjectProperties> General { get; set; }
            [JsonProperty("data")] public List<VisualDto.ObjectProperties> Data { get; set; }
            [JsonProperty("title")] public List<VisualDto.ObjectProperties> Title { get; set; }
            [JsonProperty("legend")] public List<VisualDto.ObjectProperties> Legend { get; set; }
            [JsonProperty("labels")] public List<VisualDto.ObjectProperties> Labels { get; set; }
            [JsonProperty("dataPoint")] public List<VisualDto.ObjectProperties> DataPoint { get; set; }
            [JsonProperty("columnFormatting")]
            public List<VisualDto.ObjectProperties> ColumnFormatting { get; set; }

            [JsonProperty("referenceLabel")] public List<VisualDto.ObjectProperties> ReferenceLabel { get; set; }
            [JsonProperty("referenceLabelDetail")] public List<VisualDto.ObjectProperties> ReferenceLabelDetail { get; set; }
            [JsonProperty("referenceLabelValue")] public List<VisualDto.ObjectProperties> ReferenceLabelValue { get; set; }

            [JsonProperty("values")] public List<VisualDto.ObjectProperties> Values { get; set; }

            [JsonProperty("y1AxisReferenceLine")] public List<VisualDto.ObjectProperties> Y1AxisReferenceLine { get; set; }


            [JsonExtensionData] public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class ObjectProperties
        {
            [JsonProperty("properties")]
            [JsonConverter(typeof(PropertiesConverter))]
            public Dictionary<string, object> Properties { get; set; }

            [JsonProperty("selector")]
            public Selector Selector { get; set; }


            [JsonExtensionData] public IDictionary<string, JToken> ExtensionData { get; set; }
        }




        public class VisualObjectProperty
        {
            [JsonProperty("expr")] public Field Expr { get; set; }
            [JsonProperty("solid")] public SolidColor Solid { get; set; }
            [JsonProperty("color")] public ColorExpression Color { get; set; }

            [JsonProperty("paragraphs")]
            public List<Paragraph> Paragraphs { get; set; }

            [JsonExtensionData] public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class Paragraph
        {
            [JsonProperty("textRuns")]
            public List<TextRun> TextRuns { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class TextRun
        {
            [JsonProperty("value")]
            public string Value { get; set; }

            [JsonProperty("url")]
            public string Url { get; set; }

            [JsonProperty("textStyle")]
            public Dictionary<string, object> TextStyle { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class SolidColor
        {
            [JsonProperty("color")] public ColorExpression Color { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class ColorExpression
        {
            [JsonProperty("expr")]
            public VisualColorExprWrapper Expr { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class FillRuleExprWrapper
        {
            [JsonProperty("FillRule")] public FillRuleExpression FillRule { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class FillRuleExpression
        {
            [JsonProperty("Input")] public VisualDto.Field Input { get; set; }
            [JsonProperty("FillRule")] public Dictionary<string, object> FillRule { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class ThemeDataColor
        {
            [JsonProperty("ColorId")] public int ColorId { get; set; }
            [JsonProperty("Percent")] public double Percent { get; set; }
            [JsonExtensionData] public Dictionary<string, JToken> ExtensionData { get; set; }
        }
        public class VisualColorExprWrapper
        {
            [JsonProperty("Measure")]
            public VisualDto.MeasureObject Measure { get; set; }

            [JsonProperty("Column")]
            public VisualDto.ColumnField Column { get; set; }

            [JsonProperty("Aggregation")]
            public VisualDto.Aggregation Aggregation { get; set; }

            [JsonProperty("NativeVisualCalculation")]
            public NativeVisualCalculation NativeVisualCalculation { get; set; }

            [JsonProperty("FillRule")]
            public FillRuleExpression FillRule { get; set; }

            public VisualLiteral Literal { get; set; }

            [JsonProperty("ThemeDataColor")] 
            public ThemeDataColor ThemeDataColor { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }


        

        public class Selector
        {
            

            [JsonProperty("id")]
            public string Id { get; set; }

            [JsonProperty("order")]
            public int? Order { get; set; }

            [JsonProperty("data")]
            public List<object> Data { get; set; }

            [JsonProperty("metadata")]
            public string Metadata { get; set; }

            [JsonProperty("scopeId")]
            public string ScopeId { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }


        public class FilterConfig
        {
            [JsonProperty("filters")]
            public List<VisualFilter> Filters { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class VisualFilter
        {
            [JsonProperty("name")] public string Name { get; set; }
            [JsonProperty("field")] public VisualDto.Field Field { get; set; }
            [JsonProperty("type")] public string Type { get; set; }
            [JsonProperty("filter")] public FilterDefinition Filter { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class FilterDefinition
        {
            [JsonProperty("Version")] public int Version { get; set; }
            [JsonProperty("From")] public List<FilterFrom> From { get; set; }
            [JsonProperty("Where")] public List<FilterWhere> Where { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class FilterFrom
        {
            [JsonProperty("Name")] public string Name { get; set; }
            [JsonProperty("Entity")] public string Entity { get; set; }
            [JsonProperty("Type")] public int Type { get; set; }
            [JsonProperty("Expression")] public FilterExpression Expression { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class FilterExpression
        {
            [JsonProperty("Subquery")] public SubqueryExpression Subquery { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class SubqueryExpression
        {
            [JsonProperty("Query")] public SubqueryQuery Query { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class SubqueryQuery
        {
            [JsonProperty("Version")] public int Version { get; set; }
            [JsonProperty("From")] public List<FilterFrom> From { get; set; }
            [JsonProperty("Select")] public List<SelectExpression> Select { get; set; }
            [JsonProperty("OrderBy")] public List<OrderByExpression> OrderBy { get; set; }
            [JsonProperty("Top")] public int? Top { get; set; }

            [JsonProperty("Where")] public List<FilterWhere> Where { get; set; } // 🔹 Added

            [JsonExtensionData] public Dictionary<string, JToken> ExtensionData { get; set; }
        }


        public class SelectExpression
        {
            [JsonProperty("Column")] public ColumnSelect Column { get; set; }
            [JsonProperty("Name")] public string Name { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class ColumnSelect
        {
            [JsonProperty("Expression")]
            public VisualDto.Expression Expression { get; set; }  // NOTE: wrapper that contains "SourceRef"

            [JsonProperty("Property")]
            public string Property { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class OrderByExpression
        {
            [JsonProperty("Direction")] public int Direction { get; set; }
            [JsonProperty("Expression")] public OrderByInnerExpression Expression { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class OrderByInnerExpression
        {
            [JsonProperty("Measure")] public VisualDto.MeasureObject Measure { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class FilterWhere
        {
            [JsonProperty("Condition")] public Condition Condition { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class Condition
        {
            [JsonProperty("In")] public InExpression In { get; set; }
            [JsonProperty("Not")] public NotExpression Not { get; set; }
            [JsonProperty("Comparison")] public ComparisonExpression Comparison { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class InExpression
        {
            [JsonProperty("Expressions")] public List<ColumnSelect> Expressions { get; set; }
            [JsonProperty("Table")] public InTable Table { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class InTable
        {
            [JsonProperty("SourceRef")] public VisualDto.SourceRef SourceRef { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class NotExpression
        {
            [JsonProperty("Expression")] public Condition Expression { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class ComparisonExpression
        {
            [JsonProperty("ComparisonKind")] public int ComparisonKind { get; set; }
            [JsonProperty("Left")] public FilterOperand Left { get; set; }
            [JsonProperty("Right")] public FilterOperand Right { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class FilterOperand
        {
            [JsonProperty("Measure")] public VisualDto.MeasureObject Measure { get; set; }
            [JsonProperty("Column")] public VisualDto.ColumnField Column { get; set; }
            [JsonProperty("Literal")] public LiteralOperand Literal { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class LiteralOperand
        {
            [JsonProperty("Value")] public string Value { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }


        public class PropertiesConverter : JsonConverter
        {
            public override bool CanConvert(Type objectType) => objectType == typeof(Dictionary<string, object>);

            public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
            {
                var result = new Dictionary<string, object>();
                var jObj = JObject.Load(reader);

                foreach (var prop in jObj.Properties())
                {
                    if (prop.Name == "paragraphs")
                    {
                        var paragraphs = prop.Value.ToObject<List<Paragraph>>(serializer);
                        result[prop.Name] = paragraphs;
                    }
                    else
                    {
                        var visualProp = prop.Value.ToObject<VisualObjectProperty>(serializer);
                        result[prop.Name] = visualProp;
                    }
                }

                return result;
            }

            public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
            {
                var dict = (Dictionary<string, object>)value;
                writer.WriteStartObject();

                foreach (var kvp in dict)
                {
                    writer.WritePropertyName(kvp.Key);

                    if (kvp.Value is VisualObjectProperty vo)
                        serializer.Serialize(writer, vo);
                    else if (kvp.Value is List<Paragraph> ps)
                        serializer.Serialize(writer, ps);
                    else
                        serializer.Serialize(writer, kvp.Value);
                }

                writer.WriteEndObject();
            }
        }
    }

    public class VisualExtended
    {
        public VisualDto.Root Content { get; set; }

        public string VisualFilePath { get; set; }

        public bool isVisualGroup => Content?.VisualGroup != null;
        public bool isGroupedVisual => Content?.ParentGroupName != null;

        public bool IsBilingualVisualGroup()
        {
            if (!isVisualGroup || string.IsNullOrEmpty(Content.VisualGroup.DisplayName))
                return false;
            return System.Text.RegularExpressions.Regex.IsMatch(Content.VisualGroup.DisplayName, @"^P\d{2}-\d{3}$");
        }

        public PageExtended ParentPage { get; set; }

        public bool IsInBilingualVisualGroup()
        {
            if (ParentPage == null || ParentPage.Visuals == null || Content.ParentGroupName == null)
                return false;
            return ParentPage.Visuals.Any(v => v.IsBilingualVisualGroup() && v.Content.Name == Content.ParentGroupName);
        }

        [JsonIgnore]
        public string AltText
        {
            get
            {
                var general = Content?.Visual?.VisualContainerObjects?.General;
                if (general == null || general.Count == 0)
                    return null;
                if (!general[0].Properties.ContainsKey("altText"))
                    return null;
                return general[0].Properties["altText"]?.Expr?.Literal?.Value?.Trim('\'');
            }
            set
            {
                if (Content?.Visual == null)
                    Content.Visual = new VisualDto.Visual();

                if (Content?.Visual?.VisualContainerObjects == null)
                    Content.Visual.VisualContainerObjects = new VisualDto.VisualContainerObjects();

                if (Content.Visual?.VisualContainerObjects.General == null || Content.Visual?.VisualContainerObjects.General.Count == 0)
                    Content.Visual.VisualContainerObjects.General =
                        new List<VisualDto.VisualContainerObject> {
                        new VisualDto.VisualContainerObject {
                            Properties = new Dictionary<string, VisualDto.VisualContainerProperty>()
                        }
                        };

                var general = Content.Visual.VisualContainerObjects.General[0];

                if (general.Properties == null)
                    general.Properties = new Dictionary<string, VisualDto.VisualContainerProperty>();

                general.Properties["altText"] = new VisualDto.VisualContainerProperty
                {
                    Expr = new VisualDto.VisualExpr
                    {
                        Literal = new VisualDto.VisualLiteral
                        {
                            Value = value == null ? null : "'" + value.Replace("'", "\\'") + "'"
                        }
                    }
                };
            }
        }

        private IEnumerable<VisualDto.Field> GetAllFields()
        {
            var fields = new List<VisualDto.Field>();
            var queryState = Content?.Visual?.Query?.QueryState;

            if (queryState != null)
            {
                fields.AddRange(GetFieldsFromProjections(queryState.Values));
                fields.AddRange(GetFieldsFromProjections(queryState.Y));
                fields.AddRange(GetFieldsFromProjections(queryState.Y2));
                fields.AddRange(GetFieldsFromProjections(queryState.Category));
                fields.AddRange(GetFieldsFromProjections(queryState.Series));
                fields.AddRange(GetFieldsFromProjections(queryState.Data));
                fields.AddRange(GetFieldsFromProjections(queryState.Rows));
            }

            var sortList = Content?.Visual?.Query?.SortDefinition?.Sort;
            if (sortList != null)
                fields.AddRange(sortList.Select(s => s.Field));

            var objects = Content?.Visual?.Objects;
            if (objects != null)
            {
                fields.AddRange(GetFieldsFromObjectList(objects.DataPoint));
                fields.AddRange(GetFieldsFromObjectList(objects.Data));
                fields.AddRange(GetFieldsFromObjectList(objects.Labels));
                fields.AddRange(GetFieldsFromObjectList(objects.Title));
                fields.AddRange(GetFieldsFromObjectList(objects.Legend));
                fields.AddRange(GetFieldsFromObjectList(objects.General));
                fields.AddRange(GetFieldsFromObjectList(objects.ValueAxis));
                fields.AddRange(GetFieldsFromObjectList(objects.Y1AxisReferenceLine));
                fields.AddRange(GetFieldsFromObjectList(objects.ReferenceLabel));
                fields.AddRange(GetFieldsFromObjectList(objects.ReferenceLabelDetail));
                fields.AddRange(GetFieldsFromObjectList(objects.ReferenceLabelValue));
            }

            fields.AddRange(GetFieldsFromFilterConfig(Content?.FilterConfig as VisualDto.FilterConfig));

            return fields.Where(f => f != null);
        }

        private IEnumerable<VisualDto.Field> GetFieldsFromProjections(VisualDto.ProjectionsSet set)
        {
            return set?.Projections?.Select(p => p.Field) ?? Enumerable.Empty<VisualDto.Field>();
        }

        private IEnumerable<VisualDto.Field> GetFieldsFromObjectList(List<VisualDto.ObjectProperties> objectList)
        {
            if (objectList == null) yield break;

            foreach (var obj in objectList)
            {
                if (obj.Properties == null) continue;

                foreach (var val in obj.Properties.Values)
                {
                    var prop = val as VisualDto.VisualObjectProperty;
                    if (prop == null) continue;

                    if (prop.Expr != null)
                    {
                        if (prop.Expr.Measure != null)
                            yield return new VisualDto.Field { Measure = prop.Expr.Measure };

                        if (prop.Expr.Column != null)
                            yield return new VisualDto.Field { Column = prop.Expr.Column };
                    }

                    if (prop.Color?.Expr?.FillRule?.Input != null)
                        yield return prop.Color.Expr.FillRule.Input;

                    if (prop.Solid?.Color?.Expr?.FillRule?.Input != null)
                        yield return prop.Solid.Color.Expr.FillRule.Input;

                    var solidExpr = prop.Solid?.Color?.Expr;
                    if (solidExpr?.Measure != null)
                        yield return new VisualDto.Field { Measure = solidExpr.Measure };
                    if (solidExpr?.Column != null)
                        yield return new VisualDto.Field { Column = solidExpr.Column };
                }
            }
        }

        private IEnumerable<VisualDto.Field> GetFieldsFromFilterConfig(VisualDto.FilterConfig filterConfig)
        {
            var fields = new List<VisualDto.Field>();

            if (filterConfig?.Filters == null)
                return fields;

            foreach (var filter in filterConfig.Filters ?? Enumerable.Empty<VisualDto.VisualFilter>())
            {
                if (filter.Field != null)
                    fields.Add(filter.Field);

                if (filter.Filter != null)
                {
                    var aliasMap = BuildAliasMap(filter.Filter.From);

                    foreach (var from in filter.Filter.From ?? Enumerable.Empty<VisualDto.FilterFrom>())
                    {
                        if (from.Expression?.Subquery?.Query != null)
                            ExtractFieldsFromSubquery(from.Expression.Subquery.Query, fields);
                    }

                    foreach (var where in filter.Filter.Where ?? Enumerable.Empty<VisualDto.FilterWhere>())
                        ExtractFieldsFromCondition(where.Condition, fields, aliasMap);
                }
            }

            return fields;
        }

        private void ExtractFieldsFromSubquery(VisualDto.SubqueryQuery query, List<VisualDto.Field> fields)
        {
            var aliasMap = BuildAliasMap(query.From);

            // SELECT columns
            foreach (var sel in query.Select ?? Enumerable.Empty<VisualDto.SelectExpression>())
            {
                var srcRef = sel.Column?.Expression?.SourceRef ?? new VisualDto.SourceRef();
                srcRef.Source = ResolveSource(srcRef.Source, aliasMap);

                var columnExpr = sel.Column ?? new VisualDto.ColumnSelect();
                columnExpr.Expression ??= new VisualDto.Expression();
                columnExpr.Expression.SourceRef ??= new VisualDto.SourceRef();
                columnExpr.Expression.SourceRef.Source = srcRef.Source;

                fields.Add(new VisualDto.Field
                {
                    Column = new VisualDto.ColumnField
                    {
                        Property = sel.Column.Property,
                        Expression = new VisualDto.Expression
                        {
                            SourceRef = columnExpr.Expression.SourceRef
                        }
                    }
                });
            }

            // ORDER BY measures
            foreach (var ob in query.OrderBy ?? Enumerable.Empty<VisualDto.OrderByExpression>())
            {
                var measureExpr = ob.Expression?.Measure?.Expression ?? new VisualDto.Expression();
                measureExpr.SourceRef ??= new VisualDto.SourceRef();
                measureExpr.SourceRef.Source = ResolveSource(measureExpr.SourceRef.Source, aliasMap);

                fields.Add(new VisualDto.Field
                {
                    Measure = new VisualDto.MeasureObject
                    {
                        Property = ob.Expression.Measure.Property,
                        Expression = measureExpr
                    }
                });
            }

            // Nested subqueries
            foreach (var from in query.From ?? Enumerable.Empty<VisualDto.FilterFrom>())
                if (from.Expression?.Subquery?.Query != null)
                    ExtractFieldsFromSubquery(from.Expression.Subquery.Query, fields);

            // WHERE conditions
            foreach (var where in query.Where ?? Enumerable.Empty<VisualDto.FilterWhere>())
                ExtractFieldsFromCondition(where.Condition, fields, aliasMap);
        }
        private Dictionary<string, string> BuildAliasMap(List<VisualDto.FilterFrom> fromList)
        {
            var map = new Dictionary<string, string>();
            foreach (var from in fromList ?? Enumerable.Empty<VisualDto.FilterFrom>())
            {
                if (!string.IsNullOrEmpty(from.Name) && !string.IsNullOrEmpty(from.Entity))
                    map[from.Name] = from.Entity;
            }
            return map;
        }

        private string ResolveSource(string source, Dictionary<string, string> aliasMap)
        {
            if (string.IsNullOrEmpty(source))
                return source;
            return aliasMap.TryGetValue(source, out var entity) ? entity : source;
        }

        private void ExtractFieldsFromCondition(VisualDto.Condition condition, List<VisualDto.Field> fields, Dictionary<string, string> aliasMap)
        {
            if (condition == null) return;

            // IN Expression
            if (condition.In != null)
            {
                foreach (var expr in condition.In.Expressions ?? Enumerable.Empty<VisualDto.ColumnSelect>())
                {
                    var srcRef = expr.Expression?.SourceRef ?? new VisualDto.SourceRef();
                    srcRef.Source = ResolveSource(srcRef.Source, aliasMap);

                    fields.Add(new VisualDto.Field
                    {
                        Column = new VisualDto.ColumnField
                        {
                            Property = expr.Property,
                            Expression = new VisualDto.Expression
                            {
                                SourceRef = srcRef
                            }
                        }
                    });
                }
            }

            // NOT Expression
            if (condition.Not != null)
                ExtractFieldsFromCondition(condition.Not.Expression, fields, aliasMap);

            // COMPARISON Expression
            if (condition.Comparison != null)
            {
                AddOperandField(condition.Comparison.Left, fields, aliasMap);
                AddOperandField(condition.Comparison.Right, fields, aliasMap);
            }
        }
        private void AddOperandField(VisualDto.FilterOperand operand, List<VisualDto.Field> fields, Dictionary<string, string> aliasMap)
        {
            if (operand == null) return;

            // MEASURE
            if (operand.Measure != null)
            {
                var srcRef = operand.Measure.Expression?.SourceRef ?? new VisualDto.SourceRef();
                srcRef.Source = ResolveSource(srcRef.Source, aliasMap);

                fields.Add(new VisualDto.Field
                {
                    Measure = new VisualDto.MeasureObject
                    {
                        Property = operand.Measure.Property,
                        Expression = new VisualDto.Expression
                        {
                            SourceRef = srcRef
                        }
                    }
                });
            }

            // COLUMN
            if (operand.Column != null)
            {
                var srcRef = operand.Column.Expression?.SourceRef ?? new VisualDto.SourceRef();
                srcRef.Source = ResolveSource(srcRef.Source, aliasMap);

                fields.Add(new VisualDto.Field
                {
                    Column = new VisualDto.ColumnField
                    {
                        Property = operand.Column.Property,
                        Expression = new VisualDto.Expression
                        {
                            SourceRef = srcRef
                        }
                    }
                });
            }
        }
        public IEnumerable<string> GetAllReferencedMeasures()
        {
            return GetAllFields()
                .Select(f => f.Measure)
                .Where(m => m?.Expression?.SourceRef?.Entity != null && m.Property != null)
                .Select(m => $"'{m.Expression.SourceRef.Entity}'[{m.Property}]")
                .Distinct();
        }

        public IEnumerable<string> GetAllReferencedColumns()
        {
            return GetAllFields()
                .Select(f => f.Column)
                .Where(c => c?.Expression?.SourceRef?.Entity != null && c.Property != null)
                .Select(c => $"'{c.Expression.SourceRef.Entity}'[{c.Property}]")
                .Distinct();
        }

        public void ReplaceMeasure(string oldFieldKey, Measure newMeasure, HashSet<VisualExtended> modifiedSet = null)
        {
            var newField = new VisualDto.Field
            {
                Measure = new VisualDto.MeasureObject
                {
                    Property = newMeasure.Name,
                    Expression = new VisualDto.Expression
                    {
                        SourceRef = new VisualDto.SourceRef { Entity = newMeasure.Table.Name }
                    }
                }
            };
            ReplaceField(oldFieldKey, newField, isMeasure: true, modifiedSet);
        }

        public void ReplaceColumn(string oldFieldKey, Column newColumn, HashSet<VisualExtended> modifiedSet = null)
        {
            var newField = new VisualDto.Field
            {
                Column = new VisualDto.ColumnField
                {
                    Property = newColumn.Name,
                    Expression = new VisualDto.Expression
                    {
                        SourceRef = new VisualDto.SourceRef { Entity = newColumn.Table.Name }
                    }
                }
            };
            ReplaceField(oldFieldKey, newField, isMeasure: false, modifiedSet);
        }

        private string ToFieldKey(VisualDto.Field f)
        {
            if (f?.Measure?.Expression?.SourceRef?.Entity is string mEntity && f.Measure.Property is string mProp)
                return $"'{mEntity}'[{mProp}]";

            if (f?.Column?.Expression?.SourceRef?.Entity is string cEntity && f.Column.Property is string cProp)
                return $"'{cEntity}'[{cProp}]";

            return null;
        }

        private void ReplaceField(string oldFieldKey, VisualDto.Field newField, bool isMeasure, HashSet<VisualExtended> modifiedSet = null)
        {
            var query = Content?.Visual?.Query;
            var objects = Content?.Visual?.Objects;
            bool wasModified = false;

            void Replace(VisualDto.Field f)
            {
                if (f == null) return;

                if (isMeasure && newField.Measure != null)
                {
                    // Preserve Expression with SourceRef
                    f.Measure ??= new VisualDto.MeasureObject();
                    f.Measure.Property = newField.Measure.Property;
                    f.Measure.Expression ??= new VisualDto.Expression();
                    f.Measure.Expression.SourceRef = newField.Measure.Expression?.SourceRef != null
                        ? new VisualDto.SourceRef
                        {
                            Entity = newField.Measure.Expression.SourceRef.Entity,
                            Source = newField.Measure.Expression.SourceRef.Source
                        }
                        : f.Measure.Expression.SourceRef;
                    f.Column = null;
                    wasModified = true;
                }
                else if (!isMeasure && newField.Column != null)
                {
                    // Preserve Expression with SourceRef
                    f.Column ??= new VisualDto.ColumnField();
                    f.Column.Property = newField.Column.Property;
                    f.Column.Expression ??= new VisualDto.Expression();
                    f.Column.Expression.SourceRef = newField.Column.Expression?.SourceRef != null
                        ? new VisualDto.SourceRef
                        {
                            Entity = newField.Column.Expression.SourceRef.Entity,
                            Source = newField.Column.Expression.SourceRef.Source
                        }
                        : f.Column.Expression.SourceRef;
                    f.Measure = null;
                    wasModified = true;
                }
            }

            void UpdateProjection(VisualDto.Projection proj)
            {
                if (proj == null) return;

                if (ToFieldKey(proj.Field) == oldFieldKey)
                {
                    Replace(proj.Field);

                    string entity = isMeasure
                        ? proj.Field.Measure.Expression?.SourceRef?.Entity
                        : proj.Field.Column.Expression?.SourceRef?.Entity;

                    string prop = isMeasure
                        ? proj.Field.Measure.Property
                        : proj.Field.Column.Property;

                    if (!string.IsNullOrEmpty(entity) && !string.IsNullOrEmpty(prop))
                    {
                        proj.QueryRef = $"{entity}.{prop}";
                    }

                    wasModified = true;
                }
            }

            foreach (var proj in query?.QueryState?.Values?.Projections ?? Enumerable.Empty<VisualDto.Projection>())
                UpdateProjection(proj);

            foreach (var proj in query?.QueryState?.Y?.Projections ?? Enumerable.Empty<VisualDto.Projection>())
                UpdateProjection(proj);

            foreach (var proj in query?.QueryState?.Y2?.Projections ?? Enumerable.Empty<VisualDto.Projection>())
                UpdateProjection(proj);

            foreach (var proj in query?.QueryState?.Category?.Projections ?? Enumerable.Empty<VisualDto.Projection>())
                UpdateProjection(proj);

            foreach (var proj in query?.QueryState?.Series?.Projections ?? Enumerable.Empty<VisualDto.Projection>())
                UpdateProjection(proj);

            foreach (var proj in query?.QueryState?.Data?.Projections ?? Enumerable.Empty<VisualDto.Projection>())
                UpdateProjection(proj);

            foreach (var proj in query?.QueryState?.Rows?.Projections ?? Enumerable.Empty<VisualDto.Projection>())
                UpdateProjection(proj);

            foreach (var sort in query?.SortDefinition?.Sort ?? Enumerable.Empty<VisualDto.Sort>())
                if (ToFieldKey(sort.Field) == oldFieldKey) Replace(sort.Field);

            string oldMetadata = oldFieldKey.Replace("'", "").Replace("[", ".").Replace("]", "");
            string newMetadata = isMeasure
                ? $"{newField.Measure.Expression.SourceRef.Entity}.{newField.Measure.Property}"
                : $"{newField.Column.Expression.SourceRef.Entity}.{newField.Column.Property}";

            IEnumerable<VisualDto.ObjectProperties> AllObjectProperties() =>
                (objects?.DataPoint ?? Enumerable.Empty<VisualDto.ObjectProperties>())
                .Concat(objects?.Data ?? Enumerable.Empty<VisualDto.ObjectProperties>())
                .Concat(objects?.Labels ?? Enumerable.Empty<VisualDto.ObjectProperties>())
                .Concat(objects?.Title ?? Enumerable.Empty<VisualDto.ObjectProperties>())
                .Concat(objects?.Legend ?? Enumerable.Empty<VisualDto.ObjectProperties>())
                .Concat(objects?.General ?? Enumerable.Empty<VisualDto.ObjectProperties>())
                .Concat(objects?.ValueAxis ?? Enumerable.Empty<VisualDto.ObjectProperties>())
                .Concat(objects?.ReferenceLabel ?? Enumerable.Empty<VisualDto.ObjectProperties>())
                .Concat(objects?.ReferenceLabelDetail ?? Enumerable.Empty<VisualDto.ObjectProperties>())
                .Concat(objects?.ReferenceLabelValue ?? Enumerable.Empty<VisualDto.ObjectProperties>())
                .Concat(objects?.Values ?? Enumerable.Empty<VisualDto.ObjectProperties>())
                .Concat(objects?.Y1AxisReferenceLine ?? Enumerable.Empty<VisualDto.ObjectProperties>());

            foreach (var obj in AllObjectProperties())
            {
                foreach (var prop in obj.Properties.Values.OfType<VisualDto.VisualObjectProperty>())
                {
                    var field = isMeasure ? new VisualDto.Field { Measure = prop.Expr?.Measure } : new VisualDto.Field { Column = prop.Expr?.Column };
                    if (ToFieldKey(field) == oldFieldKey)
                    {
                        if (prop.Expr != null)
                        {
                            if (isMeasure)
                            {
                                prop.Expr.Measure ??= new VisualDto.MeasureObject();
                                prop.Expr.Measure.Property = newField.Measure.Property;
                                prop.Expr.Measure.Expression ??= new VisualDto.Expression();
                                prop.Expr.Measure.Expression.SourceRef = newField.Measure.Expression?.SourceRef;
                                prop.Expr.Column = null;
                                wasModified = true;
                            }
                            else
                            {
                                prop.Expr.Column ??= new VisualDto.ColumnField();
                                prop.Expr.Column.Property = newField.Column.Property;
                                prop.Expr.Column.Expression ??= new VisualDto.Expression();
                                prop.Expr.Column.Expression.SourceRef = newField.Column.Expression?.SourceRef;
                                prop.Expr.Measure = null;
                                wasModified = true;
                            }
                        }
                    }

                    var fillInput = prop.Color?.Expr?.FillRule?.Input;
                    if (ToFieldKey(fillInput) == oldFieldKey)
                    {
                        if (isMeasure)
                        {
                            fillInput.Measure ??= new VisualDto.MeasureObject();
                            fillInput.Measure.Property = newField.Measure.Property;
                            fillInput.Measure.Expression ??= new VisualDto.Expression();
                            fillInput.Measure.Expression.SourceRef = newField.Measure.Expression?.SourceRef;
                            fillInput.Column = null;
                            wasModified = true;
                        }
                        else
                        {
                            fillInput.Column ??= new VisualDto.ColumnField();
                            fillInput.Column.Property = newField.Column.Property;
                            fillInput.Column.Expression ??= new VisualDto.Expression();
                            fillInput.Column.Expression.SourceRef = newField.Column.Expression?.SourceRef;
                            fillInput.Measure = null;
                            wasModified = true;
                        }
                    }

                    var solidInput = prop.Solid?.Color?.Expr?.FillRule?.Input;
                    if (ToFieldKey(solidInput) == oldFieldKey)
                    {
                        if (isMeasure)
                        {
                            solidInput.Measure ??= new VisualDto.MeasureObject();
                            solidInput.Measure.Property = newField.Measure.Property;
                            solidInput.Measure.Expression ??= new VisualDto.Expression();
                            solidInput.Measure.Expression.SourceRef = newField.Measure.Expression?.SourceRef;
                            solidInput.Column = null;
                            wasModified = true;
                        }
                        else
                        {
                            solidInput.Column ??= new VisualDto.ColumnField();
                            solidInput.Column.Property = newField.Column.Property;
                            solidInput.Column.Expression ??= new VisualDto.Expression();
                            solidInput.Column.Expression.SourceRef = newField.Column.Expression?.SourceRef;
                            solidInput.Measure = null;
                            wasModified = true;
                        }
                    }

                    var solidExpr = prop.Solid?.Color?.Expr;
                    if (solidExpr != null)
                    {
                        var solidField = isMeasure
                            ? new VisualDto.Field { Measure = solidExpr.Measure }
                            : new VisualDto.Field { Column = solidExpr.Column };

                        if (ToFieldKey(solidField) == oldFieldKey)
                        {
                            if (isMeasure)
                            {
                                solidExpr.Measure ??= new VisualDto.MeasureObject();
                                solidExpr.Measure.Property = newField.Measure.Property;
                                solidExpr.Measure.Expression ??= new VisualDto.Expression();
                                solidExpr.Measure.Expression.SourceRef = newField.Measure.Expression?.SourceRef;
                                solidExpr.Column = null;
                                wasModified = true;
                            }
                            else
                            {
                                solidExpr.Column ??= new VisualDto.ColumnField();
                                solidExpr.Column.Property = newField.Column.Property;
                                solidExpr.Column.Expression ??= new VisualDto.Expression();
                                solidExpr.Column.Expression.SourceRef = newField.Column.Expression?.SourceRef;
                                solidExpr.Measure = null;
                                wasModified = true;
                            }
                        }
                    }
                }

                if (obj.Selector?.Metadata == oldMetadata)
                {
                    obj.Selector.Metadata = newMetadata;
                    wasModified = true;
                }
            }

            //if (Content.FilterConfig != null)
            //{
            //    var filterConfigString = Content.FilterConfig.ToString();
            //    string table = isMeasure ? newField.Measure.Expression.SourceRef.Entity : newField.Column.Expression.SourceRef.Entity;
            //    string prop = isMeasure ? newField.Measure.Property : newField.Column.Property;

            //    string oldPattern = oldFieldKey;
            //    string newPattern = $"'{table}'[{prop}]";

            //    if (filterConfigString.Contains(oldPattern))
            //    {
            //        Content.FilterConfig = filterConfigString.Replace(oldPattern, newPattern);
            //        wasModified = true;
            //    }
            //}
            if (wasModified && modifiedSet != null)
                modifiedSet.Add(this);
        }

    }

    public class PageExtended
    {
        public PageDto Page { get; set; }

        public ReportExtended ParentReport { get; set; }

        public int PageIndex
        {
            get
            {
                if (ParentReport == null || ParentReport.PagesConfig == null || ParentReport.PagesConfig.PageOrder == null)
                    return -1;
                return ParentReport.PagesConfig.PageOrder.IndexOf(Page.Name);
            }
        }


        public IList<VisualExtended> Visuals { get; set; } = new List<VisualExtended>();
        public string PageFilePath { get; set; }
    }

    public class ReportExtended
    {
        public IList<PageExtended> Pages { get; set; } = new List<PageExtended>();
        public string PagesFilePath { get; set; }
        public PagesDto PagesConfig { get; set; }
    }

}
