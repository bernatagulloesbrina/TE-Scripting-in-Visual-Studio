﻿using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TabularEditor.TOMWrapper;
using TabularEditor.Scripting;
using Newtonsoft.Json.Linq;

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
            [JsonProperty("visualContainerObjects")] public object VisualContainerObjects { get; set; }
            [JsonProperty("filterConfig")] public object FilterConfig { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class Position
        {
            [JsonProperty("x")] public double X { get; set; }
            [JsonProperty("y")] public double Y { get; set; }
            [JsonProperty("z")] public int Z { get; set; }
            [JsonProperty("height")] public double Height { get; set; }
            [JsonProperty("width")] public double Width { get; set; }
            [JsonProperty("tabOrder")] public int TabOrder { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class Visual
        {
            [JsonProperty("visualType", Order = 1)] public string VisualType { get; set; }
            [JsonProperty("query", Order = 2)] public Query Query { get; set; }
            [JsonProperty("objects", Order = 3)] public Objects Objects { get; set; }
            [JsonProperty("drillFilterOtherVisuals", Order = 4)] public bool DrillFilterOtherVisuals { get; set; }
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

        public class QueryState
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


            [JsonProperty("referenceLabel")] public List<VisualDto.ObjectProperties> ReferenceLabel { get; set; }
            [JsonProperty("referenceLabelDetail")] public List<VisualDto.ObjectProperties> ReferenceLabelDetail { get; set; }
            [JsonProperty("referenceLabelValue")] public List<VisualDto.ObjectProperties> ReferenceLabelValue { get; set; }

            [JsonProperty("values")] public List<VisualDto.ObjectProperties> Values { get; set; }


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
                fields.AddRange(GetFieldsFromObjectList(objects.ReferenceLabel));
                fields.AddRange(GetFieldsFromObjectList(objects.ReferenceLabelDetail));
                fields.AddRange(GetFieldsFromObjectList(objects.ReferenceLabelValue));

            }

            fields.AddRange(GetFieldsFromFilterConfig(Content?.FilterConfig));

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

                    if (prop.Color != null &&
                        prop.Color.Expr != null &&
                        prop.Color.Expr.FillRule != null &&
                        prop.Color.Expr.FillRule.Input != null)
                    {
                        yield return prop.Color.Expr.FillRule.Input;
                    }

                    if (prop.Solid != null &&
                        prop.Solid.Color != null &&
                        prop.Solid.Color.Expr != null &&
                        prop.Solid.Color.Expr.FillRule != null &&
                        prop.Solid.Color.Expr.FillRule.Input != null)
                    {
                        yield return prop.Solid.Color.Expr.FillRule.Input;
                    }

                    var solidExpr = prop.Solid != null &&
                                    prop.Solid.Color != null
                                    ? prop.Solid.Color.Expr
                                    : null;

                    if (solidExpr != null)
                    {
                        if (solidExpr.Measure != null)
                            yield return new VisualDto.Field { Measure = solidExpr.Measure };

                        if (solidExpr.Column != null)
                            yield return new VisualDto.Field { Column = solidExpr.Column };
                    }
                }
            }
        }

        private IEnumerable<VisualDto.Field> GetFieldsFromFilterConfig(object filterConfig)
        {
            var fields = new List<VisualDto.Field>();

            if (filterConfig is JObject jObj)
            {
                foreach (var token in jObj.DescendantsAndSelf().OfType<JObject>())
                {
                    var table = token["table"]?.ToString();
                    var property = token["column"]?.ToString() ?? token["measure"]?.ToString();

                    if (!string.IsNullOrEmpty(table) && !string.IsNullOrEmpty(property))
                    {
                        var field = new VisualDto.Field();

                        if (token["measure"] != null)
                        {
                            field.Measure = new VisualDto.MeasureObject
                            {
                                Property = property,
                                Expression = new VisualDto.Expression
                                {
                                    SourceRef = new VisualDto.SourceRef { Entity = table }
                                }
                            };
                        }
                        else if (token["column"] != null)
                        {
                            field.Column = new VisualDto.ColumnField
                            {
                                Property = property,
                                Expression = new VisualDto.Expression
                                {
                                    SourceRef = new VisualDto.SourceRef { Entity = table }
                                }
                            };
                        }

                        fields.Add(field);
                    }
                }
            }

            return fields;
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

                if (isMeasure)
                {
                    f.Measure = newField.Measure;
                    f.Column = null;
                    wasModified = true;
                }
                else
                {
                    f.Column = newField.Column;
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
                        ? newField.Measure.Expression?.SourceRef?.Entity
                        : newField.Column.Expression?.SourceRef?.Entity;

                    string prop = isMeasure
                        ? newField.Measure.Property
                        : newField.Column.Property;

                    if (!string.IsNullOrEmpty(entity) && !string.IsNullOrEmpty(prop))
                    {
                        proj.QueryRef = $"{entity}.{prop}";
                        //proj.NativeQueryRef = prop;
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
                .Concat(objects?.Values ?? Enumerable.Empty<VisualDto.ObjectProperties>());

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
                                prop.Expr.Measure = newField.Measure;
                                prop.Expr.Column = null;
                                wasModified = true;
                            }
                            else
                            {
                                prop.Expr.Column = newField.Column;
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
                            fillInput.Measure = newField.Measure;
                            fillInput.Column = null;
                            wasModified = true;
                        }
                        else
                        {
                            fillInput.Column = newField.Column;
                            fillInput.Measure = null;
                            wasModified = true;
                        }
                    }

                    var solidInput = prop.Solid?.Color?.Expr?.FillRule?.Input;
                    if (ToFieldKey(solidInput) == oldFieldKey)
                    {
                        if (isMeasure)
                        {
                            solidInput.Measure = newField.Measure;
                            solidInput.Column = null;
                            wasModified = true;
                        }
                        else
                        {
                            solidInput.Column = newField.Column;
                            solidInput.Measure = null;
                            wasModified = true;
                        }
                    }

                    // ✅ NEW: handle direct measure/column under solid.color.expr
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
                                solidExpr.Measure = newField.Measure;
                                solidExpr.Column = null;
                                wasModified = true;
                            }
                            else
                            {
                                solidExpr.Column = newField.Column;
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

            if (Content.FilterConfig != null)
            {
                var filterConfigString = Content.FilterConfig.ToString();
                string table = isMeasure ? newField.Measure.Expression.SourceRef.Entity : newField.Column.Expression.SourceRef.Entity;
                string prop = isMeasure ? newField.Measure.Property : newField.Column.Property;

                string oldPattern = oldFieldKey;
                string newPattern = $"'{table}'[{prop}]";

                if (filterConfigString.Contains(oldPattern))
                {
                    Content.FilterConfig = filterConfigString.Replace(oldPattern, newPattern);
                    wasModified = true;
                }
            }
            if (wasModified && modifiedSet != null)
                modifiedSet.Add(this);

        }

        public void ReplaceInFilterConfigRaw(
            Dictionary<string, string> tableMap,
            Dictionary<string, string> fieldMap,
            HashSet<VisualExtended> modifiedVisuals = null)
        {
            if (Content.FilterConfig == null) return;

            string originalJson = JsonConvert.SerializeObject(Content.FilterConfig);
            string updatedJson = originalJson;

            foreach (var kv in tableMap)
                updatedJson = updatedJson.Replace($"\"{kv.Key}\"", $"\"{kv.Value}\"");

            foreach (var kv in fieldMap)
                updatedJson = updatedJson.Replace($"\"{kv.Key}\"", $"\"{kv.Value}\"");

            // Only update and track if something actually changed
            if (updatedJson != originalJson)
            {
                Content.FilterConfig = JsonConvert.DeserializeObject(updatedJson);
                modifiedVisuals?.Add(this);
            }
        }

    }


    public class PageExtended
    {
        public PageDto Page { get; set; }
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
