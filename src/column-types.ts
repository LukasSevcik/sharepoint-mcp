import type { ColumnSpec, ColumnType } from "./types.js";

interface FieldMeta {
  fieldTypeKind: number;
  metaType: string;
}

const TYPE_MAP: Record<ColumnType, FieldMeta> = {
  text:        { fieldTypeKind: 2,  metaType: "SP.FieldText" },
  note:        { fieldTypeKind: 3,  metaType: "SP.FieldMultiLineText" },
  number:      { fieldTypeKind: 9,  metaType: "SP.FieldNumber" },
  currency:    { fieldTypeKind: 10, metaType: "SP.FieldCurrency" },
  date:        { fieldTypeKind: 4,  metaType: "SP.FieldDateTime" },
  boolean:     { fieldTypeKind: 8,  metaType: "SP.Field" },
  choice:      { fieldTypeKind: 6,  metaType: "SP.FieldChoice" },
  multichoice: { fieldTypeKind: 15, metaType: "SP.FieldMultiChoice" },
  person:      { fieldTypeKind: 20, metaType: "SP.FieldUser" },
  url:         { fieldTypeKind: 11, metaType: "SP.FieldUrl" },
  calculated:  { fieldTypeKind: 17, metaType: "SP.FieldCalculated" },
};

const OUTPUT_TYPE_KIND: Record<string, number> = {
  text: 2, number: 9, boolean: 8, date: 4, currency: 10,
};

export function buildFieldBody(spec: ColumnSpec): Record<string, unknown> {
  const { fieldTypeKind, metaType } = TYPE_MAP[spec.type];

  const base: Record<string, unknown> = {
    "__metadata": { "type": metaType },
    "Title": spec.displayName ?? spec.name,
    "StaticName": spec.name,
    "FieldTypeKind": fieldTypeKind,
    "Required": spec.required ?? false,
    "Description": spec.description ?? "",
    "EnforceUniqueValues": false,
    "Indexed": spec.indexed ?? false,
  };

  switch (spec.type) {
    case "note":
      base["RichText"] = spec.richText ?? false;
      base["AllowHyperlink"] = spec.richText ?? false;
      break;
    case "number":
      if (spec.min !== undefined) base["MinimumValue"] = spec.min;
      if (spec.max !== undefined) base["MaximumValue"] = spec.max;
      if (spec.decimals !== undefined) base["DisplayFormat"] = spec.decimals;
      break;
    case "currency":
      base["CurrencyLocaleId"] = 1033;
      if (spec.decimals !== undefined) base["DisplayFormat"] = spec.decimals;
      break;
    case "date":
      base["DisplayFormat"] = spec.dateOnly ? 0 : 1;
      base["FriendlyDisplayFormat"] = 0;
      break;
    case "choice":
    case "multichoice":
      base["Choices"] = {
        "__metadata": { "type": "Collection(Edm.String)" },
        "results": spec.choices ?? [],
      };
      if (spec.defaultValue !== undefined) base["DefaultValue"] = String(spec.defaultValue);
      break;
    case "person":
      base["AllowMultipleValues"] = spec.multiple ?? false;
      base["SelectionMode"] = 0;
      break;
    case "boolean":
      if (spec.defaultValue !== undefined) {
        base["DefaultValue"] = spec.defaultValue ? "1" : "0";
      }
      break;
    case "calculated":
      if (spec.formula) base["Formula"] = spec.formula;
      base["OutputType"] = OUTPUT_TYPE_KIND[spec.outputType ?? "text"] ?? 2;
      break;
    default:
      if (spec.defaultValue !== undefined) base["DefaultValue"] = String(spec.defaultValue);
  }

  return base;
}

export function fieldTypeKindToLabel(kind: number): string {
  const map: Record<number, string> = {
    1: "integer", 2: "text", 3: "note", 4: "date", 6: "choice",
    7: "lookup", 8: "boolean", 9: "number", 10: "currency", 11: "url",
    12: "computed", 15: "multichoice", 17: "calculated", 20: "person",
  };
  return map[kind] ?? `fieldType${kind}`;
}
