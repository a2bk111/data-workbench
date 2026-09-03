import React, { useState, useMemo, useRef, useId } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  Upload, Play, Trash2, Plus, Database, GitBranch, Table2,
  ArrowRightLeft, Download, ChevronUp, ChevronDown, X, AlertTriangle,
  CheckCircle2, FileSpreadsheet, ArrowUpDown, Search, Save, FolderOpen, Info
} from "lucide-react";

const theme = {
  bg: "#12161B",
  panel: "#181D23",
  surface: "#1E252C",
  surfaceHover: "#242C34",
  border: "#2B333C",
  borderLight: "#39424C",
  text: "#F1F4F6",
  textDim: "#A6AFB9",
  textFaint: "#767F89",
  accent: "#57D9CC",
  accentDim: "rgba(87,217,204,0.16)",
  accentText: "#8EE7DD",
  warn: "#F4C04B",
  warnDim: "rgba(244,192,75,0.16)",
  error: "#F17179",
  errorDim: "rgba(241,113,121,0.16)",
  success: "#72E0A0",
  successDim: "rgba(114,224,160,0.16)",
  focus: "#8EE7DD",
};

const fontDisplay = "'Space Grotesk', sans-serif";
const fontBody = "'Inter', sans-serif";
const fontMono = "'JetBrains Mono', monospace";

const TYPE_COLORS = {
  date: theme.accent, number: "#8FB3F0", currency: theme.warn,
  email: "#D6A6F5", phone: theme.success, text: theme.textDim,
};

const focusRing = `0 0 0 2px ${theme.bg}, 0 0 0 4px ${theme.focus}`;

// ---------- type detection ----------
function detectType(values) {
  const sample = values.filter((v) => v !== null && v !== undefined && String(v).trim() !== "").slice(0, 30);
  if (sample.length === 0) return "text";
  const emailRe = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  const phoneRe = /^[+]?[\d][\d\-\s()]{6,}$/;
  const currencyRe = /^[₹$€£]\s?-?\d[\d,]*(\.\d+)?$/;
  const numberRe = /^-?\d[\d,]*(\.\d+)?%?$/;
  const dateOk = sample.every((v) => {
    const s = String(v).trim();
    return /^\d{4}-\d{2}-\d{2}/.test(s) || /^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/.test(s);
  });
  if (dateOk) return "date";
  if (sample.every((v) => emailRe.test(String(v).trim()))) return "email";
  if (sample.every((v) => currencyRe.test(String(v).trim()))) return "currency";
  if (sample.every((v) => numberRe.test(String(v).trim()))) return "number";
  const phoneCount = sample.filter((v) => phoneRe.test(String(v).trim())).length;
  if (phoneCount / sample.length > 0.85 && sample.some((v) => /\d{5,}/.test(String(v).replace(/\D/g, "")))) return "phone";
  return "text";
}

function columnStats(data, col) {
  let nulls = 0;
  const seen = new Set();
  data.forEach((r) => {
    const v = r[col];
    if (v === null || v === undefined || String(v).trim() === "") nulls++;
    else seen.add(String(v));
  });
  return { total: data.length, nulls, unique: seen.size };
}

// ---------- levenshtein / similarity ----------
function levenshtein(a, b) {
  a = a || ""; b = b || "";
  const m = a.length, n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;
  let prev = new Array(n + 1), curr = new Array(n + 1);
  for (let j = 0; j <= n; j++) prev[j] = j;
  for (let i = 1; i <= m; i++) {
    curr[0] = i;
    for (let j = 1; j <= n; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      curr[j] = Math.min(curr[j - 1] + 1, prev[j] + 1, prev[j - 1] + cost);
    }
    [prev, curr] = [curr, prev];
  }
  return prev[n];
}
function similarity(a, b) {
  a = String(a || "").trim().toLowerCase();
  b = String(b || "").trim().toLowerCase();
  if (a === b) return 1;
  const maxLen = Math.max(a.length, b.length);
  if (maxLen === 0) return 1;
  return 1 - levenshtein(a, b) / maxLen;
}
function compositeSimilarity(rowA, rowB, columns) {
  if (!columns || !columns.length) return 0;
  const sims = columns.map((c) => similarity(rowA[c], rowB[c]));
  return sims.reduce((a, b) => a + b, 0) / sims.length;
}

// ---------- pipeline step processors ----------
function applyStep(data, step) {
  const meta = { before: data.length };
  let out = data;
  switch (step.type) {
    case "trim": {
      out = data.map((row) => {
        const nr = { ...row };
        Object.keys(nr).forEach((k) => { if (typeof nr[k] === "string") nr[k] = nr[k].trim().replace(/\s+/g, " "); });
        return nr;
      });
      meta.detail = "Trimmed whitespace across all columns";
      break;
    }
    case "case": {
      const { column, mode } = step.params;
      out = data.map((row) => {
        const nr = { ...row };
        const v = String(nr[column] ?? "");
        if (mode === "upper") nr[column] = v.toUpperCase();
        else if (mode === "lower") nr[column] = v.toLowerCase();
        else if (mode === "title") nr[column] = v.replace(/\w\S*/g, (t) => t[0].toUpperCase() + t.slice(1).toLowerCase());
        return nr;
      });
      meta.detail = `Standardized case on "${column}"`;
      break;
    }
    case "dedupe_exact": {
      const seen = new Set();
      out = data.filter((row) => {
        const key = JSON.stringify(row);
        if (seen.has(key)) return false;
        seen.add(key); return true;
      });
      meta.detail = `Removed ${data.length - out.length} exact duplicate row(s)`;
      break;
    }
    case "dedupe_fuzzy": {
      const { columns, threshold } = step.params;
      const cols = columns && columns.length ? columns : [];
      const used = new Array(data.length).fill(false);
      const keepIdx = [];
      for (let i = 0; i < data.length; i++) {
        if (used[i]) continue;
        keepIdx.push(i); used[i] = true;
        for (let j = i + 1; j < data.length; j++) {
          if (used[j]) continue;
          if (compositeSimilarity(data[i], data[j], cols) >= threshold) used[j] = true;
        }
      }
      const keepSet = new Set(keepIdx);
      out = data.filter((_, idx) => keepSet.has(idx));
      meta.detail = `Merged ${data.length - out.length} near-duplicate row(s) on ${cols.map((c) => `"${c}"`).join(" + ") || "(no column selected)"} (≥${Math.round(threshold * 100)}% match)`;
      break;
    }
    case "split_column": {
      const { column, delimiter, names } = step.params;
      const list = (names && names.length ? names : ["", ""]).map((n, i) => n || `${column}_${i + 1}`);
      const delim = delimiter === "" ? " " : delimiter;
      out = data.map((row) => {
        const nr = { ...row };
        const parts = String(row[column] ?? "").split(delim);
        list.forEach((colName, i) => {
          if (i < list.length - 1) nr[colName] = (parts[i] ?? "").trim();
          else nr[colName] = (parts.slice(i).join(delim) || "").trim();
        });
        return nr;
      });
      meta.detail = `Split "${column}" into ${list.map((n) => `"${n}"`).join(", ")}`;
      break;
    }
    case "merge_columns": {
      const { col1, col2, separator, name } = step.params;
      out = data.map((row) => {
        const nr = { ...row };
        nr[name || `${col1}_${col2}`] = `${row[col1] ?? ""}${separator}${row[col2] ?? ""}`;
        return nr;
      });
      meta.detail = `Merged "${col1}" + "${col2}" into "${name || col1 + "_" + col2}"`;
      break;
    }
    case "fill_missing": {
      const { column, strategy, customValue } = step.params;
      const nums = data.map((r) => parseFloat(String(r[column]).replace(/[^0-9.\-]/g, ""))).filter((n) => !isNaN(n));
      let fillVal = customValue || "";
      if (strategy === "mean" && nums.length) fillVal = (nums.reduce((a, b) => a + b, 0) / nums.length).toFixed(2);
      if (strategy === "median" && nums.length) {
        const sorted = [...nums].sort((a, b) => a - b);
        fillVal = String(sorted[Math.floor(sorted.length / 2)]);
      }
      if (strategy === "mode") {
        const counts = {};
        data.forEach((r) => { const v = r[column]; if (v !== null && v !== undefined && String(v).trim() !== "") counts[v] = (counts[v] || 0) + 1; });
        let best = null, bestCount = 0;
        Object.entries(counts).forEach(([k, c]) => { if (c > bestCount) { best = k; bestCount = c; } });
        fillVal = best ?? "";
      }
      let filled = 0;
      out = data.map((row) => {
        const nr = { ...row };
        if (nr[column] === null || nr[column] === undefined || String(nr[column]).trim() === "") { nr[column] = fillVal; filled++; }
        return nr;
      });
      meta.detail = `Filled ${filled} blank value(s) in "${column}" with ${strategy === "value" ? `"${customValue}"` : strategy}`;
      break;
    }
    case "remove_column": {
      const { column } = step.params;
      out = data.map((row) => { const nr = { ...row }; delete nr[column]; return nr; });
      meta.detail = `Removed column "${column}"`;
      break;
    }
    case "keep_columns": {
      const { columns } = step.params;
      const keep = columns && columns.length ? columns : Object.keys(data[0] || {});
      out = data.map((row) => {
        const nr = {};
        keep.forEach((c) => { nr[c] = row[c]; });
        return nr;
      });
      meta.detail = `Kept ${keep.length} column(s): ${keep.join(", ")}`;
      break;
    }
    case "filter_rows": {
      const { column, operator, value, action } = step.params;
      const isBlank = (v) => v === null || v === undefined || String(v).trim() === "";
      const asNum = (v) => parseFloat(String(v).replace(/[^0-9.\-]/g, ""));
      const matches = (row) => {
        const v = row[column];
        switch (operator) {
          case "is_blank": return isBlank(v);
          case "is_not_blank": return !isBlank(v);
          case "is_zero": { const n = asNum(v); return !isNaN(n) && n === 0; }
          case "equals": return String(v ?? "").trim().toLowerCase() === String(value).trim().toLowerCase();
          case "not_equals": return String(v ?? "").trim().toLowerCase() !== String(value).trim().toLowerCase();
          case "contains": return String(v ?? "").toLowerCase().includes(String(value).toLowerCase());
          case "greater_than": { const n = asNum(v); return !isNaN(n) && n > parseFloat(value); }
          case "less_than": { const n = asNum(v); return !isNaN(n) && n < parseFloat(value); }
          default: return false;
        }
      };
      out = data.filter((row) => (action === "keep" ? matches(row) : !matches(row)));
      const removedCount = data.length - out.length;
      meta.detail = `${action === "keep" ? "Kept only" : "Removed"} rows where "${column}" ${operator.replace(/_/g, " ")}${["equals", "not_equals", "contains", "greater_than", "less_than"].includes(operator) ? ` "${value}"` : ""} (${removedCount} row(s) ${action === "keep" ? "excluded" : "removed"})`;
      break;
    }
    case "find_replace": {
      const { column, find, replace } = step.params;
      out = find
        ? data.map((row) => {
            const nr = { ...row };
            nr[column] = String(nr[column] ?? "").split(find).join(replace ?? "");
            return nr;
          })
        : data;
      meta.detail = find ? `Replaced "${find}" with "${replace || ""}" in "${column}"` : `No find text specified for "${column}" — nothing changed`;
      break;
    }
    case "clean_number": {
      const { column } = step.params;
      let cleaned = 0;
      out = data.map((row) => {
        const nr = { ...row };
        const raw = String(nr[column] ?? "");
        const num = parseFloat(raw.replace(/[^0-9.\-]/g, ""));
        if (!isNaN(num)) { nr[column] = num; cleaned++; }
        return nr;
      });
      meta.detail = `Converted ${cleaned} value(s) in "${column}" to clean numbers`;
      break;
    }
    case "rename_column": {
      const { column, newName } = step.params;
      out = data.map((row) => {
        const nr = {};
        Object.keys(row).forEach((k) => { nr[k === column ? (newName || column) : k] = row[k]; });
        return nr;
      });
      meta.detail = newName ? `Renamed "${column}" to "${newName}"` : `No new name given for "${column}" — nothing changed`;
      break;
    }
    case "sort_rows": {
      const { column, direction } = step.params;
      out = [...data].sort((a, b) => {
        const av = a[column], bv = b[column];
        const an = parseFloat(av), bn = parseFloat(bv);
        const cmp = (!isNaN(an) && !isNaN(bn)) ? an - bn : String(av ?? "").localeCompare(String(bv ?? ""));
        return direction === "desc" ? -cmp : cmp;
      });
      meta.detail = `Sorted by "${column}" (${direction === "desc" ? "descending" : "ascending"})`;
      break;
    }
    default: out = data;
  }
  meta.after = out.length;
  return { data: out, meta };
}

const STEP_LIBRARY = [
  { type: "trim", label: "Trim whitespace", group: "Clean" },
  { type: "case", label: "Standardize case", group: "Clean" },
  { type: "fill_missing", label: "Fill missing values", group: "Clean" },
  { type: "clean_number", label: "Clean number / currency", group: "Clean" },
  { type: "find_replace", label: "Find & replace text", group: "Clean" },
  { type: "filter_rows", label: "Filter rows (keep/remove)", group: "Clean" },
  { type: "dedupe_exact", label: "Remove exact duplicates", group: "Dedupe" },
  { type: "dedupe_fuzzy", label: "Fuzzy duplicate match", group: "Dedupe" },
  { type: "split_column", label: "Split column", group: "Reshape" },
  { type: "merge_columns", label: "Merge columns", group: "Reshape" },
  { type: "rename_column", label: "Rename column", group: "Reshape" },
  { type: "sort_rows", label: "Sort rows", group: "Reshape" },
  { type: "remove_column", label: "Remove a column", group: "Reshape" },
  { type: "keep_columns", label: "Keep only selected columns", group: "Reshape" },
];

function defaultParams(type, columns) {
  const c0 = columns[0] || "";
  const c1 = columns[1] || columns[0] || "";
  switch (type) {
    case "case": return { column: c0, mode: "title" };
    case "dedupe_fuzzy": return { columns: c0 ? [c0] : [], threshold: 0.85 };
    case "split_column": return { column: c0, delimiter: " ", names: ["", ""] };
    case "merge_columns": return { col1: c0, col2: c1, separator: " ", name: "" };
    case "fill_missing": return { column: c0, strategy: "value", customValue: "" };
    case "remove_column": return { column: c0 };
    case "keep_columns": return { columns: [...columns] };
    case "filter_rows": return { column: c0, operator: "is_blank", value: "", action: "remove" };
    case "find_replace": return { column: c0, find: "", replace: "" };
    case "clean_number": return { column: c0 };
    case "rename_column": return { column: c0, newName: "" };
    case "sort_rows": return { column: c0, direction: "asc" };
    default: return {};
  }
}

// ---------- file parsing / export ----------
function parseFile(file) {
  return new Promise((resolve, reject) => {
    const ext = file.name.split(".").pop().toLowerCase();
    if (ext === "csv") {
      Papa.parse(file, { header: true, skipEmptyLines: true, complete: (res) => resolve({ data: res.data, columns: res.meta.fields || [] }), error: reject });
    } else {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const wb = XLSX.read(e.target.result, { type: "array" });
          const sheet = wb.Sheets[wb.SheetNames[0]];
          const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
          resolve({ data: json, columns: json.length ? Object.keys(json[0]) : [] });
        } catch (err) { reject(err); }
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    }
  });
}
function downloadBlob(content, filename, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}
function exportCSV(data, filename) { downloadBlob(Papa.unparse(data), filename, "text/csv"); }
function exportXLSX(data, filename) {
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Cleaned Data");
  const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  downloadBlob(out, filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
}

// ---------- accessible UI primitives ----------
function Btn({ children, onClick, variant = "default", disabled, style, title, ariaLabel }) {
  const base = {
    fontFamily: fontBody, fontSize: 13, fontWeight: 500, cursor: disabled ? "not-allowed" : "pointer",
    borderRadius: 6, padding: "7px 12px", display: "inline-flex", alignItems: "center", gap: 6,
    border: "1px solid transparent", transition: "filter .15s", opacity: disabled ? 0.45 : 1,
  };
  const variants = {
    default: { background: theme.surface, color: theme.text, border: `1px solid ${theme.border}` },
    accent: { background: theme.accentDim, color: theme.accentText, border: `1px solid rgba(87,217,204,0.4)` },
    ghost: { background: "transparent", color: theme.textDim, border: `1px solid transparent` },
    danger: { background: theme.errorDim, color: theme.error, border: `1px solid rgba(241,113,121,0.35)` },
  };
  return (
    <button type="button" title={title} aria-label={ariaLabel || title} disabled={disabled} onClick={onClick}
      style={{ ...base, ...variants[variant], ...style }}
      onMouseEnter={(e) => { if (!disabled) e.currentTarget.style.filter = "brightness(1.18)"; }}
      onMouseLeave={(e) => { e.currentTarget.style.filter = "none"; }}
      onFocus={(e) => { e.currentTarget.style.boxShadow = focusRing; }}
      onBlur={(e) => { e.currentTarget.style.boxShadow = "none"; }}>
      {children}
    </button>
  );
}

function FieldLabel({ htmlFor, children }) {
  return <label htmlFor={htmlFor} style={{ fontFamily: fontBody, fontSize: 11.5, color: theme.textFaint, display: "block", marginBottom: 4 }}>{children}</label>;
}

function Select({ value, onChange, options, style, label }) {
  const id = useId();
  return (
    <div>
      {label && <FieldLabel htmlFor={id}>{label}</FieldLabel>}
      <select id={id} value={value} onChange={(e) => onChange(e.target.value)}
        style={{
          fontFamily: fontBody, fontSize: 13, background: theme.bg, color: theme.text,
          border: `1px solid ${theme.border}`, borderRadius: 6, padding: "6px 8px", outline: "none", ...style,
        }}
        onFocus={(e) => { e.currentTarget.style.boxShadow = focusRing; }}
        onBlur={(e) => { e.currentTarget.style.boxShadow = "none"; }}>
        {options.map((o) => (typeof o === "string" ? <option key={o} value={o}>{o}</option> : <option key={o.value} value={o.value}>{o.label}</option>))}
      </select>
    </div>
  );
}

function TextInput({ value, onChange, placeholder, style, label }) {
  const id = useId();
  return (
    <div>
      {label && <FieldLabel htmlFor={id}>{label}</FieldLabel>}
      <input id={id} value={value} onChange={(e) => onChange(e.target.value)} placeholder={placeholder}
        aria-label={!label ? placeholder : undefined}
        style={{
          fontFamily: fontBody, fontSize: 13, background: theme.bg, color: theme.text,
          border: `1px solid ${theme.border}`, borderRadius: 6, padding: "6px 8px", outline: "none", ...style,
        }}
        onFocus={(e) => { e.currentTarget.style.boxShadow = focusRing; }}
        onBlur={(e) => { e.currentTarget.style.boxShadow = "none"; }} />
    </div>
  );
}

function RangeInput({ value, onChange, min, max, step, label, id }) {
  const genId = useId();
  const inputId = id || genId;
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
      {label && <label htmlFor={inputId} style={{ fontFamily: fontBody, fontSize: 12, color: theme.textDim }}>{label}</label>}
      <input id={inputId} type="range" min={min} max={max} step={step} value={value}
        onChange={(e) => onChange(parseFloat(e.target.value))}
        aria-valuetext={`${Math.round(value * 100)} percent`}
        style={{ width: 110 }}
        onFocus={(e) => { e.currentTarget.style.boxShadow = focusRing; }}
        onBlur={(e) => { e.currentTarget.style.boxShadow = "none"; }} />
      <span aria-hidden="true" style={{ fontFamily: fontMono, fontSize: 12, color: theme.accentText, minWidth: 36 }}>{Math.round(value * 100)}%</span>
    </div>
  );
}

function CheckboxGroup({ columns, selected, onChange, legend }) {
  const instanceId = useId();
  const toggle = (c) => {
    if (selected.includes(c)) onChange(selected.filter((s) => s !== c));
    else onChange([...selected, c]);
  };
  return (
    <fieldset style={{ border: `1px solid ${theme.border}`, borderRadius: 6, padding: "8px 10px", margin: 0 }}>
      <legend style={{ fontFamily: fontBody, fontSize: 11.5, color: theme.textFaint, padding: "0 4px" }}>{legend}</legend>
      <div style={{ display: "flex", flexWrap: "wrap", gap: "6px 14px" }}>
        {columns.map((c) => {
          const id = `${instanceId}-${c}`;
          return (
            <label key={c} htmlFor={id} style={{ display: "flex", alignItems: "center", gap: 6, fontFamily: fontBody, fontSize: 12.5, color: theme.textDim, cursor: "pointer" }}>
              <input id={id} type="checkbox" checked={selected.includes(c)} onChange={() => toggle(c)}
                style={{ accentColor: theme.accent, width: 14, height: 14 }}
                onFocus={(e) => { e.currentTarget.style.boxShadow = focusRing; }}
                onBlur={(e) => { e.currentTarget.style.boxShadow = "none"; }} />
              {c}
            </label>
          );
        })}
      </div>
    </fieldset>
  );
}

const DELIMITER_PRESETS = [
  { value: " ", label: "Space" },
  { value: ",", label: "Comma (,)" },
  { value: "-", label: "Hyphen (-)" },
  { value: "_", label: "Underscore (_)" },
  { value: "|", label: "Pipe (|)" },
  { value: "\t", label: "Tab" },
  { value: "__custom__", label: "Custom..." },
];

function DelimiterInput({ value, onChange, label = "Delimiter" }) {
  const isKnownPreset = DELIMITER_PRESETS.some((p) => p.value === value);
  const [mode, setMode] = useState(isKnownPreset ? value : "__custom__");
  const handlePresetChange = (v) => {
    setMode(v);
    if (v !== "__custom__") onChange(v);
    else onChange("");
  };
  return (
    <div style={{ display: "flex", gap: 6, alignItems: "flex-end" }}>
      <Select label={label} value={mode} onChange={handlePresetChange} options={DELIMITER_PRESETS} />
      {mode === "__custom__" && (
        <TextInput label="Custom character(s)" value={value} onChange={onChange} placeholder="e.g. ;" style={{ width: 90 }} />
      )}
    </div>
  );
}

function TypeBadge({ type }) {
  const color = TYPE_COLORS[type] || theme.textDim;
  return <span style={{ fontFamily: fontMono, fontSize: 10, letterSpacing: 0.4, textTransform: "uppercase", color, border: `1px solid ${color}55`, background: `${color}18`, borderRadius: 4, padding: "1px 6px" }}>{type}</span>;
}

// ---------- Data table ----------
function DataTable({ data, columns, maxRows = 60, caption }) {
  const [sortCol, setSortCol] = useState(null);
  const [sortDir, setSortDir] = useState(1);
  const [filter, setFilter] = useState("");
  const filterId = useId();

  const types = useMemo(() => { const t = {}; columns.forEach((c) => { t[c] = detectType(data.map((r) => r[c])); }); return t; }, [data, columns]);
  const filtered = useMemo(() => {
    if (!filter.trim()) return data;
    const f = filter.toLowerCase();
    return data.filter((row) => columns.some((c) => String(row[c] ?? "").toLowerCase().includes(f)));
  }, [data, columns, filter]);
  const sorted = useMemo(() => {
    if (!sortCol) return filtered;
    return [...filtered].sort((a, b) => {
      const av = a[sortCol], bv = b[sortCol];
      const an = parseFloat(av), bn = parseFloat(bv);
      const cmp = (!isNaN(an) && !isNaN(bn)) ? an - bn : String(av ?? "").localeCompare(String(bv ?? ""));
      return cmp * sortDir;
    });
  }, [filtered, sortCol, sortDir]);

  if (!columns.length) return <div style={{ color: theme.textFaint, fontFamily: fontBody, fontSize: 13, padding: 24, textAlign: "center" }}>No data loaded yet</div>;

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
        <Search size={14} color={theme.textFaint} aria-hidden="true" />
        <TextInput value={filter} onChange={setFilter} placeholder="Filter rows..." style={{ width: 220 }} />
        <div aria-live="polite" style={{ marginLeft: "auto", fontFamily: fontMono, fontSize: 12, color: theme.textDim }}>
          {sorted.length.toLocaleString()} rows · {columns.length} columns
        </div>
      </div>
      <div style={{ overflowX: "auto", border: `1px solid ${theme.border}`, borderRadius: 8 }}>
        <table style={{ borderCollapse: "collapse", width: "100%", fontFamily: fontMono, fontSize: 12.5 }}>
          {caption && <caption style={{ textAlign: "left", padding: "8px 10px", fontFamily: fontBody, fontSize: 12, color: theme.textFaint }}>{caption}</caption>}
          <thead>
            <tr style={{ background: theme.panel }}>
              {columns.map((c) => {
                const active = sortCol === c;
                return (
                  <th key={c} scope="col" aria-sort={active ? (sortDir === 1 ? "ascending" : "descending") : "none"}>
                    <button type="button" onClick={() => { setSortCol(c); setSortDir(active ? -sortDir : 1); }}
                      style={{
                        all: "unset", cursor: "pointer", display: "flex", flexDirection: "column", gap: 4,
                        padding: "8px 10px", whiteSpace: "nowrap", borderBottom: `1px solid ${theme.border}`, width: "100%", boxSizing: "border-box",
                      }}
                      onFocus={(e) => { e.currentTarget.style.boxShadow = `inset ${focusRing}`; }}
                      onBlur={(e) => { e.currentTarget.style.boxShadow = "none"; }}
                      aria-label={`Sort by ${c}${active ? (sortDir === 1 ? ", ascending" : ", descending") : ""}`}>
                      <span style={{ display: "flex", alignItems: "center", gap: 5 }}>
                        <span style={{ color: theme.text, fontFamily: fontBody, fontWeight: 500 }}>{c}</span>
                        {active ? (sortDir === 1 ? <ChevronUp size={12} color={theme.accent} aria-hidden="true" /> : <ChevronDown size={12} color={theme.accent} aria-hidden="true" />) : <ArrowUpDown size={11} color={theme.textFaint} aria-hidden="true" />}
                      </span>
                      <TypeBadge type={types[c]} />
                    </button>
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {sorted.slice(0, maxRows).map((row, i) => (
              <tr key={i} style={{ background: i % 2 ? "transparent" : "rgba(255,255,255,0.02)" }}>
                {columns.map((c) => (
                  <td key={c} style={{ padding: "6px 10px", color: theme.textDim, borderBottom: `1px solid ${theme.border}`, whiteSpace: "nowrap", maxWidth: 220, overflow: "hidden", textOverflow: "ellipsis" }}>
                    {String(row[c] ?? "")}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {sorted.length > maxRows && <div style={{ fontFamily: fontBody, fontSize: 12, color: theme.textFaint, marginTop: 8, textAlign: "center" }}>Showing first {maxRows} of {sorted.length.toLocaleString()} rows</div>}
    </div>
  );
}

// ---------- Step editor ----------
function StepEditor({ step, columns, sampleRows = [], onChange }) {
  const p = step.params;
  const set = (k, v) => onChange({ ...step, params: { ...p, [k]: v } });
  switch (step.type) {
    case "case":
      return (
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <Select label="Column" value={p.column} onChange={(v) => set("column", v)} options={columns} />
          <Select label="Case" value={p.mode} onChange={(v) => set("mode", v)} options={[{ value: "title", label: "Title Case" }, { value: "upper", label: "UPPER CASE" }, { value: "lower", label: "lower case" }]} />
        </div>
      );
    case "dedupe_fuzzy":
      return (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          <CheckboxGroup legend="Match on column(s)" columns={columns} selected={p.columns} onChange={(v) => set("columns", v)} />
          <RangeInput label="Match sensitivity" value={p.threshold} onChange={(v) => set("threshold", v)} min="0.5" max="1" step="0.01" />
          {(!p.columns || !p.columns.length) && (
            <div style={{ display: "flex", gap: 5, alignItems: "center", fontFamily: fontBody, fontSize: 11.5, color: theme.warn }}>
              <AlertTriangle size={12} aria-hidden="true" /> Select at least one column, or nothing will be matched.
            </div>
          )}
        </div>
      );
    case "split_column": {
      const names = p.names && p.names.length ? p.names : ["", ""];
      const setNames = (next) => set("names", next);
      const updateName = (i, v) => { const next = [...names]; next[i] = v; setNames(next); };
      const addColumn = () => setNames([...names, ""]);
      const removeColumn = (i) => { if (names.length > 2) setNames(names.filter((_, idx) => idx !== i)); };
      const sampleVal = (sampleRows.find((r) => r[p.column] && String(r[p.column]).trim() !== "") || {})[p.column];
      const delim = p.delimiter === "" ? " " : p.delimiter;
      const previewParts = sampleVal ? String(sampleVal).split(delim) : null;
      return (
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <Select label="Column" value={p.column} onChange={(v) => set("column", v)} options={columns} />
            <DelimiterInput value={p.delimiter} onChange={(v) => set("delimiter", v)} />
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            <span style={{ fontFamily: fontBody, fontSize: 11.5, color: theme.textFaint }}>New columns (in order)</span>
            {names.map((n, i) => (
              <div key={i} style={{ display: "flex", gap: 6, alignItems: "center" }}>
                <TextInput value={n} onChange={(v) => updateName(i, v)} placeholder={`${p.column || "column"}_${i + 1}${i === names.length - 1 && names.length > 1 ? " (remainder)" : ""}`} style={{ width: 200 }} />
                {names.length > 2 && (
                  <Btn variant="ghost" onClick={() => removeColumn(i)} style={{ padding: 5 }} ariaLabel={`Remove new column ${i + 1}`}><X size={13} /></Btn>
                )}
              </div>
            ))}
            <Btn variant="ghost" onClick={addColumn} style={{ fontSize: 11.5, alignSelf: "flex-start" }}><Plus size={12} aria-hidden="true" /> Add another column</Btn>
          </div>
          <div aria-live="polite" style={{ fontFamily: fontMono, fontSize: 11.5, color: previewParts ? theme.accentText : theme.textFaint, background: theme.bg, border: `1px solid ${theme.border}`, borderRadius: 6, padding: "6px 8px" }}>
            {sampleVal
              ? `Preview: "${sampleVal}" → ${names.map((n, i) => `"${(i < names.length - 1 ? (previewParts[i] ?? "") : previewParts.slice(i).join(delim)).trim()}"`).join(" | ")}`
              : "No sample value found in this column yet — upload data or pick a different column."}
          </div>
        </div>
      );
    }
    case "merge_columns": {
      const sampleRow = sampleRows.find((r) => (r[p.col1] || r[p.col2]) && String(r[p.col1] ?? r[p.col2]).trim() !== "");
      const sep = p.separator === "" ? " " : p.separator;
      return (
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <Select label="First column" value={p.col1} onChange={(v) => set("col1", v)} options={columns} />
            <Select label="Second column" value={p.col2} onChange={(v) => set("col2", v)} options={columns} />
            <DelimiterInput label="Separator" value={p.separator} onChange={(v) => set("separator", v)} />
            <TextInput label="New column name" value={p.name} onChange={(v) => set("name", v)} placeholder={`${p.col1}_${p.col2}`} style={{ width: 130 }} />
          </div>
          <div aria-live="polite" style={{ fontFamily: fontMono, fontSize: 11.5, color: sampleRow ? theme.accentText : theme.textFaint, background: theme.bg, border: `1px solid ${theme.border}`, borderRadius: 6, padding: "6px 8px" }}>
            {sampleRow ? `Preview: "${sampleRow[p.col1] ?? ""}" + "${sampleRow[p.col2] ?? ""}" → "${sampleRow[p.col1] ?? ""}${sep}${sampleRow[p.col2] ?? ""}"` : "No sample row found yet — upload data first."}
          </div>
        </div>
      );
    }
    case "fill_missing":
      return (
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "flex-end" }}>
          <Select label="Column" value={p.column} onChange={(v) => set("column", v)} options={columns} />
          <Select label="Fill with" value={p.strategy} onChange={(v) => set("strategy", v)} options={[{ value: "value", label: "Custom value" }, { value: "mean", label: "Mean" }, { value: "median", label: "Median" }, { value: "mode", label: "Mode" }]} />
          {p.strategy === "value" && <TextInput label="Value" value={p.customValue} onChange={(v) => set("customValue", v)} style={{ width: 100 }} />}
        </div>
      );
    case "remove_column":
      return <Select label="Column to remove" value={p.column} onChange={(v) => set("column", v)} options={columns} />;
    case "keep_columns":
      return <CheckboxGroup legend="Columns to keep" columns={columns} selected={p.columns} onChange={(v) => set("columns", v)} />;
    case "filter_rows": {
      const needsValue = ["equals", "not_equals", "contains", "greater_than", "less_than"].includes(p.operator);
      return (
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "flex-end" }}>
          <Select label="Column" value={p.column} onChange={(v) => set("column", v)} options={columns} />
          <Select label="Condition" value={p.operator} onChange={(v) => set("operator", v)} options={[
            { value: "is_blank", label: "is blank" },
            { value: "is_not_blank", label: "is not blank" },
            { value: "is_zero", label: "is zero" },
            { value: "equals", label: "equals" },
            { value: "not_equals", label: "does not equal" },
            { value: "contains", label: "contains" },
            { value: "greater_than", label: "greater than" },
            { value: "less_than", label: "less than" },
          ]} />
          {needsValue && <TextInput label="Value" value={p.value} onChange={(v) => set("value", v)} style={{ width: 110 }} />}
          <Select label="Action" value={p.action} onChange={(v) => set("action", v)} options={[{ value: "remove", label: "Remove matching rows" }, { value: "keep", label: "Keep only matching rows" }]} />
        </div>
      );
    }
    case "find_replace":
      return (
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <Select label="Column" value={p.column} onChange={(v) => set("column", v)} options={columns} />
          <TextInput label="Find" value={p.find} onChange={(v) => set("find", v)} placeholder="text to find" style={{ width: 130 }} />
          <TextInput label="Replace with" value={p.replace} onChange={(v) => set("replace", v)} placeholder="replacement" style={{ width: 130 }} />
        </div>
      );
    case "clean_number": {
      const sampleVal = (sampleRows.find((r) => r[p.column] && String(r[p.column]).trim() !== "") || {})[p.column];
      const cleaned = sampleVal !== undefined ? parseFloat(String(sampleVal).replace(/[^0-9.\-]/g, "")) : null;
      return (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          <Select label="Column" value={p.column} onChange={(v) => set("column", v)} options={columns} />
          <div aria-live="polite" style={{ fontFamily: fontMono, fontSize: 11.5, color: cleaned !== null && !isNaN(cleaned) ? theme.accentText : theme.textFaint, background: theme.bg, border: `1px solid ${theme.border}`, borderRadius: 6, padding: "6px 8px" }}>
            {sampleVal !== undefined ? `Preview: "${sampleVal}" → ${isNaN(cleaned) ? "(not numeric)" : cleaned}` : "No sample value found yet."}
          </div>
        </div>
      );
    }
    case "rename_column":
      return (
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <Select label="Column" value={p.column} onChange={(v) => set("column", v)} options={columns} />
          <TextInput label="New name" value={p.newName} onChange={(v) => set("newName", v)} placeholder={p.column} style={{ width: 160 }} />
        </div>
      );
    case "sort_rows":
      return (
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <Select label="Column" value={p.column} onChange={(v) => set("column", v)} options={columns} />
          <Select label="Direction" value={p.direction} onChange={(v) => set("direction", v)} options={[{ value: "asc", label: "Ascending (A→Z, 0→9)" }, { value: "desc", label: "Descending (Z→A, 9→0)" }]} />
        </div>
      );
    default:
      return <span style={{ fontFamily: fontBody, fontSize: 12, color: theme.textFaint }}>No configuration needed</span>;
  }
}

// ---------- Pipeline tab ----------
function PipelineTab({ file, setFile }) {
  const [steps, setSteps] = useState([]);
  const fileInputRef = useRef(null);
  const pipelineFileRef = useRef(null);

  const handleUpload = async (e) => {
    const f = e.target.files[0];
    if (!f) return;
    try {
      const { data, columns } = await parseFile(f);
      setFile({ name: f.name, data, columns });
      setSteps([]);
    } catch { alert("Could not read that file. Make sure it's a valid CSV or Excel file."); }
  };

  const columnsAtStep = useMemo(() => {
    if (!file) return [];
    const list = [file.columns];
    let cols = file.columns, data = file.data;
    for (const step of steps) {
      const res = applyStep(data, step);
      data = res.data;
      cols = data.length ? Object.keys(data[0]) : cols;
      list.push(cols);
    }
    return list;
  }, [file, steps]);

  const dataAtStep = useMemo(() => {
    if (!file) return [];
    const list = [file.data];
    let data = file.data;
    for (const step of steps) {
      const res = applyStep(data, step);
      data = res.data;
      list.push(data);
    }
    return list;
  }, [file, steps]);

  const pipelineResult = useMemo(() => {
    if (!file) return { data: [], results: [] };
    let data = file.data;
    const results = [];
    for (const step of steps) { const res = applyStep(data, step); data = res.data; results.push(res.meta); }
    return { data, results };
  }, [file, steps]);

  const addStep = (type) => {
    const cols = columnsAtStep[columnsAtStep.length - 1] || [];
    setSteps([...steps, { id: Date.now() + Math.random(), type, params: defaultParams(type, cols) }]);
  };
  const updateStep = (idx, newStep) => { const next = [...steps]; next[idx] = newStep; setSteps(next); };
  const removeStep = (idx) => setSteps(steps.filter((_, i) => i !== idx));
  const moveStep = (idx, dir) => {
    const next = [...steps]; const target = idx + dir;
    if (target < 0 || target >= next.length) return;
    [next[idx], next[target]] = [next[target], next[idx]];
    setSteps(next);
  };

  const loadTemplate = (tpl) => {
    if (!file) return;
    const cols = file.columns;
    if (tpl === "crm") {
      setSteps([
        { id: Date.now() + 1, type: "trim", params: {} },
        { id: Date.now() + 2, type: "case", params: { column: cols[0] || "", mode: "title" } },
        { id: Date.now() + 3, type: "dedupe_fuzzy", params: { columns: cols[0] ? [cols[0]] : [], threshold: 0.85 } },
      ]);
    }
  };

  const savePipeline = () => {
    const savable = steps.map(({ id, type, params }) => ({ id, type, params }));
    downloadBlob(JSON.stringify(savable, null, 2), "pipeline_template.json", "application/json");
  };
  const loadPipeline = (e) => {
    const f = e.target.files[0];
    if (!f) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const loaded = JSON.parse(ev.target.result);
        setSteps(loaded);
      } catch { alert("That file doesn't look like a saved pipeline (expected JSON)."); }
    };
    reader.readAsText(f);
  };

  return (
    <div style={{ display: "grid", gridTemplateColumns: "360px 1fr", gap: 20 }}>
      <div>
        <section aria-labelledby="source-heading" style={{ border: `1px solid ${theme.border}`, borderRadius: 8, background: theme.panel, padding: 14 }}>
          <h2 id="source-heading" style={{ fontFamily: fontDisplay, fontSize: 14, fontWeight: 600, color: theme.text, margin: "0 0 10px" }}>Data source</h2>
          <input ref={fileInputRef} type="file" accept=".csv,.xlsx,.xls" onChange={handleUpload} style={{ display: "none" }} aria-hidden="true" tabIndex={-1} />
          <Btn variant="accent" onClick={() => fileInputRef.current.click()} style={{ width: "100%", justifyContent: "center" }}>
            <Upload size={14} aria-hidden="true" /> {file ? "Replace file" : "Upload CSV or Excel"}
          </Btn>
          {file && (
            <div style={{ marginTop: 10, fontFamily: fontMono, fontSize: 11.5, color: theme.textDim, display: "flex", alignItems: "center", gap: 6 }}>
              <FileSpreadsheet size={13} color={theme.accent} aria-hidden="true" /> {file.name} · {file.data.length.toLocaleString()} rows
            </div>
          )}
          {file && (
            <div style={{ marginTop: 10 }}>
              <Btn variant="ghost" onClick={() => loadTemplate("crm")} style={{ fontSize: 11.5 }}>Use CRM dedupe template</Btn>
            </div>
          )}
        </section>

        <section aria-labelledby="pipeline-heading" style={{ marginTop: 16 }}>
          <div style={{ display: "flex", alignItems: "center", marginBottom: 10 }}>
            <h2 id="pipeline-heading" style={{ fontFamily: fontDisplay, fontSize: 14, fontWeight: 600, color: theme.text, margin: 0 }}>Pipeline</h2>
            <div style={{ marginLeft: "auto", display: "flex", gap: 6 }}>
              <input ref={pipelineFileRef} type="file" accept=".json" onChange={loadPipeline} style={{ display: "none" }} aria-hidden="true" tabIndex={-1} />
              <Btn variant="ghost" title="Load a saved pipeline template" onClick={() => pipelineFileRef.current.click()} style={{ fontSize: 11.5 }}><FolderOpen size={13} aria-hidden="true" /> Load</Btn>
              <Btn variant="ghost" title="Save this pipeline as a reusable template" disabled={!steps.length} onClick={savePipeline} style={{ fontSize: 11.5 }}><Save size={13} aria-hidden="true" /> Save</Btn>
            </div>
          </div>

          {steps.length === 0 && <div style={{ fontFamily: fontBody, fontSize: 12.5, color: theme.textFaint, padding: "14px 0" }}>No steps yet. Add one below to start cleaning your data.</div>}

          <ol style={{ display: "flex", flexDirection: "column", gap: 0, listStyle: "none", margin: 0, padding: 0 }}>
            {steps.map((step, idx) => {
              const label = STEP_LIBRARY.find((s) => s.type === step.type)?.label || step.type;
              const meta = pipelineResult.results[idx];
              const delta = meta ? meta.after - meta.before : 0;
              return (
                <li key={step.id}>
                  {idx > 0 && <div aria-hidden="true" style={{ width: 1, height: 14, background: theme.borderLight, marginLeft: 15 }} />}
                  <div style={{ border: `1px solid ${theme.border}`, borderRadius: 8, background: theme.surface, padding: 12 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                      <div aria-hidden="true" style={{ width: 22, height: 22, borderRadius: "50%", background: theme.accentDim, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: fontMono, fontSize: 11, color: theme.accentText, flexShrink: 0 }}>{idx + 1}</div>
                      <span style={{ fontFamily: fontBody, fontSize: 13, fontWeight: 500, color: theme.text, flex: 1 }}>{label}</span>
                      <Btn variant="ghost" onClick={() => moveStep(idx, -1)} disabled={idx === 0} style={{ padding: 5 }} ariaLabel={`Move ${label} step up`}><ChevronUp size={13} /></Btn>
                      <Btn variant="ghost" onClick={() => moveStep(idx, 1)} disabled={idx === steps.length - 1} style={{ padding: 5 }} ariaLabel={`Move ${label} step down`}><ChevronDown size={13} /></Btn>
                      <Btn variant="ghost" onClick={() => removeStep(idx)} style={{ padding: 5 }} ariaLabel={`Remove ${label} step`}><X size={13} /></Btn>
                    </div>
                    <StepEditor step={step} columns={columnsAtStep[idx] || []} sampleRows={dataAtStep[idx] || []} onChange={(ns) => updateStep(idx, ns)} />
                    {meta && (
                      <div aria-live="polite" style={{ marginTop: 8, fontFamily: fontMono, fontSize: 11, color: delta !== 0 ? theme.warn : theme.textFaint, display: "flex", alignItems: "center", gap: 6 }}>
                        <span>{meta.detail}</span>
                        {delta !== 0 && <span>({delta > 0 ? "+" : ""}{delta} rows)</span>}
                      </div>
                    )}
                  </div>
                </li>
              );
            })}
          </ol>

          <div style={{ marginTop: 12 }}>
            <h3 style={{ fontFamily: fontBody, fontSize: 11.5, color: theme.textFaint, margin: "0 0 6px", textTransform: "uppercase", letterSpacing: 0.5, fontWeight: 500 }}>Add step</h3>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
              {STEP_LIBRARY.map((s) => (
                <Btn key={s.type} variant="default" disabled={!file} onClick={() => addStep(s.type)} style={{ fontSize: 12 }}>
                  <Plus size={12} aria-hidden="true" /> {s.label}
                </Btn>
              ))}
            </div>
          </div>
        </section>
      </div>

      <div>
        <div style={{ display: "flex", alignItems: "center", marginBottom: 12 }}>
          <h2 style={{ fontFamily: fontDisplay, fontSize: 14, fontWeight: 600, color: theme.text, margin: 0 }}>Result preview</h2>
          <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
            <Btn variant="default" disabled={!file} onClick={() => exportXLSX(pipelineResult.data, `cleaned_${file?.name?.replace(/\.[^.]+$/, "") || "data"}.xlsx`)}>
              <Download size={13} aria-hidden="true" /> Export as Excel
            </Btn>
            <Btn variant="accent" disabled={!file} onClick={() => exportCSV(pipelineResult.data, `cleaned_${file?.name?.replace(/\.[^.]+$/, "") || "data"}.csv`)}>
              <Download size={13} aria-hidden="true" /> Export cleaned CSV
            </Btn>
          </div>
        </div>
        <DataTable data={pipelineResult.data} columns={pipelineResult.data.length ? Object.keys(pipelineResult.data[0]) : (file?.columns || [])} caption="Cleaned data preview, reflects all pipeline steps applied in order" />
      </div>
    </div>
  );
}

// ---------- Compare tab ----------
function CompareTab() {
  const [fileA, setFileA] = useState(null);
  const [fileB, setFileB] = useState(null);
  const [keyA, setKeyA] = useState("");
  const [keyB, setKeyB] = useState("");
  const [valueA, setValueA] = useState("");
  const [valueB, setValueB] = useState("");
  const [threshold, setThreshold] = useState(0.85);
  const [ran, setRan] = useState(false);
  const refA = useRef(null), refB = useRef(null);

  const upload = async (e, setter, setKey, setVal) => {
    const f = e.target.files[0];
    if (!f) return;
    try {
      const { data, columns } = await parseFile(f);
      setter({ name: f.name, data, columns });
      setKey(columns[0] || ""); setVal(columns[1] || columns[0] || "");
      setRan(false);
    } catch { alert("Could not read that file."); }
  };

  const results = useMemo(() => {
    if (!ran || !fileA || !fileB) return [];
    const usedB = new Array(fileB.data.length).fill(false);
    const out = [];
    fileA.data.forEach((rowA) => {
      let bestIdx = -1, bestSim = 0;
      fileB.data.forEach((rowB, j) => {
        if (usedB[j]) return;
        const sim = similarity(rowA[keyA], rowB[keyB]);
        if (sim > bestSim) { bestSim = sim; bestIdx = j; }
      });
      if (bestIdx >= 0 && bestSim >= threshold) {
        usedB[bestIdx] = true;
        const rowB = fileB.data[bestIdx];
        const vA = parseFloat(String(rowA[valueA]).replace(/[^0-9.\-]/g, ""));
        const vB = parseFloat(String(rowB[valueB]).replace(/[^0-9.\-]/g, ""));
        const delta = !isNaN(vA) && !isNaN(vB) ? (vB - vA) : null;
        out.push({ keyA: rowA[keyA], keyB: rowB[keyB], match: Math.round(bestSim * 100), valA: rowA[valueA], valB: rowB[valueB], delta, status: "matched" });
      } else {
        out.push({ keyA: rowA[keyA], keyB: "—", match: 0, valA: rowA[valueA], valB: "—", delta: null, status: "unmatched" });
      }
    });
    return out;
  }, [ran, fileA, fileB, keyA, keyB, valueA, valueB, threshold]);

  const matchedCount = results.filter((r) => r.status === "matched").length;

  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 16 }}>
        {[{ file: fileA, setter: setFileA, keyCol: keyA, setKeyCol: setKeyA, valCol: valueA, setValCol: setValueA, ref: refA, label: "File A" },
          { file: fileB, setter: setFileB, keyCol: keyB, setKeyCol: setKeyB, valCol: valueB, setValCol: setValueB, ref: refB, label: "File B" }].map((c, i) => (
          <section aria-labelledby={`file-${i}-heading`} key={i} style={{ border: `1px solid ${theme.border}`, borderRadius: 8, background: theme.panel, padding: 14 }}>
            <h2 id={`file-${i}-heading`} style={{ fontFamily: fontDisplay, fontSize: 13, fontWeight: 600, color: theme.text, margin: "0 0 8px" }}>{c.label}</h2>
            <input ref={c.ref} type="file" accept=".csv,.xlsx,.xls" style={{ display: "none" }} aria-hidden="true" tabIndex={-1} onChange={(e) => upload(e, c.setter, c.setKeyCol, c.setValCol)} />
            <Btn variant="accent" onClick={() => c.ref.current.click()} style={{ width: "100%", justifyContent: "center" }}>
              <Upload size={13} aria-hidden="true" /> {c.file ? "Replace file" : "Upload file"}
            </Btn>
            {c.file && (
              <>
                <div style={{ fontFamily: fontMono, fontSize: 11, color: theme.textDim, margin: "8px 0" }}>{c.file.name} · {c.file.data.length} rows</div>
                <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                  <Select label="Match key column" value={c.keyCol} onChange={c.setKeyCol} options={c.file.columns} />
                  <Select label="Value column to compare" value={c.valCol} onChange={c.setValCol} options={c.file.columns} />
                </div>
              </>
            )}
          </section>
        ))}
      </div>

      <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 16, flexWrap: "wrap" }}>
        <RangeInput label="Match sensitivity" value={threshold} onChange={setThreshold} min="0.5" max="1" step="0.01" />
        <Btn variant="accent" disabled={!fileA || !fileB} onClick={() => setRan(true)}>
          <Play size={13} aria-hidden="true" /> Run comparison
        </Btn>
        {ran && <Btn variant="default" onClick={() => exportCSV(results, "comparison_result.csv")}><Download size={13} aria-hidden="true" /> Export result</Btn>}
      </div>

      {ran && (
        <>
          <div aria-live="polite" style={{ display: "flex", gap: 12, marginBottom: 14 }}>
            <div style={{ background: theme.successDim, color: theme.success, fontFamily: fontMono, fontSize: 12, padding: "6px 12px", borderRadius: 6 }}>{matchedCount} matched</div>
            <div style={{ background: theme.errorDim, color: theme.error, fontFamily: fontMono, fontSize: 12, padding: "6px 12px", borderRadius: 6 }}>{results.length - matchedCount} unmatched</div>
          </div>
          <div style={{ overflowX: "auto", border: `1px solid ${theme.border}`, borderRadius: 8 }}>
            <table style={{ borderCollapse: "collapse", width: "100%", fontFamily: fontMono, fontSize: 12.5 }}>
              <caption style={{ textAlign: "left", padding: "8px 10px", fontFamily: fontBody, fontSize: 12, color: theme.textFaint }}>Row-by-row comparison between file A and file B</caption>
              <thead>
                <tr style={{ background: theme.panel }}>
                  {["Status", "Key (A)", "Key (B)", "Match %", "Value (A)", "Value (B)", "Delta"].map((h) => (
                    <th key={h} scope="col" style={{ textAlign: "left", padding: "8px 10px", borderBottom: `1px solid ${theme.border}`, color: theme.text, fontFamily: fontBody, fontWeight: 500 }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {results.slice(0, 80).map((r, i) => (
                  <tr key={i}>
                    <td style={{ padding: "6px 10px", borderBottom: `1px solid ${theme.border}` }}>
                      {r.status === "matched" ? <><CheckCircle2 size={14} color={theme.success} aria-hidden="true" /><span className="sr-only" style={{ position: "absolute", width: 1, height: 1, overflow: "hidden" }}>Matched</span></> : <><AlertTriangle size={14} color={theme.error} aria-hidden="true" /><span style={{ position: "absolute", width: 1, height: 1, overflow: "hidden" }}>Unmatched</span></>}
                    </td>
                    <td style={{ padding: "6px 10px", borderBottom: `1px solid ${theme.border}`, color: theme.textDim }}>{String(r.keyA)}</td>
                    <td style={{ padding: "6px 10px", borderBottom: `1px solid ${theme.border}`, color: theme.textDim }}>{String(r.keyB)}</td>
                    <td style={{ padding: "6px 10px", borderBottom: `1px solid ${theme.border}`, color: theme.textFaint }}>{r.match}%</td>
                    <td style={{ padding: "6px 10px", borderBottom: `1px solid ${theme.border}`, color: theme.textDim }}>{String(r.valA)}</td>
                    <td style={{ padding: "6px 10px", borderBottom: `1px solid ${theme.border}`, color: theme.textDim }}>{String(r.valB)}</td>
                    <td style={{ padding: "6px 10px", borderBottom: `1px solid ${theme.border}`, color: r.delta > 0 ? theme.warn : r.delta < 0 ? theme.success : theme.textFaint }}>{r.delta === null ? "—" : (r.delta > 0 ? "+" : "") + r.delta.toFixed(2)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      )}
    </div>
  );
}

// ---------- Explorer tab ----------
function ExplorerTab({ file }) {
  if (!file) return <div style={{ color: theme.textFaint, fontFamily: fontBody, fontSize: 13, padding: 40, textAlign: "center" }}>Upload a file in the Pipeline tab first</div>;
  const stats = file.columns.map((c) => ({ col: c, ...columnStats(file.data, c), type: detectType(file.data.map((r) => r[c])) }));
  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(160px, 1fr))", gap: 10, marginBottom: 18 }}>
        {stats.map((s) => (
          <div key={s.col} style={{ border: `1px solid ${theme.border}`, borderRadius: 8, background: theme.panel, padding: 10 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
              <span style={{ fontFamily: fontBody, fontSize: 12.5, color: theme.text, fontWeight: 500, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{s.col}</span>
              <TypeBadge type={s.type} />
            </div>
            <div style={{ fontFamily: fontMono, fontSize: 11, color: theme.textFaint }}>{s.unique.toLocaleString()} unique · {s.nulls > 0 ? <span style={{ color: theme.warn }}>{s.nulls} blank</span> : "0 blank"}</div>
          </div>
        ))}
      </div>
      <DataTable data={file.data} columns={file.columns} maxRows={100} caption="Full uploaded dataset, unmodified" />
    </div>
  );
}

export default function DataWorkbench() {
  const [tab, setTab] = useState("pipeline");
  const [file, setFile] = useState(null);
  const tabs = [
    { id: "pipeline", label: "Pipeline", icon: GitBranch },
    { id: "compare", label: "Compare", icon: ArrowRightLeft },
    { id: "explorer", label: "Explorer", icon: Table2 },
  ];

  return (
    <div style={{ background: theme.bg, minHeight: "100vh", padding: "28px 32px", fontFamily: fontBody }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');
        input[type=range] { accent-color: ${theme.accent}; }
        ::-webkit-scrollbar { height: 8px; width: 8px; }
        ::-webkit-scrollbar-thumb { background: ${theme.border}; border-radius: 4px; }
        @media (prefers-reduced-motion: reduce) { * { transition: none !important; } }
        a.skip-link:focus { position: fixed; top: 8px; left: 8px; z-index: 999; }
      `}</style>

      <a href="#main-content" className="skip-link" style={{
        position: "absolute", left: -9999, top: "auto", background: theme.accent, color: theme.bg,
        padding: "8px 14px", borderRadius: 6, fontFamily: fontBody, fontSize: 13, fontWeight: 600, zIndex: 999,
      }} onFocus={(e) => { e.currentTarget.style.left = "8px"; e.currentTarget.style.top = "8px"; e.currentTarget.style.position = "fixed"; }}>
        Skip to main content
      </a>

      <div style={{ display: "flex", alignItems: "baseline", gap: 12, marginBottom: 4 }}>
        <Database size={20} color={theme.accent} aria-hidden="true" />
        <h1 style={{ fontFamily: fontDisplay, fontSize: 22, fontWeight: 600, color: theme.text, margin: 0 }}>Data Workbench</h1>
      </div>
      <p style={{ fontFamily: fontBody, fontSize: 13.5, color: theme.textDim, margin: "4px 0 22px", maxWidth: 560 }}>
        Clean, dedupe, and compare spreadsheets entirely in your browser. Nothing is uploaded anywhere — your data never leaves this device.
      </p>

      <div role="tablist" aria-label="Data Workbench sections" style={{ display: "flex", gap: 4, borderBottom: `1px solid ${theme.border}`, marginBottom: 22 }}>
        {tabs.map((t) => {
          const Icon = t.icon;
          const active = tab === t.id;
          return (
            <button key={t.id} role="tab" aria-selected={active} aria-controls={`panel-${t.id}`} id={`tab-${t.id}`}
              onClick={() => setTab(t.id)}
              style={{
                display: "flex", alignItems: "center", gap: 6, padding: "10px 14px", background: "transparent",
                border: "none", borderBottom: active ? `2px solid ${theme.accent}` : "2px solid transparent",
                color: active ? theme.text : theme.textFaint, fontFamily: fontBody, fontSize: 13.5,
                fontWeight: 500, cursor: "pointer", marginBottom: -1,
              }}
              onFocus={(e) => { e.currentTarget.style.boxShadow = focusRing; }}
              onBlur={(e) => { e.currentTarget.style.boxShadow = "none"; }}>
              <Icon size={15} aria-hidden="true" /> {t.label}
            </button>
          );
        })}
      </div>

      <main id="main-content">
        <div role="tabpanel" id="panel-pipeline" aria-labelledby="tab-pipeline" hidden={tab !== "pipeline"}>
          {tab === "pipeline" && <PipelineTab file={file} setFile={setFile} />}
        </div>
        <div role="tabpanel" id="panel-compare" aria-labelledby="tab-compare" hidden={tab !== "compare"}>
          {tab === "compare" && <CompareTab />}
        </div>
        <div role="tabpanel" id="panel-explorer" aria-labelledby="tab-explorer" hidden={tab !== "explorer"}>
          {tab === "explorer" && <ExplorerTab file={file} />}
        </div>
      </main>
    </div>
  );
}
