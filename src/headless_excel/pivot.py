"""Fluent pivot-table construction for headless-excel.

Pivot tables are *declared* here (layout, fields, aggregation functions, style)
but never *computed*. The structure is written into the workbook with the source
cache flagged ``refreshOnLoad`` so that LibreOffice (during ``ctx.sync()``) — or
Excel, when a user opens the file — materialises every value and grand total.

The API mirrors ClosedXML / EPPlus / SpreadJS:

    pt = (
        ws.add_pivot(source="Data!A1:D7", anchor="A3", name="SalesByRegion")
          .rows("Region")
          .columns("Product")
          .values("Sales", func="sum", name="Total Sales", num_format="$#,##0")
          .filters("Year")
          .style("PivotStyleMedium9", row_stripes=True)
    )

A one-shot declarative form is equivalent::

    ws.add_pivot(
        source="Data!A1:D7", anchor="A3", name="SalesByRegion",
        rows=["Region"], columns=["Product"],
        values=[("Sales", "sum", "Total Sales", "$#,##0")],
        filters=["Year"], style="PivotStyleMedium9",
    )

Aggregation functions accepted (Excel ``subtotal`` values):
``sum, count, average, max, min, product, countNums, stdDev, stdDevp, var, varp``.
"""

from collections.abc import Iterable, Sequence

from openpyxl.pivot.cache import (
    CacheDefinition,
    CacheField,
    CacheSource,
    SharedItems,
    WorksheetSource,
)
from openpyxl.pivot.record import RecordList
from openpyxl.pivot.table import (
    DataField,
    Location,
    PivotField,
    PivotTableStyle,
    RowColField,
    TableDefinition,
)
from openpyxl.utils import range_boundaries

__all__ = ["PivotTable", "FUNCTIONS"]

FUNCTIONS = frozenset(
    {
        "sum",
        "count",
        "countNums",
        "average",
        "max",
        "min",
        "product",
        "stdDev",
        "stdDevp",
        "var",
        "varp",
    }
)

_FUNC_ALIASES = {
    "avg": "average",
    "mean": "average",
    "counta": "count",
    "countnums": "countNums",
    "std": "stdDev",
    "stddev": "stdDev",
    "stddevp": "stdDevp",
}

_BUILTIN_NUMFMT = {
    "0": 1,
    "0.00": 2,
    "#,##0": 3,
    "#,##0.00": 4,
    "0%": 9,
    "0.00%": 10,
    "0.00e+00": 11,
    "mm-dd-yy": 14,
    "$#,##0": 5,
    "$#,##0.00": 7,
    "$#,##0_);($#,##0)": 5,
    "$#,##0.00;[red]($#,##0.00)": 8,
}


def _resolve_func(func: str) -> str:
    f = func.strip()
    if f in _FUNC_ALIASES:
        f = _FUNC_ALIASES[f]
    elif f.lower() in _FUNC_ALIASES:
        f = _FUNC_ALIASES[f.lower()]
    else:
        canon = {fn.lower(): fn for fn in FUNCTIONS}
        f = canon.get(f.lower(), f)
    if f not in FUNCTIONS:
        raise ValueError(
            f"Unknown aggregation function {func!r}. "
            f"Choose one of: {', '.join(sorted(FUNCTIONS))}"
        )
    return f


def _split_source(source: str, default_sheet: str | None) -> tuple[str, str]:
    """Split ``'Sheet!A1:D7'`` into (sheet_name, 'A1:D7')."""
    if "!" in source:
        sheet, ref = source.rsplit("!", 1)
        sheet = sheet.strip().strip("'")
    else:
        if default_sheet is None:
            raise ValueError(
                "source has no sheet prefix and no default sheet is known; "
                "use 'SheetName!A1:D7'"
            )
        sheet, ref = default_sheet, source
    return sheet, ref


class PivotTable:
    """Fluent builder for a single pivot table.

    Construct via :meth:`WorksheetProxy.add_pivot`. The underlying
    :class:`openpyxl.pivot.table.TableDefinition` and its cache are attached to
    the destination worksheet's ``_pivots`` list immediately and re-rendered in
    place on every fluent mutation, so the object is always consistent.
    """

    def __init__(
        self,
        dest_ws,
        source: str,
        anchor: str = "A3",
        name: str = "PivotTable1",
        cache_id: int | None = None,
    ) -> None:
        wb = dest_ws.parent
        src_sheet, src_ref = _split_source(source, _active_title(wb))
        self._dest_ws = dest_ws
        self._wb = wb
        self._name = name
        self._anchor = anchor
        self._src_sheet = src_sheet
        self._src_ref = src_ref
        self._fields = self._read_header(src_sheet, src_ref)
        self._index = {n: i for i, n in enumerate(self._fields)}

        self._rows: list[str] = []
        self._cols: list[str] = []
        self._filters: list[str] = []
        self._values: list[tuple[str, str, str | None, str | None]] = []
        self._style_name: str | None = None
        self._style_opts: dict[str, bool] = {}
        self._values_on_rows = False

        existing = {p.cacheId for p in getattr(wb, "_pivots", []) if p.cacheId}
        self._cache_id = cache_id or (max(existing) + 1 if existing else 1)

        self._table: TableDefinition | None = None
        self._cache: CacheDefinition | None = None
        self._attach()
        self._render()

    def _read_header(self, sheet_name: str, ref: str) -> list[str]:
        if sheet_name not in self._wb.sheetnames:
            raise ValueError(f"source sheet {sheet_name!r} not found in workbook")
        ws = self._wb[sheet_name]
        min_col, min_row, max_col, _ = range_boundaries(ref)
        headers = []
        for col in range(min_col, max_col + 1):
            val = ws.cell(row=min_row, column=col).value
            headers.append(str(val) if val is not None else f"Column{col}")
        if len(headers) != len(set(headers)):
            raise ValueError(f"source header row has duplicate names: {headers}")
        return headers

    def _field_idx(self, name: str) -> int:
        try:
            return self._index[name]
        except KeyError:
            raise ValueError(
                f"field {name!r} not in source header {self._fields}"
            ) from None

    def rows(self, *names: str) -> "PivotTable":
        """Add one or more fields to the row axis (in order)."""
        self._rows.extend(self._dedupe(names, self._rows))
        return self._render()

    def columns(self, *names: str) -> "PivotTable":
        """Add one or more fields to the column axis (in order)."""
        self._cols.extend(self._dedupe(names, self._cols))
        return self._render()

    def filters(self, *names: str) -> "PivotTable":
        """Add one or more report-filter (page) fields."""
        self._filters.extend(self._dedupe(names, self._filters))
        return self._render()

    def values(
        self,
        name: str,
        func: str = "sum",
        display_name: str | None = None,
        num_format: str | None = None,
    ) -> "PivotTable":
        """Add a value (data) field with an aggregation function.

        Args:
            name: source column name.
            func: aggregation function (see :data:`FUNCTIONS`); aliases like
                ``"avg"`` are accepted.
            display_name: header shown in the pivot (defaults to
                ``"Sum of <name>"`` etc.).
            num_format: number-format string. Set natively for Excel; for
                guaranteed rendering everywhere call :meth:`format_values`
                after ``ctx.sync()``.
        """
        f = _resolve_func(func)
        self._field_idx(name)
        self._values.append((name, f, display_name, num_format))
        return self._render()

    def values_axis(self, axis: str = "columns") -> "PivotTable":
        """Place the value-field labels on ``"columns"`` (default) or ``"rows"``.

        Only affects pivots with more than one value field.
        """
        if axis not in ("rows", "columns"):
            raise ValueError("axis must be 'rows' or 'columns'")
        self._values_on_rows = axis == "rows"
        return self._render()

    def style(
        self,
        name: str = "PivotStyleMedium9",
        *,
        row_stripes: bool = True,
        col_stripes: bool = False,
        row_headers: bool = True,
        col_headers: bool = True,
        last_column: bool = False,
    ) -> "PivotTable":
        """Apply a built-in pivot table style (rendered natively by Excel)."""
        self._style_name = name
        self._style_opts = dict(
            showRowStripes=row_stripes,
            showColStripes=col_stripes,
            showRowHeaders=row_headers,
            showColHeaders=col_headers,
            showLastColumn=last_column,
        )
        return self._render()

    @staticmethod
    def _dedupe(names: Iterable[str], existing: Sequence[str]) -> list[str]:
        seen = set(existing)
        out = []
        for n in names:
            if n not in seen:
                out.append(n)
                seen.add(n)
        return out

    def _attach(self) -> None:
        """Create empty model objects and register them on the worksheet."""
        cache = CacheDefinition(
            cacheSource=CacheSource(
                type="worksheet",
                worksheetSource=WorksheetSource(
                    ref=self._src_ref, sheet=self._src_sheet
                ),
            ),
            cacheFields=[self._cache_field(n) for n in self._fields],
            refreshOnLoad=True,
            recordCount=0,
            createdVersion=4,
            refreshedVersion=4,
            minRefreshableVersion=3,
        )
        cache.records = RecordList(count=0, r=[])
        cache._id = self._cache_id

        table = TableDefinition(
            name=self._name,
            cacheId=self._cache_id,
            dataCaption="Values",
            location=Location(
                ref=self._anchor, firstHeaderRow=1, firstDataRow=2, firstDataCol=1
            ),
            updatedVersion=4,
            minRefreshableVersion=3,
            createdVersion=4,
            indent=0,
            outline=True,
            outlineData=True,
            useAutoFormatting=True,
            itemPrintTitles=True,
            multipleFieldFilters=False,
            rowGrandTotals=True,
            colGrandTotals=True,
        )
        table.cache = cache
        self._cache = cache
        self._table = table

        pivots = getattr(self._dest_ws, "_pivots", None)
        if pivots is None:
            self._dest_ws._pivots = []
        self._dest_ws._pivots.append(table)

    def _cache_field(self, name: str) -> CacheField:
        # Flag numeric columns so Excel refreshes the cache correctly.
        ws = self._wb[self._src_sheet]
        min_col, min_row, max_col, max_row = range_boundaries(self._src_ref)
        col = min_col + self._fields.index(name)
        numeric = True
        seen_any = False
        for row in range(min_row + 1, max_row + 1):
            v = ws.cell(row=row, column=col).value
            if v is None:
                continue
            seen_any = True
            if not isinstance(v, int | float) or isinstance(v, bool):
                numeric = False
                break
        if seen_any and numeric:
            shared = SharedItems(
                containsSemiMixedTypes=False,
                containsString=False,
                containsNumber=True,
            )
        else:
            shared = SharedItems()
        return CacheField(name=name, sharedItems=shared)

    def _render(self) -> "PivotTable":
        """(Re)build the pivotField / row / col / data layout from the spec."""
        assert self._table is not None
        n = len(self._fields)
        axis_of: dict[int, str] = {}
        for nm in self._rows:
            axis_of[self._field_idx(nm)] = "axisRow"
        for nm in self._cols:
            axis_of[self._field_idx(nm)] = "axisCol"
        for nm in self._filters:
            axis_of[self._field_idx(nm)] = "axisPage"
        data_idx = {self._field_idx(nm) for nm, *_ in self._values}

        pivot_fields = []
        for i in range(n):
            kw: dict = {"showAll": False}
            if i in axis_of:
                kw["axis"] = axis_of[i]
            if i in data_idx:
                kw["dataField"] = True
            pivot_fields.append(PivotField(**kw))
        self._table.pivotFields = pivot_fields

        row_fields = [RowColField(x=self._field_idx(nm)) for nm in self._rows]
        col_fields = [RowColField(x=self._field_idx(nm)) for nm in self._cols]
        # x=-2 is the synthetic "data" field that positions value labels.
        if len(self._values) > 1:
            if self._values_on_rows:
                row_fields.append(RowColField(x=-2))
            else:
                col_fields.append(RowColField(x=-2))
        self._table.rowFields = row_fields
        self._table.colFields = col_fields
        self._table.pageFields = self._build_page_fields()

        data_fields = []
        for nm, func, disp, fmt in self._values:
            idx = self._field_idx(nm)
            df = DataField(
                name=disp or self._default_value_name(func, nm),
                fld=idx,
                subtotal=func,
                baseField=-1,
                baseItem=1048832,
            )
            if fmt is not None:
                df.numFmtId = self._numfmt_id(fmt)
            data_fields.append(df)
        self._table.dataFields = data_fields
        self._table.dataOnRows = self._values_on_rows

        if self._style_name is not None:
            self._table.pivotTableStyleInfo = PivotTableStyle(
                name=self._style_name, **self._style_opts
            )
        return self

    def _numfmt_id(self, fmt: str) -> int:
        """Resolve a format string to a numFmtId Excel will honor on refresh.

        Built-ins map to their fixed id; anything else is registered as a custom
        format on the workbook (numFmtId 164+).
        """
        from openpyxl.styles.numbers import builtin_format_id

        key = fmt.lower().replace(" ", "")
        if key in _BUILTIN_NUMFMT:
            return _BUILTIN_NUMFMT[key]
        try:
            bid = builtin_format_id(fmt)
            if bid is not None:
                return bid
        except Exception:
            pass
        idx = self._wb._number_formats.add(fmt)
        return 164 + idx

    def _build_page_fields(self):
        from openpyxl.pivot.table import PageField

        return [PageField(fld=self._field_idx(nm), hier=-1) for nm in self._filters]

    @staticmethod
    def _default_value_name(func: str, field: str) -> str:
        label = {
            "sum": "Sum",
            "count": "Count",
            "countNums": "Count",
            "average": "Average",
            "max": "Max",
            "min": "Min",
            "product": "Product",
            "stdDev": "StdDev",
            "stdDevp": "StdDevp",
            "var": "Var",
            "varp": "Varp",
        }.get(func, func.title())
        return f"{label} of {field}"

    def format_values(self, ctx, num_format: str | None = None) -> None:
        """Apply value number-formats to the *materialised* output cells.

        LibreOffice does not always apply a data-field's number format to the
        rendered cells. Call this after ``ctx.sync()`` to format the value
        region of the refreshed pivot directly, guaranteeing it renders in any
        consumer. ``num_format`` overrides per-field formats for all values.
        """
        formats = [num_format or fmt for _, _, _, fmt in self._values]
        fmt = next((f for f in formats if f), None)
        if fmt is None:
            return
        ctx.sync()
        loc = self._refreshed_location(ctx.path)
        if loc is None:
            return
        ref, first_data_row, first_data_col = loc
        _, _, max_col, max_row = range_boundaries(ref)
        from openpyxl.utils import get_column_letter

        c1 = get_column_letter(first_data_col)
        c2 = get_column_letter(max_col)
        ws = ctx.sheet(self._dest_ws.title)
        ws.range(f"{c1}{first_data_row}:{c2}{max_row}").apply_style(number_format=fmt)

    def _refreshed_location(self, path):
        """Read the post-refresh location of this pivot from the saved file."""
        from openpyxl import load_workbook

        wb = load_workbook(path)
        title = self._dest_ws.title
        if title not in wb.sheetnames:
            return None
        for p in wb[title]._pivots:
            if p.name == self._name and p.location and ":" in (p.location.ref or ""):
                min_col, min_row, _, _ = range_boundaries(p.location.ref)
                first_data_row = min_row + (p.location.firstDataRow or 2) - 1
                first_data_col = min_col + (p.location.firstDataCol or 1)
                return p.location.ref, first_data_row, first_data_col
        return None

    def __repr__(self) -> str:  # pragma: no cover
        return (
            f"PivotTable(name={self._name!r}, source={self._src_sheet}!{self._src_ref}, "
            f"rows={self._rows}, cols={self._cols}, values={[v[0] for v in self._values]})"
        )


def _active_title(wb) -> str | None:
    try:
        return wb.active.title
    except Exception:
        return None
