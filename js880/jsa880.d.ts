// TypeScript type definitions for JSA880.js
// Project: js880 — WPS JSA spreadsheet automation framework
// Definitions by: Engineer-1 <noreply@paperclip.ing>
// Issue: XXD-113 / XXD-116

// ============================================================
// Common type aliases
// ============================================================

// === Callback types (match Array2D's actual calling convention: (value, index)) ===

/** A column selector: "f1".."fN" string, 0-based column index, or lambda */
type ColumnSelector = string | number | ((row: any[], index: number) => any);

/** A predicate: returns true to keep/include, false to exclude */
type Predicate = (row: any[], index: number) => boolean;

/** A mapper: returns transformed value */
type Mapper = (row: any, index: number) => any;

/** A comparer: returns negative/zero/positive for sort order */
type Comparer = (a: any, b: any) => number;

/** An aggregator / reducer callback */
type Aggregator = (accumulator: any, current: any, index: number, array: any[]) => any;

/** A key selector — same shape as ColumnSelector but semantically "key extractor" */
type KeySelector = ColumnSelector;

/** A result selector — a mapper used in join/lookup operations */
type ResultSelector = Mapper;

// ============================================================
// QueryBuilder — returned by Array2D.prototype.where()
// ============================================================

declare class QueryBuilder {
    /** Switch to a different column for subsequent conditions */
    column(col: string | number): this;

    /** Greater than */
    gt(value: any): this;
    greaterThan(value: any): this;
    大于(value: any): this;

    /** Greater than or equal */
    gte(value: any): this;
    greaterThanOrEqual(value: any): this;
    大于等于(value: any): this;

    /** Less than */
    lt(value: any): this;
    lessThan(value: any): this;
    小于(value: any): this;

    /** Less than or equal */
    lte(value: any): this;
    lessThanOrEqual(value: any): this;
    小于等于(value: any): this;

    /** Equal */
    eq(value: any): this;
    equals(value: any): this;
    equal(value: any): this;
    等于(value: any): this;

    /** Not equal */
    neq(value: any): this;
    notEqual(value: any): this;
    不等于(value: any): this;

    /** Contains substring */
    contains(value: any): this;
    contain(value: any): this;
    包含(value: any): this;

    /** Value is in an array */
    in(values: any[]): this;
    在列表中(values: any[]): this;

    /** Value is not in an array */
    nin(values: any[]): this;
    notIn(values: any[]): this;
    不在列表中(values: any[]): this;

    /** Value is between min and max (inclusive) */
    between(min: any, max: any): this;
    在范围内(min: any, max: any): this;

    /** Matches a regex */
    match(regex: RegExp | string): this;
    regex(regex: RegExp | string): this;
    匹配(regex: RegExp | string): this;

    /** Is null or empty */
    isNull(): this;
    isEmpty(): this;
    为空(): this;

    /** Is not null and not empty */
    isNotNull(): this;
    isNotEmpty(): this;
    不为空(): this;

    /** Switch to AND logic, optionally set column */
    and(column?: string | number): this;
    且(column?: string | number): this;

    /** Switch to OR logic, optionally set column */
    or(column?: string | number): this;
    或(column?: string | number): this;

    /** Execute the query and return filtered Array2D */
    execute(): Array2D;
    exec(): Array2D;
    run(): Array2D;
    执行(): Array2D;
    val(): Array2D;
}

// ============================================================
// Array2D — the core data structure
// ============================================================

/**
 * Array2D — the core 2D array data structure.
 *
 * Behaves as a 0-indexed array of rows (each row is an array of values).
 * Supports numeric indexing: `arr[0]` returns the first row, `arr.length` gives row count.
 */
declare class Array2D {
    /** Numeric index access: returns the row at the given index */
    [index: number]: any[];

    /** Number of rows */
    length: number;

    /**
     * Create a new Array2D instance.
     * @param data — 2D array, 1D array, Array2D, null, or scalar value
     */
    constructor(data?: any[][] | any[] | Array2D | null | any);

    // ---- Internal / utility ----

    /** Internal: create a new Array2D instance from data, preserving _header */
    _new(data: any[][]): this;

    /**
     * Get current data as plain 2D array. If newData is provided, replace data
     * in place and return this for chaining.
     */
    val(): any[][];
    val(newData: any[][]): this;

    /** Promise-like: call onFulfilled with this Array2D */
    then(onFulfilled: (arr: this) => any): any;

    /** Promise-like: call onRejected on error */
    catch(onRejected: (err: any) => any): any;

    // ---- Inspection ----

    /** Check if empty */
    z是否为空(): boolean;
    isEmpty(): boolean;

    /** Total number of rows */
    z数量(): number;
    z行数(): number;
    count(): number;
    rowCount(): number;

    /** Total number of columns (from first row) */
    z列数(): number;
    colCount(): number;

    /** Matrix dimensions as { rows, cols } */
    z矩阵信息(): { rows: number; cols: number };
    matrixInfo(): { rows: number; cols: number };

    /** Check if a value is an error message */
    z错误值(msg: string): boolean;
    isError(msg: string): boolean;

    /** Framework version */
    z版本(): string;
    version(): string;

    // ---- Access ----

    /** Get cell value at (row, col) */
    z单元格(row: number, col: number): any;
    cell(row: number, col: number): any;

    /** Set cell value at (row, col). Returns this. */
    z设置单元格(row: number, col: number, value: any): this;
    setCell(row: number, col: number, value: any): this;

    /** Get first row */
    z首行(): any[];
    firstRow(): any[];

    /** Get last row */
    z末行(): any[];
    lastRow(): any[];

    /** Get first column as flat array */
    z首列(): any[];
    firstCol(): any[];

    /** Get last column as flat array */
    z末列(): any[];
    lastCol(): any[];

    /** Get a single row by index */
    z获取行(index: number): any[];
    getRow(index: number): any[];

    /** Get a single column by index */
    z获取列(index: number): any[];
    getCol(index: number): any[];

    /** Get first element (first row, or first cell of 1D) */
    z第一个(): any;
    first(): any;

    /** Get last element (last row, or last cell of 1D) */
    z最后一个(): any;
    last(): any;

    // ---- Modification (in-place) ----

    /** Add a row at the end. Returns this. */
    z添加行(row: any[]): this;
    z追加一项(...items: any[]): this;
    addRow(row: any[]): this;
    push(...items: any[]): number;

    /** Add a column at the given index. Returns this. */
    z添加列(col: any[], index?: number): this;
    addCol(col: any[], index?: number): this;

    /** Delete a row by index. Returns this. */
    z删除行(index: number): this;
    deleteRow(index: number): this;

    /** Delete a column by index. Returns this. */
    z删除列(index: number): this;
    deleteCol(index: number): this;

    /** Remove and return the last row */
    z尾部弹出一项(): any[];
    pop(): any[];

    /** Remove and return the first row */
    z删除第一个(): any[];
    shift(): any[];

    // ---- Clone / convert ----

    /** Create a deep copy */
    z克隆(): this;
    copy(): this;

    /** Clone as plain 2D array */
    z结果(): any[][];
    res(): any[][];

    /** Return an empty Array2D */
    z空结果(): this;

    /** Flatten to 1D array */
    z扁平化(): any[];
    flat(): any[];

    /** Transpose rows ↔ cols */
    z转置(): this;
    transpose(): this;

    /** Reverse row order (in-place) */
    z反转(): this;
    reverse(): this;

    /** Convert to JSON string */
    z转JSON(pretty?: boolean): string;
    toJson(pretty?: boolean): string;

    /** Convert to HTML table */
    z输出HTML(options?: any): string;
    html(options?: any): string;
    toHtml(options?: any): string;

    /** Convert to string with separators */
    z转字符串(rowSeparator?: string, colSeparator?: string): string;
    toString(): string;

    // ---- Write to sheet ----

    /** Write this Array2D to a Range */
    toRange(rng: any, clearBelow?: boolean): any;
    z写入单元格(rng: any, clearBelow?: boolean): any;

    // ---- Slice / splice ----

    /** Slice rows by index range. Returns new Array2D. */
    z行切片(start: number, end?: number): this;
    slice(start?: number, end?: number): this;

    /** Splice rows: remove and/or insert. Returns removed rows as new Array2D. */
    z行切片删除行(start: number, deleteCount?: number, ...items: any[][]): this;
    splice(start: number, deleteCount?: number, ...items: any[][]): this;

    // ---- Take / skip / pick ----

    /** Take first N rows */
    z取前N个(n: number): this;
    z取前几个(n: number): this;
    take(n: number): this;

    /** Skip first N rows */
    z跳过(n: number): this;
    z跳过前N个(n: number): this;
    z跳过前几个(n: number): this;
    skip(n: number): this;

    /** Pick random N rows */
    z挑选(n: number): this;
    pick(n: number): this;

    /** Repeat the array N times */
    z重复N次(n: number): this;
    repeat(n: number): this;

    /** Take every N-th row, optionally with offset */
    z间隔取数(offset: number, step?: number): this;
    nth(offset: number, step?: number): this;

    /** Skip rows while predicate is true */
    z跳过前面连续满足(predicate: Predicate): this;
    skipWhile(predicate: Predicate): this;

    /** Take rows while predicate is true */
    z取前面连续满足(predicate: Predicate): this;
    takeWhile(predicate: Predicate): this;

    // ---- Filter / map / reduce ----

    /**
     * Filter rows. Accepts variadic predicates (AND logic).
     * The last argument may be a number = skipHeader rows.
     * A single object argument filters by value equality on each key.
     */
    z筛选(...args: any[]): this;
    filter(...args: any[]): this;

    /**
     * Map rows. Accepts variadic mappers (pipe left-to-right).
     */
    z映射(...mappers: Array<string | Mapper>): this;
    map(...mappers: Array<string | Mapper>): this;

    /** Reduce (left fold) */
    z归约(callback: Aggregator, initialValue?: any): any;
    reduce(callback: Aggregator, initialValue?: any): any;

    /** Reduce right (right fold) */
    z倒序归约(callback: Aggregator, initialValue?: any): any;
    reduceRight(callback: Aggregator, initialValue?: any): any;

    /** Reduce right (right fold) */
    z倒序归约(callback: Aggregator, initialValue?: any): any;
    reduceRight(callback: Aggregator, initialValue?: any): any;

    /** Test if every row satisfies predicate */
    z全部满足(predicate: Predicate): boolean;
    every(predicate: Predicate): boolean;

    /** Test if some row satisfies predicate */
    z有满足(predicate: Predicate): boolean;
    some(predicate: Predicate): boolean;

    /** Iterate over each row (no return value) */
    forEach(callback: (row: any[], index: number) => void): void;

    /** Iterate in reverse */
    z倒序遍历执行(callback: (row: any[], index: number) => void): void;
    forEachRev(callback: (row: any[], index: number) => void): void;

    // ---- Chainable where() filter ----

    /** Start a chainable query builder on a column */
    where(column: string | number): QueryBuilder;
    z筛选链(column: string | number): QueryBuilder;

    // ---- Sort ----

    /** Standard sort (in-place). Returns this. */
    sort(compareFn?: Comparer): this;

    /** Sort ascending by natural order */
    z升序排序(): this;
    sortAsc(): this;

    /** Sort descending by natural order */
    z降序排序(): this;
    sortDesc(): this;

    /** Sort by key selector ascending */
    z按规则升序(keySelector: KeySelector): this;
    sortBy(keySelector: KeySelector): this;

    /** Sort by key selector descending */
    z按规则降序(keySelector: KeySelector): this;
    sortByDesc(keySelector: KeySelector): this;

    /** Sort rows by a single column */
    z行排序(colIndex: number | string, ascending?: boolean): this;
    sortRow(colIndex: number | string, ascending?: boolean): this;

    /** Sort columns by a single row */
    z列排序(rowIndex: number, ascending?: boolean): this;
    sortCol(rowIndex: number, ascending?: boolean): this;

    /** Multi-column sort */
    z多列排序(sortParams: Array<{ col: number | string; asc?: boolean }>, headerRows?: number, customOrder?: any[]): this;
    sortByCols(sortParams: Array<{ col: number | string; asc?: boolean }>, headerRows?: number, customOrder?: any[]): this;

    /** Custom list sort */
    z自定义排序(colIndex: number | string, orderList: any[], headerRows?: number): this;
    sortByList(colIndex: number | string, orderList: any[], headerRows?: number): this;

    /** Smart sort: auto-detect column type, then sort */
    z智能排序(col: number | string, direction?: 'asc' | 'desc', skipHeader?: number): this;

    // ---- Group ----

    /** Group rows by key selector, aggregate values */
    z分组(keySelector: KeySelector, valSelector?: ColumnSelector): Map<any, any[]>;
    groupBy(keySelector: KeySelector, valSelector?: ColumnSelector): Map<any, any[]>;

    /** Smart group: auto-detect column type and group */
    z智能分组(col: number | string, groupBy?: string): Map<any, any[]>;

    /** Pivot table from raw data */
    z透视(rowField: ColumnSelector, colField: ColumnSelector, valueField: ColumnSelector, aggregator?: Aggregator): this;
    pivotBy(rowField: ColumnSelector, colField: ColumnSelector, valueField: ColumnSelector, aggregator?: Aggregator): this;

    // ---- Distinct / deduplicate ----

    /** Deduplicate rows by key */
    z去重(keySelector?: KeySelector, resultSelector?: ResultSelector): this;
    distinct(keySelector?: KeySelector, resultSelector?: ResultSelector): this;
    zDistinct(keySelector?: KeySelector, resultSelector?: ResultSelector): this;

    // ---- Join operations ----

    /** Left join with another array */
    z左连接(brr: any[][], leftKeySelector: KeySelector, rightKeySelector: KeySelector, resultSelector?: ResultSelector): this;
    leftjoin(brr: any[][], leftKeySelector: KeySelector, rightKeySelector: KeySelector, resultSelector?: ResultSelector): this;

    /** Inner join */
    z内连接(brr: any[][], leftKeySelector: KeySelector, rightKeySelector: KeySelector, resultSelector?: ResultSelector): this;
    innerjoin(brr: any[][], leftKeySelector: KeySelector, rightKeySelector: KeySelector, resultSelector?: ResultSelector): this;

    /** Full outer join */
    z左右全连接(brr: any[][], leftKeySelector: KeySelector, rightKeySelector: KeySelector, resultSelector?: ResultSelector): this;
    fulljoin(brr: any[][], leftKeySelector: KeySelector, rightKeySelector: KeySelector, resultSelector?: ResultSelector): this;

    /** One-to-many left join (left table → expand right matches) */
    z一对多连接(brr: any[][], leftKeySelector: KeySelector, rightKeySelector: KeySelector, resultSelector?: ResultSelector): this;
    leftFulljoin(brr: any[][], leftKeySelector: KeySelector, rightKeySelector: KeySelector, resultSelector?: ResultSelector): this;

    /** Zip two arrays side-by-side */
    z左右连接(...arrays: any[][]): this;
    zip(...arrays: any[][]): this;

    // ---- Set operations ----

    /** Rows in this but not in brr */
    z排除(brr: any[][], leftSelector?: KeySelector, rightSelector?: KeySelector): this;
    except(brr: any[][], leftSelector?: KeySelector, rightSelector?: KeySelector): this;

    /** Rows in both this and brr */
    z取交集(brr: any[][], leftSelector?: KeySelector, rightSelector?: KeySelector): this;
    intersect(brr: any[][], leftSelector?: KeySelector, rightSelector?: KeySelector): this;

    /** Union (deduplicated merge) */
    z去重并集(brr: any[][], leftSelector?: KeySelector, rightSelector?: KeySelector): this;
    union(brr: any[][], leftSelector?: KeySelector, rightSelector?: KeySelector): this;

    // ---- Super lookup ----

    /** Multi-condition lookup */
    z超级查找(brr: any[][], leftKeySelector: KeySelector, rightKeySelector: KeySelector, resultSelector?: ResultSelector): this;
    superLookup(brr: any[][], leftKeySelector: KeySelector, rightKeySelector: KeySelector, resultSelector?: ResultSelector): this;

    // ---- Search / find ----

    /** Find first matching row */
    z查找单个(predicate: Predicate): any[] | undefined;
    find(predicate: Predicate): any[] | undefined;

    /** Find all matching row indices */
    z查找所有下标(predicate: Predicate): number[];
    findAllIndex(predicate: Predicate): number[];

    /** Find all matching row indices */
    z查找所有行下标(predicate: Predicate): number[];
    findRowsIndex(predicate: Predicate): number[];

    /** Find matching column indices within a given row */
    z查找所有列下标(rowIndex: number, predicate: Predicate): number[];
    findColsIndex(rowIndex: number, predicate: Predicate): number[];

    /** Find the first index where a cell matches value */
    z查找元素下标(predicate: Predicate): number;
    findIndexByPredicate(predicate: Predicate): number;

    /** Find index of a value */
    z值位置(value: any, fromIndex?: number): number;
    indexOf(value: any, fromIndex?: number): number;

    /** Find last index of a value */
    z从后往前值位置(value: any, fromIndex?: number): number;
    lastIndexOf(value: any, fromIndex?: number): number;

    /** Find index of a value in the 2D array */
    z查找索引(value: any): number;
    findIndex(value: any): number;

    /** Check if a value exists */
    z包含(value: any): boolean;
    includes(value: any): boolean;

    // ---- Insert / delete (bulk) ----

    /** Delete columns by index or array of indices */
    z批量删除列(cols: number | number[]): this;
    z批量删除列2(cols: number | number[]): this;
    deleteCols(cols: number | number[]): this;
    delcols(cols: number | number[]): this;

    /** Delete rows by index or array of indices */
    z批量删除行(rows: number | number[]): this;
    deleteRows(rows: number | number[]): this;

    /** Insert columns at position(s) */
    z批量插入列(colSelector: number | number[], value?: any, count?: number): this;
    insertCols(colSelector: number | number[], value?: any, count?: number): this;

    /** Insert rows at position(s) */
    z批量插入行(rowSelector: number | number[], value?: any, count?: number): this;
    insertRows(rowSelector: number | number[], value?: any, count?: number): this;

    /** Insert a row-number column starting from startNum */
    z插入行号(startNum?: number): this;
    insertRowNum(startNum?: number): this;

    // ---- Fill / pad / resize ----

    /** Fill with a value to specified dimensions */
    z填充(value: any, rows: number, cols: number): this;
    fill(value: any, rows: number, cols: number): this;

    /** Fill blank cells */
    z补齐空位(direction?: string, rangeAddress?: string, fillValue?: any): this;
    fillBlank(direction?: string, rangeAddress?: string, fillValue?: any): this;

    /** Pad to target column count and row count */
    z补齐数组(cols: number, rows: number, fillValue?: any): this;
    pad(cols: number, rows: number, fillValue?: any): this;

    /** Resize to target dimensions */
    z重设大小(rows: number, cols: number): this;
    resize(rows: number, cols: number): this;

    // ---- Convert to matrix ----

    /** Convert to a dense matrix, filling blanks */
    z转矩阵(fillValue?: any): any[][];
    toMatrix(fillValue?: any): any[][];

    // ---- Pagination ----

    /** Split into N pages */
    z按页数分页(pageCount: number): this[];
    pageByCount(pageCount: number): this[];

    /** Split into pages of given row size */
    z按行数分页(pageSize: number): this[];
    pageByRows(pageSize: number): this[];

    /** Split at given row indices */
    z按下标分页(indexes: number[]): this[];
    pageByIndexs(indexes: number[]): this[];

    // ---- Range operations ----

    /** Select a range of rows */
    z按范围选择(start: number, end: number): this;
    rangeSelect(start: number, end: number): this;

    /** Iterate over a range of rows */
    z按范围遍历(start: number, end: number, callback: (row: any[], index: number) => void): void;
    rangeForEach(start: number, end: number, callback: (row: any[], index: number) => void): void;

    /** Map over a rectangular region */
    z区域映射(startRow: number, endRow: number, startCol: number, endCol: number, mapper: Mapper): any[][];
    rangeMap(startRow: number, endRow: number, startCol: number, endCol: number, mapper: Mapper): any[][];

    // ---- Chunk / split ----

    /** Split into chunks of given size */
    z分块(size: number): this[];
    chunk(size: number): this[];

    // ---- Aggregate ----

    /** Sum a column */
    z求和(colSelector?: ColumnSelector): number;
    sum(colSelector?: ColumnSelector): number;

    /** Average of a column */
    z平均值(colSelector?: ColumnSelector): number;
    average(colSelector?: ColumnSelector): number;

    /** Max of a column */
    z最大值(colSelector?: ColumnSelector): number;
    max(colSelector?: ColumnSelector): number;

    /** Min of a column */
    z最小值(colSelector?: ColumnSelector): number;
    min(colSelector?: ColumnSelector): number;

    /** Median of a column */
    z中位数(colSelector?: ColumnSelector): number;
    median(colSelector?: ColumnSelector): number;

    // ---- Join text ----

    /** Join all rows into a single string */
    z连接(separator?: string): string;
    join(separator?: string): string;

    /** Join values from a column into a string */
    z文本连接(selector: ColumnSelector, separator?: string): string;
    textjoin(selector: ColumnSelector, separator?: string): string;

    // ---- Random ----

    /** Shuffle rows */
    z随机打乱(): this;
    shuffle(): this;

    /** Pick N random rows */
    random(n: number): this;
    z随机一项(n: number): this;

    // ---- Matrix layout / block ----

    /** Layout flat data into rows×cols matrix */
    z矩阵排版(cols: number, direction?: 'r' | 'c'): any[][];

    /** Apply an operation to every cell */
    z矩阵运算(op: (cell: any) => any): any[][];

    /** Block matrix aggregation */
    z分块矩阵(arr: any[][], rowSize: number, colSize: number, aggFunc?: Aggregator): any[][];

    // ---- Utility ----

    /** Replace null/undefined cells with empty string */
    z处理空值(): this;
    noNull(): this;

    /** Extract a single column as flat array */
    z提取列(colIndex: number): any[];
    pluck(colIndex: number): any[];

    /** Select specific columns */
    z选择列(cols: Array<number | string>, newHeaders?: string[]): this;
    selectCols(cols: Array<number | string>, newHeaders?: string[]): this;
    SelectCols(cols: Array<number | string>, newHeaders?: string[]): this;

    /** Select specific rows by index */
    z选择行(rowIndexes: number[]): this;
    selectRows(rowIndexes: number[]): this;

    /** Detect the data type of a column */
    z检测类型(col: number | string): { type: string; [key: string]: any };

    // ---- Concatenation ----

    /** Concatenate multiple arrays vertically */
    z上下连接(...others: any[][]): this;
    concat(...others: any[][]): this;

    // ---- Super pivot ----

    /**
     * Super pivot — advanced pivot table with multi-level row/col headers.
     * Returns a wrapped result with .val(), .res(), and .getMeta().
     */
    z超级透视(
        rowFields: string,
        colFields: string,
        dataFields: string,
        headerRows?: number,
        outputHeader?: number,
        separator?: string,
        options?: any
    ): {
        val(): any[][];
        res(): any[][];
        getMeta(): {
            version: string;
            rowFields: string[];
            rowTitles: string[];
            colFields: string[];
            colTitles: string[];
            dataFields: string[];
            grandTotal: any;
        };
    };
    superPivot(
        rowFields: string,
        colFields: string,
        dataFields: string,
        headerRows?: number,
        outputHeader?: number,
        separator?: string,
        options?: any
    ): {
        val(): any[][];
        res(): any[][];
        getMeta(): {
            version: string;
            rowFields: string[];
            rowTitles: string[];
            colFields: string[];
            colTitles: string[];
            dataFields: string[];
            grandTotal: any;
        };
    };

    // ============================================================
    // Static methods
    // ============================================================

    /** Parse a lambda expression string into a function */
    static parseLambda(expr: string): Function | null;
    static z解析函数表达式(expr: string): Function | null;

    /** Internal: filter by object conditions */
    static _filterByObject(data: any[][], condition: object): any[][];
    /** Internal: check a single condition */
    static _checkCondition(row: any[], condition: object, index: number): boolean;

    // -- Static helpers (mirroring instance methods) --

    static where(data: any[][], column: string | number): QueryBuilder;
    static html(arr: any[][], options?: any): string;

    // Block matrix
    static z分块矩阵(arr: any[][], rowSize: number, colSize: number, aggFunc?: Aggregator): any[][];

    // Matrix
    static getMatrix(totalRows: number, cols: number, direction?: 'r' | 'c'): any[][];
    static z矩阵分布(totalRows: number, cols: number, direction?: 'r' | 'c'): any[][];
    static rangeMatrix(arr: any[][], keySelector: KeySelector, dataArrays: any[][], aggregator?: Aggregator): any[][];
    static z区域矩阵(arr: any[][], keySelector: KeySelector, dataArrays: any[][], aggregator?: Aggregator): any[][];
    static rangeMatric(arr: any[][], keySelector: KeySelector, dataArrays: any[][], aggregator?: Aggregator): any[][];
    static toMatrix(arr: any[][], rows: number, cols: number, direction?: 'r' | 'c'): any[][];

    // Index generation
    static getIndexs(start: number, end: number, step?: number): number[];
    static z生成下标数组(start: number, end: number, step?: number): number[];
    static indexArray(arr: any[][], predicate: Predicate): number[];
    static z下标数组(arr: any[][], predicate: Predicate): number[];

    // Range operations
    static rangeForEach(arr: any[][], start: number, end: number, callback: (row: any[], index: number) => void): void;
    static z按范围遍历(arr: any[][], start: number, end: number, callback: (row: any[], index: number) => void): void;
    static rangeMap(arr: any[][], address: string, mapper: Mapper): any[][];
    static z局部映射(arr: any[][], address: string, mapper: Mapper): any[][];
    static findRange(arr: any[][], value: any): { row: number; col: number } | null;
    static rangeSelect(arr: any[][], start: number, end: number, startCol?: number, endCol?: number): any[][];

    // Rank
    static rank(arr: any[][], colSelector: ColumnSelector, type?: 'asc' | 'desc'): number[];
    static z排名(arr: any[][], colSelector: ColumnSelector, type?: 'asc' | 'desc'): number[];
    static rankGroup(arr: any[][], colSelector: ColumnSelector, groupCol: ColumnSelector, type?: 'asc' | 'desc', outputAll?: boolean): number[];
    static z分组排名(arr: any[][], colSelector: ColumnSelector, groupCol: ColumnSelector, type?: 'asc' | 'desc', outputAll?: boolean): number[];

    // Cross join
    static crossjoin(arr: any[][], brr: any[][]): any[][];
    static z笛卡尔积(arr: any[][], brr: any[][]): any[][];

    // Group into
    static groupInto(arr: any[][], keySelector: KeySelector, valueSelector: ColumnSelector, separator?: string): any[][];
    static z分组汇总(arr: any[][], keySelector: KeySelector, valueSelector: ColumnSelector, separator?: string): any[][];
    static groupIntoMap(arr: any[][], keySelector: KeySelector, valueSelector: ColumnSelector): Map<any, any[]>;
    static groupIntoJoin(targetData: any[][], sourceData: any[][], keySelector: KeySelector, valueSelector: ColumnSelector, separator?: string): any[][];

    // Aggregate
    static agg(arr: any[][], colSelector: ColumnSelector, aggType: string): number;
    static max(arr: any[][], selector: ColumnSelector): number;
    static min(arr: any[][], selector: ColumnSelector): number;
    static sum(arr: any[][], selector: ColumnSelector): number;
    static average(arr: any[][], selector: ColumnSelector): number;
    static median(arr: any[][], selector: ColumnSelector): number;

    // Join / set
    static leftjoin(arr: any[][], brr: any[][], leftKey: KeySelector, rightKey: KeySelector, resultSelector?: ResultSelector): any[][];
    static innerjoin(arr: any[][], brr: any[][], leftKey: KeySelector, rightKey: KeySelector, resultSelector?: ResultSelector): any[][];
    static fulljoin(arr: any[][], brr: any[][], leftKey: KeySelector, rightKey: KeySelector, resultSelector?: ResultSelector): any[][];
    static leftFulljoin(arr: any[][], brr: any[][], leftKey: KeySelector, rightKey: KeySelector, resultSelector?: ResultSelector): any[][];
    static except(arr: any[][], brr: any[][]): any[][];
    static intersect(arr: any[][], brr: any[][]): any[][];
    static union(arr: any[][], brr: any[][]): any[][];
    static superLookup(arr: any[][], lookupValue: any, colIndex: number, returnCol: number): any;

    // Filter / map / reduce
    static filter(arr: any[][] /* , predicate1, predicate2, ..., [skipHeader] */): any[][];
    static map(arr: any[][] /* , mapper1, mapper2, ... */): any[][];
    static reduce(arr: any[][], callback: Aggregator, initialValue?: any): any;
    static reduceRight(arr: any[][], callback: Aggregator, initialValue?: any): any;
    static forEach(arr: any[][], callback: (row: any[], index: number) => void): void;
    static forEachRev(arr: any[][], callback: (row: any[], index: number) => void): void;
    static some(arr: any[][], predicate: Predicate): boolean;
    static every(arr: any[][], predicate: Predicate): boolean;

    // Distinct
    static distinct(arr: any[][], keySelector?: KeySelector, resultSelector?: ResultSelector): any[][];

    // Sort
    static sort(arr: any[][], comparer: Comparer): any[][];
    static sortDesc(arr: any[][], comparer: Comparer): any[][];
    static sortByCols(arr: any[][], colsConfig: any[], skipHeader?: number): any[][];
    static sortByList(arr: any[][], col: number | string, orderList: any[], skipHeader?: number): any[][];

    // Smart operations
    static smartSort(arr: any[][], col: number | string, direction?: 'asc' | 'desc', skipHeader?: number): any[][];
    static smartGroup(arr: any[][], col: number | string, groupBy?: string): Map<any, any[]>;
    static detectType(data: any[][], colIndex: number): { type: string; [key: string]: any };
    static _parseDateForSort(dateVal: any): any;

    // Insert / delete
    static insertCols(arr: any[][], colPos: number | number[], values?: any, totalCols?: number): any[][];
    static deleteCols(arr: any[][], cols: number | number[]): any[][];
    static delCols(arr: any[][], cols: number | number[]): any[][];
    static deleteRows(arr: any[][], rows: number | number[]): any[][];
    static insertRows(arr: any[][], rowPos: number | number[], values?: any): any[][];
    static insertRowNum(arr: any[][], start?: number, title?: string): any[][];

    // Access
    static first(arr: any[][], predicate?: Predicate): any[];
    static last(arr: any[][], predicate?: Predicate): any[];
    static find(arr: any[][], predicate: Predicate): any[] | undefined;
    static findIndex(arr: any[][], predicate: Predicate): number;
    static findRowsIndex(arr: any[][], predicate: Predicate): number[];
    static findColsIndex(arr: any[][], fn: Predicate): number[];
    static findAllIndex(arr: any[][], fn: Predicate): number[];
    static includes(arr: any[][], value: any): boolean;
    static indexOf(arr: any[][], value: any): number;
    static lastIndexOf(arr: any[][], value: any): number;

    // Slice / splice / take / skip
    static slice(arr: any[][], start: number, end?: number): any[][];
    static splice(arr: any[][], start: number, deleteCount?: number): any[][];
    static take(arr: any[][], count: number): any[][];
    static skip(arr: any[][], count: number): any[][];

    // Fill / pad
    static pad(arr: any[][], cols: number, rows: number, fillValue?: any): any[][];
    static fillBlank(arr: any[][], direction?: string, rangeAddress?: string): any[][];

    // Pagination
    static pageByCount(arr: any[][], pageCount: number, pageNumber?: number): any[][];
    static pageByIndexs(arr: any[][], idxs: number[]): any[][][];

    // Convert
    static copy(arr: any[][], within?: boolean): any[][];
    static transpose(arr: any[][]): any[][];
    static flat(arr: any[][], mapper?: Mapper): any[];
    static join(arr: any[][], separator?: string): string;
    static textjoin(arr: any[][], selector: ColumnSelector, separator?: string): string;
    static toJson(arr: any[][], indent?: number): string;
    static toString(arr: any[][], separator?: string): string;
    static concat(arr: any[][], ...others: any[][]): any[][];
    static selectRows(arr: any[][], rows: number[]): any[][];
    static toRange(arr: any[][], rng: any): any;

    // Random
    static random(arr: any[][], n: number): any[][];
    static shuffle(arr: any[][]): any[][][];

    // Repeat
    static repeat(arr: any[][], count: number): any[][];
    static nth(arr: any[][], n: number, offset?: number): any[][];

    // Misc
    static copyWithin(arr: any[][], target: number, start: number, end: number): any[][];
    static version(): string;
    static count(arr: any[][]): number;
    static isEmpty(arr: any[][]): boolean;
    static push(arr: any[][], item: any): number;
    static pop(arr: any[][]): any[];
    static shift(arr: any[][]): any[];
    static reverse(arr: any[][]): any[][];
    static res(arr: any[][]): any[][];
    static groupBy(arr: any[][], keySelector: KeySelector): Map<any, any[]>;
}

// ============================================================
// JSA namespace — global utility functions
// ============================================================

declare namespace JSA {
    // -- Array operations --

    /** Transpose a 2D array */
    function z转置(arr: any[][]): any[][];
    function transpose(arr: any[][]): any[][];

    /** Select columns from a 2D array */
    function z选择列(arr: any[][], colIndexes: Array<number | string>, newHeaders?: string[]): any[][];
    function selectCols(arr: any[][], colIndexes: Array<number | string>, newHeaders?: string[]): any[][];

    /** Write an array to a Range */
    function z写入单元格(arr: any[][] | any[], rng: any, clearDown?: boolean): any;
    function toRange(arr: any[][] | any[], rng: any, clearDown?: boolean): any;

    // -- Type conversion --

    /** Convert a value to a number */
    function z转数值(text: string | number): number;
    function val(text: string | number): number;

    /** Convert a value to a string */
    function z转文本(v: any): string;
    function cstr(v: any): string;

    /** Convert to integer */
    function z取整数(v: any): number;
    function cint(v: any): number;

    /** Get decimal portion */
    function z取小数(v: any): number;
    function getDecimal(v: any): number;

    /** Convert an array to WPS formula array format */
    function z转公式数组(arr: any[][]): string;
    function toExcelArray(arr: any[][]): string;

    // -- Date / time --

    /** Get today's date */
    function z今天(): Date;
    function today(): Date;

    /** Convert a value to a date serial number */
    function z转日期数值(d: any): number;
    function cdate(d: any): number;

    /** Get month number (1-12) from a date */
    function month(date?: Date | string): number;

    /** Get current date and time */
    function now(): Date;

    /** Date difference */
    function z日期间隔(d1: any, d2: any, format: string): number;
    function datedif(d1: any, d2: any, format: string): number;

    // -- String operations --

    /** Replace substring */
    function z替换(str: string, find: string, replaceWith: string): string;
    function replace(str: string, find: string, replaceWith: string): string;

    /** Mid — extract substring */
    function z截取字符(str: string, start: number, len: number): string;
    function mid(str: string, start: number, len: number): string;

    /** Left — extract N characters from start */
    function z左取字符(str: string, len: number): string;
    function left(str: string, len: number): string;

    /** Right — extract N characters from end */
    function z右取字符(str: string, len: number): string;
    function right(str: string, len: number): string;

    /** Wildcard pattern matching */
    function z模糊匹配(str: string, pattern: string): boolean;
    function like(str: string, pattern: string): boolean;

    // -- Math / aggregate --

    /** Sum of arguments */
    function z求和(...values: number[]): number;
    function sum(...values: number[]): number;

    /** Max of arguments */
    function z最大值(...values: number[]): number;
    function max(...values: number[]): number;

    /** Min of arguments */
    function z最小值(...values: number[]): number;
    function min(...values: number[]): number;

    /** Average of arguments */
    function z平均值(...values: number[]): number;
    function average(...values: number[]): number;

    // -- Lookup --

    /** Match — find index of key in array, with mode */
    function z查找索引(
        关键字: any,
        数组: any[],
        结果列?: number,
        模式?: number,
        错误值?: any
    ): any;
    function match(
        关键字: any,
        数组: any[],
        结果列?: number,
        模式?: number,
        错误值?: any
    ): any;

    /** VLOOKUP — left-side lookup */
    function z左侧查找(
        关键字: any,
        数组: any[][],
        结果列: number,
        模式?: string,
        错误值?: any
    ): any;
    function vlookup(
        关键字: any,
        数组: any[][],
        结果列: number,
        模式?: string,
        错误值?: any
    ): any;

    /** XLOOKUP — enhanced lookup with separate search and return arrays */
    function z增强查找(
        关键字: any,
        查找数组: any[],
        结果数组: any[],
        错误值?: any
    ): any;
    function xlookup(
        关键字: any,
        查找数组: any[],
        结果数组: any[],
        错误值?: any
    ): any;

    // -- Number / sequence --

    /** Generate a numeric sequence */
    function z生成数字序列(start: number, end: number, step?: number): number[];
    function getIndexs(start: number, end: number, step?: number): number[];
    function getNumberArray(start: number, end: number, step?: number): number[];

    /** Random integer in range [start, end] */
    function z随机整数(start: number, end: number): number;
    function rndInt(start: number, end: number): number;

    /** Array of N random integers */
    function z随机整数数组(start: number, end: number, n: number): number[];
    function rndIntArray(start: number, end: number, n: number): number[];

    /** Shuffle an array */
    function z随机打乱(array: any[]): any[];
    function shuffle(array: any[]): any[];

    // -- Chinese RMB uppercase --

    /** Convert number to Chinese RMB uppercase */
    function z人民币大写(n: number): string;
    function rmbdx(n: number): string;

    // -- Evaluation --

    /** Evaluate an expression string */
    function z表达式求值(expr: string): any;
    function eval880(expr: string): any;

    // -- Path normalization --

    /** Normalize path separators */
    function z统一路径分隔符(path: string): string;
    function normalPath(path: string): string;

    // -- Delay --

    /** Sleep/delay for ts milliseconds */
    function z延时(ts: number): void;
    function delay(ts: number): void;

    // -- Matrix distribution --

    /** Distribute totalRows × cols as a matrix */
    function z矩阵分布(totalRows: number, cols: number, direction?: 'r' | 'c'): any[][];
    function getMatrix(totalRows: number, cols: number, direction?: 'r' | 'c'): any[][];

    // -- Lambda parsing --

    /** Parse a lambda expression string into a function */
    function z解析函数表达式(expr: string): Function | null;

    // -- VBA interop --

    /** Call a VBA procedure */
    function __jsaToVBA(procName: string, ...args: any[]): any;

    // -- Internal helpers --

    /** Internal: returns {} */
    function m(): {};
    /** Internal: returns {} */
    function S(): {};

    // -- Core k() / jsaLambda --

    /**
     * jsaLambda — core lambda execution engine.
     *
     * Supports:
     *   - Path strings: "Array2D.filter", "$$.filter"
     *   - Arrow expressions: "(x) => x > 2"
     *   - Pipe expressions: "x => x.filter(...).map(...)"
     *   - Multi-line code blocks
     *   - "-r" parameter for Range address conversion
     *   - __KJ_ARGS__ metadata extraction
     */
    function jsaLambda(fn: string | Function, ...args: any[]): any;

    /**
     * k() — formula wrapper with error handling.
     *
     * Wraps jsaLambda with:
     *   - Error formatting as "#K_ERR: pos=N, KIND, msg=..."
     *   - Null/undefined → "#K_ERR: ..."
     *   - Framework-not-loaded guard
     *
     * @example
     *   =k("JSA.sum", A1:A10)
     *   =k("x => x.filter(r => r.f3 > 100).map(r => r.f1)", A1:H40)
     *   =k("$$.filter", A1:H40, "r => r.f3 > 100", 1)
     */
    function k(fn: string | Function, ...args: any[]): any;

    namespace k {
        /** Print troubleshooting guide to Console */
        function help(): void;
    }
}

// ============================================================
// Global k() / jsaLambda — available without JSA. prefix
// ============================================================

/**
 * jsaLambda — core lambda execution engine (global).
 * Alias of JSA.jsaLambda.
 */
declare function jsaLambda(fn: string | Function, ...args: any[]): any;

/**
 * k() — formula wrapper with error handling (global).
 * Alias of JSA.k.
 */
declare function k(fn: string | Function, ...args: any[]): any;

// ============================================================
// ShtUtils — worksheet utilities
// ============================================================

declare class ShtUtils {
    constructor(initialSheet?: any);

    /** Get safe UsedRange */
    z安全已使用区域(工作表?: any): any;
    safeUsedRange(工作表?: any): any;

    /** Check if sheet name matches wildcard */
    z包含表名(表名: string, 表集合?: any): boolean;
    includesSht(表名: string, 表集合?: any): boolean;

    /** Filter sheets by name */
    z表名筛选(表名: string, 表集合?: any): any[];
    filterShts(表名: string, 表集合?: any): any[];

    /** Find a sheet by name or object */
    z查找表(sht: any, shts?: any): any;
    findSht(sht: any, shts?: any): any;

    /** Check if sheet is empty */
    z判断空表(工作表: any): boolean;
    isEmptySht(工作表: any): boolean;

    /** Delete a sheet */
    z删除表(工作表: any): void;
    deleteSht(工作表: any): void;

    /** Get sheet by code name */
    z按代码名称(表名: string, 表集合?: any): any;
    byCodeName(表名: string, 表集合?: any): any;

    /** Hide sheets */
    z隐藏表(表集合?: any): void;
    hideSheets(表集合?: any): void;

    /** Show sheets */
    z显示表(表集合?: any): void;
    showSheets(表集合?: any): void;

    /** Activate a sheet */
    z激活表(工作表: any): any;
    shtActivate(工作表: any): any;

    /** Get last row number */
    z最后一行(工作表?: any): number;
    lastRow(工作表?: any): number;

    /** Correct an invalid sheet name */
    z纠正表名(工作表名: string): string;
    correctShtName(工作表名: string): string;

    /** Sort sheets */
    z工作表排序(shts: any, sortFn?: Function | string[]): void;

    // Static aliases
    static z安全已使用区域(工作表?: any): any;
    static safeUsedRange(工作表?: any): any;
    static z包含表名(表名: string, 表集合?: any): boolean;
    static includesSht(表名: string, 表集合?: any): boolean;
    static z表名筛选(表名: string, 表集合?: any): any[];
    static filterShts(表名: string, 表集合?: any): any[];
    static z查找表(sht: any, shts?: any): any;
    static findSht(sht: any, shts?: any): any;
    static z判断空表(工作表: any): boolean;
    static isEmptySht(工作表: any): boolean;
    static z删除表(工作表: any): void;
    static deleteSht(工作表: any): void;
    static z按代码名称(表名: string, 表集合?: any): any;
    static byCodeName(表名: string, 表集合?: any): any;
    static z隐藏表(表集合?: any): void;
    static hideSheets(表集合?: any): void;
    static z显示表(表集合?: any): void;
    static showSheets(表集合?: any): void;
    static z激活表(工作表: any): any;
    static shtActivate(工作表: any): any;
    static z最后一行(工作表?: any): number;
    static lastRow(工作表?: any): number;
    static z纠正表名(工作表名: string): string;
    static correctShtName(工作表名: string): string;
}

// ============================================================
// RngUtils — range utilities
// ============================================================

declare class RngUtils {
    constructor(initialRange?: any);
    // Methods are worksheet-level — included for namespace coverage
    [key: string]: any;
}

// ============================================================
// IO utilities (if available in the loaded version)
// ============================================================

declare namespace IO {
    function z是否文件(path: string): boolean;
    function IsFile(path: string): boolean;
    function z是否文件夹(path: string): boolean;
    function IsDirectory(path: string): boolean;
    function z文件名(path: string): string;
    function getFileName(path: string): string;
    function z纯文件名(path: string): string;
    function getFileNameNoType(path: string): string;
    function z文件后缀(path: string): string;
    function getFileType(path: string): string;
}
