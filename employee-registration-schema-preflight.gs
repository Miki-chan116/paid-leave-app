/*
 * Phase 2B-2a-1: schema migration の読取専用 preflight。
 * plan は提案であり、このファイルに Spreadsheet の変更処理はない。
 */

function employeeRegistrationPreflightRead_(read, code) {
  try { return read(); }
  catch (error) { throw new Error("EMPLOYEE_REGISTRATION_PREFLIGHT_" + code + "_UNAVAILABLE"); }
}

function employeeRegistrationPreflightGrid_(values, code) {
  if (!isDenseEmployeeRegistrationArray_(values) || !values.length ||
      !isDenseEmployeeRegistrationArray_(values[0]) || !values[0].length) {
    throw new Error("EMPLOYEE_REGISTRATION_PREFLIGHT_" + code + "_READ_INVALID");
  }
  const width = values[0].length;
  for (const row of values) {
    if (!isDenseEmployeeRegistrationArray_(row) || row.length !== width) {
      throw new Error("EMPLOYEE_REGISTRATION_PREFLIGHT_" + code + "_READ_INVALID");
    }
  }
  return values;
}

function employeeRegistrationPreflightIntersectsColumn_(range, column) {
  const start = range.getColumn(), width = range.getNumColumns();
  if (!Number.isSafeInteger(start) || start < 1 ||
      !Number.isSafeInteger(width) || width < 1) {
    throw new Error("EMPLOYEE_REGISTRATION_PREFLIGHT_RANGE_INVALID");
  }
  return start <= column && column < start + width;
}

function employeeRegistrationPreflightColumnCells_(cells, rowCount, empty) {
  if (!isDenseEmployeeRegistrationArray_(cells) || cells.length !== rowCount) {
    throw new Error("EMPLOYEE_REGISTRATION_PREFLIGHT_COLUMN_READ_INVALID");
  }
  for (const row of cells) {
    if (!isDenseEmployeeRegistrationArray_(row) || row.length !== 1) {
      throw new Error("EMPLOYEE_REGISTRATION_PREFLIGHT_COLUMN_READ_INVALID");
    }
    if (!empty(row[0])) return false;
  }
  return true;
}

function employeeRegistrationPreflightReusableColumn_(sheet, column, maxRows) {
  // getDataRange() では見えない物理行まで調べる。読めない項目を空とみなさない。
  const range = employeeRegistrationPreflightRead_(
    () => sheet.getRange(1, column, maxRows, 1), "COLUMN_RANGE");
  const values = employeeRegistrationPreflightRead_(() => range.getValues(), "COLUMN_VALUES");
  const formulas = employeeRegistrationPreflightRead_(() => range.getFormulas(), "COLUMN_FORMULAS");
  const notes = employeeRegistrationPreflightRead_(() => range.getNotes(), "COLUMN_NOTES");
  const validations = employeeRegistrationPreflightRead_(
    () => range.getDataValidations(), "COLUMN_VALIDATIONS");
  const merged = employeeRegistrationPreflightRead_(() => range.getMergedRanges(), "COLUMN_MERGES");
  const protections = employeeRegistrationPreflightRead_(
    () => sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE), "PROTECTIONS");
  const sheetProtections = employeeRegistrationPreflightRead_(
    () => sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET), "PROTECTIONS");
  const named = employeeRegistrationPreflightRead_(() => sheet.getNamedRanges(), "NAMED_RANGES");
  const rules = employeeRegistrationPreflightRead_(
    () => sheet.getConditionalFormatRules(), "CONDITIONAL_FORMATS");
  const filter = employeeRegistrationPreflightRead_(() => sheet.getFilter(), "FILTER");
  const lists = [merged, protections, sheetProtections, named, rules];
  if (filter === undefined || lists.some(list => !isDenseEmployeeRegistrationArray_(list))) {
    throw new Error("EMPLOYEE_REGISTRATION_PREFLIGHT_METADATA_READ_INVALID");
  }
  const reasons = [];
  if (!employeeRegistrationPreflightColumnCells_(values, maxRows, value => value === "")) reasons.push("VALUE");
  if (!employeeRegistrationPreflightColumnCells_(formulas, maxRows, value => value === "")) reasons.push("FORMULA");
  if (!employeeRegistrationPreflightColumnCells_(notes, maxRows, value => value === "")) reasons.push("NOTE");
  if (!employeeRegistrationPreflightColumnCells_(validations, maxRows, value => value === null))
    reasons.push("VALIDATION");
  if (merged.length) reasons.push("MERGE");
  if (sheetProtections.length || protections.some(item =>
      employeeRegistrationPreflightIntersectsColumn_(item.getRange(), column))) reasons.push("PROTECTION");
  if (named.some(item => employeeRegistrationPreflightIntersectsColumn_(item.getRange(), column)))
    reasons.push("NAMED_RANGE");
  if (rules.some(rule => {
    const ranges = rule.getRanges();
    if (!isDenseEmployeeRegistrationArray_(ranges))
      throw new Error("EMPLOYEE_REGISTRATION_PREFLIGHT_METADATA_READ_INVALID");
    return ranges.some(item => employeeRegistrationPreflightIntersectsColumn_(item, column));
  })) reasons.push("CONDITIONAL_FORMAT");
  if (filter && employeeRegistrationPreflightIntersectsColumn_(filter.getRange(), column))
    reasons.push("FILTER");
  return { column: column, reusable: reasons.length === 0, reasons: reasons };
}

function employeeRegistrationPreflightHeaders_(headers) {
  const issues = [], positions = Object.create(null), emptyColumns = [];
  if (!isDenseEmployeeRegistrationArray_(headers)) {
    return { positions: positions, emptyColumns: emptyColumns, issues: ["EMPLOYEE_HEADERS_INVALID"] };
  }
  for (let index = 0; index < headers.length; index++) {
    const header = headers[index];
    if (header === "") { emptyColumns.push(index + 1); continue; }
    if (typeof header !== "string" || header !== header.trim()) {
      issues.push("EMPLOYEE_HEADER_INVALID"); continue;
    }
    if (Object.prototype.hasOwnProperty.call(positions, header)) issues.push("EMPLOYEE_HEADER_DUPLICATE");
    else positions[header] = index;
  }
  if (emptyColumns.length > 1) issues.push("EMPLOYEE_EMPTY_HEADERS_MULTIPLE");
  const required = ["employee_id", "display_employee_id"].concat(EMPLOYEE_REGISTRATION_INPUT_FIELDS_);
  if (required.some(key => !Object.prototype.hasOwnProperty.call(positions, key)))
    issues.push("EMPLOYEE_REQUIRED_HEADER_MISSING");
  return { positions: positions, emptyColumns: emptyColumns, issues: issues };
}

function employeeRegistrationPreflightLedger_(values) {
  if (values === null) return { status: "MISSING", rowCount: 0, issues: [] };
  const headers = values[0], expected = EMPLOYEE_REGISTRATION_LEDGER_HEADERS_;
  const names = Object.create(null), issues = [];
  for (const header of headers) {
    if (typeof header !== "string" || !header || header !== header.trim()) {
      issues.push("LEDGER_HEADER_INVALID"); continue;
    }
    if (Object.prototype.hasOwnProperty.call(names, header)) issues.push("LEDGER_HEADER_DUPLICATE");
    names[header] = true;
  }
  const prefix = headers.length > 0 && headers.length < expected.length &&
    headers.every((header, index) => header === expected[index]);
  if (prefix && values.length === 1 && !issues.length) {
    return { status: "PARTIAL_PREFIX", rowCount: 0, issues: [] };
  }
  if (headers.length !== expected.length ||
      headers.some((header, index) => header !== expected[index])) issues.push("LEDGER_HEADER_UNEXPECTED");
  if (issues.length) return { status: "INVALID", rowCount: values.length - 1, issues: issues };
  try { employeeRegistrationProjectOperationLedger_(values); }
  catch (error) { return { status: "INVALID", rowCount: values.length - 1, issues: ["LEDGER_DATA_INVALID"] }; }
  return { status: "COMPLETE", rowCount: values.length - 1, issues: [] };
}

function employeeRegistrationPreflightProjectedEmployees_(values, headers, candidate) {
  const virtual = values.map(row => row.slice());
  if (candidate) virtual[0][candidate - 1] = "registration_operation_id";
  else if (virtual[0].indexOf("registration_operation_id") === -1) {
    virtual.forEach((row, index) => row.push(index === 0 ? "registration_operation_id" : ""));
  }
  return employeeRegistrationProjectEmployees_(virtual,
    date => employeeRegistrationDateKeyWithFormatter_(date, headers.timeZone,
      (value, zone, pattern) => Utilities.formatDate(value, zone, pattern)), "registration_ready");
}

function employeeRegistrationClassifySchema_(snapshot) {
  const values = snapshot.employeeValues, meta = snapshot.emptyColumns;
  const header = employeeRegistrationPreflightHeaders_(values[0]);
  const issues = header.issues.slice(), plan = [];
  const hasOperation = Object.prototype.hasOwnProperty.call(header.positions, "registration_operation_id");
  const hasInitialGrant = Object.prototype.hasOwnProperty.call(header.positions, "initial_grant_check_target");
  if (!hasInitialGrant) issues.push("LEGACY_REQUIRED_COLUMN_MISSING");
  const candidate = header.emptyColumns.length === 1 ? header.emptyColumns[0] : null;
  if (candidate) {
    const inspection = meta.find(item => item.column === candidate);
    if (!inspection || !inspection.reusable) issues.push("EMPLOYEE_EMPTY_COLUMN_UNSAFE");
  }
  if (hasOperation && header.emptyColumns.length) issues.push("EMPLOYEE_EMPTY_HEADER_UNEXPECTED");
  const tailColumn = !hasOperation && !candidate && values[0].length + 1 <= snapshot.maxColumns
    ? snapshot.tailColumn : null;
  if (!hasOperation && !candidate && values[0].length + 1 <= snapshot.maxColumns &&
      (!tailColumn || tailColumn.column !== values[0].length + 1 || !tailColumn.reusable)) {
    issues.push("EMPLOYEE_TARGET_COLUMN_UNSAFE");
  }
  const ledger = employeeRegistrationPreflightLedger_(snapshot.ledgerValues);
  issues.push(...ledger.issues);
  if (ledger.status === "INVALID") issues.push("LEDGER_INVALID");
  let employeeSnapshot = null;
  if (!issues.some(code => code.startsWith("EMPLOYEE_"))) {
    try {
      employeeSnapshot = employeeRegistrationPreflightProjectedEmployees_(values,
        { timeZone: snapshot.timeZone }, candidate);
      const operationIds = new Set();
      for (const row of employeeSnapshot.employeeRows) {
        if (!row.registration_operation_id) continue;
        const key = row.registration_operation_id.toLowerCase();
        if (operationIds.has(key)) throw new Error("duplicate operation id");
        operationIds.add(key);
      }
    } catch (error) { issues.push("EMPLOYEE_DATA_INVALID"); }
  }
  if (ledger.status === "COMPLETE" && ledger.rowCount && employeeSnapshot) {
    if (!hasOperation) issues.push("LEDGER_DATA_WITHOUT_EMPLOYEE_OPERATION_COLUMN");
    else try {
      verifyEmployeeRegistrationLedgerEmployeeIntegrity_(
        employeeRegistrationProjectOperationLedger_(snapshot.ledgerValues), employeeSnapshot);
    } catch (error) { issues.push("LEDGER_EMPLOYEE_INTEGRITY_INVALID"); }
  }
  if (employeeSnapshot && employeeSnapshot.employeeRows.some(row => row.registration_operation_id) &&
      (ledger.status !== "COMPLETE" || ledger.rowCount === 0)) {
    issues.push("EMPLOYEE_OPERATION_WITHOUT_LEDGER_RECORD");
  }
  const fatal = issues.some(code => code !== "LEGACY_REQUIRED_COLUMN_MISSING");
  let state;
  if (fatal) state = "MANUAL_INVESTIGATION_REQUIRED";
  else {
    if (!hasOperation) {
      if (candidate) plan.push({ type: "REUSE_EMPTY_COLUMN", column: candidate,
        header: "registration_operation_id" });
      else {
        const column = values[0].length + 1;
        if (snapshot.maxColumns < column) plan.push({ type: "ENSURE_COLUMN_CAPACITY", minimum: column });
        plan.push({ type: "ADD_REGISTRATION_OPERATION_ID_COLUMN", column: column,
          header: "registration_operation_id" });
      }
    }
    if (ledger.status === "MISSING") plan.push({ type: "CREATE_OPERATION_LEDGER",
      headers: EMPLOYEE_REGISTRATION_LEDGER_HEADERS_.slice() });
    if (ledger.status === "PARTIAL_PREFIX") plan.push({ type: "COMPLETE_LEDGER_HEADERS",
      headers: EMPLOYEE_REGISTRATION_LEDGER_HEADERS_.slice() });
    state = plan.length === 0 ? "ALREADY_MIGRATED" :
      ((hasOperation || ledger.status !== "MISSING") ? "PARTIAL_MIGRATION" : "READY_FOR_MIGRATION");
  }
  return {
    ok: !fatal, state: state, issues: issues, plan: fatal ? [] : plan,
    employees: { rowCount: values.length - 1, headerCount: values[0].length,
      maxRows: snapshot.maxRows, maxColumns: snapshot.maxColumns,
      emptyHeaderColumns: header.emptyColumns, hasRegistrationOperationId: hasOperation,
      hasInitialGrantCheckTarget: hasInitialGrant,
      legacyRequiredColumnNext: hasInitialGrant ? null : values[0].length + 1,
      legacyRequiredColumnCapacityAvailable: hasInitialGrant ? null :
        snapshot.maxColumns >= values[0].length + 1,
      emptyColumns: meta, tailColumn: tailColumn },
    ledger: { exists: snapshot.ledgerValues !== null, status: ledger.status,
      rowCount: ledger.rowCount },
    timeZone: snapshot.timeZone,
    limitations: ["FILTER_VIEWS_NOT_INSPECTED", "CROSS_SHEET_FORMULA_REFERENCES_NOT_INSPECTED"]
  };
}

// 本番入口は snapshot を受け取らず、実 Spreadsheet の読取成功後だけ結果を返す。
function inspectEmployeeRegistrationSchemaReadOnly_() {
  const spreadsheet = employeeRegistrationPreflightRead_(
    () => SpreadsheetApp.openById(SS_ID), "SPREADSHEET");
  if (!spreadsheet || typeof spreadsheet.getSheetByName !== "function")
    throw new Error("EMPLOYEE_REGISTRATION_PREFLIGHT_SPREADSHEET_UNAVAILABLE");
  const employees = employeeRegistrationPreflightRead_(
    () => spreadsheet.getSheetByName("employees"), "EMPLOYEES_SHEET");
  if (!employees) throw new Error("EMPLOYEE_REGISTRATION_PREFLIGHT_EMPLOYEES_SHEET_UNAVAILABLE");
  const maxRows = employeeRegistrationPreflightRead_(
    () => employees.getMaxRows(), "MAX_ROWS");
  const maxColumns = employeeRegistrationPreflightRead_(
    () => employees.getMaxColumns(), "MAX_COLUMNS");
  if (!Number.isSafeInteger(maxRows) || maxRows < 1 ||
      !Number.isSafeInteger(maxColumns) || maxColumns < 1)
    throw new Error("EMPLOYEE_REGISTRATION_PREFLIGHT_DIMENSIONS_INVALID");
  const employeeValues = employeeRegistrationPreflightGrid_(employeeRegistrationPreflightRead_(
    () => employees.getDataRange().getValues(), "EMPLOYEES_READ"), "EMPLOYEES");
  if (employeeValues.length > maxRows || employeeValues[0].length > maxColumns)
    throw new Error("EMPLOYEE_REGISTRATION_PREFLIGHT_DIMENSIONS_INVALID");
  const header = employeeRegistrationPreflightHeaders_(employeeValues[0]);
  const emptyColumns = header.emptyColumns.map(column =>
    employeeRegistrationPreflightReusableColumn_(employees, column, maxRows));
  const targetColumn = employeeValues[0].length + 1;
  const hasOperation = Object.prototype.hasOwnProperty.call(
    header.positions, "registration_operation_id");
  const tailColumn = !hasOperation && header.emptyColumns.length === 0 &&
      targetColumn <= maxColumns
    ? employeeRegistrationPreflightReusableColumn_(employees, targetColumn, maxRows) : null;
  const ledgerSheet = employeeRegistrationPreflightRead_(
    () => spreadsheet.getSheetByName("employee_registration_operations"), "LEDGER_LOOKUP");
  if (ledgerSheet === undefined) {
    throw new Error("EMPLOYEE_REGISTRATION_PREFLIGHT_LEDGER_LOOKUP_UNAVAILABLE");
  }
  let ledgerValues = null;
  if (ledgerSheet !== null) ledgerValues = employeeRegistrationPreflightGrid_(employeeRegistrationPreflightRead_(
    () => ledgerSheet.getDataRange().getValues(), "LEDGER_READ"), "LEDGER");
  const timeZone = employeeRegistrationPreflightRead_(
    () => spreadsheet.getSpreadsheetTimeZone(), "TIMEZONE");
  if (typeof timeZone !== "string" || !timeZone.trim())
    throw new Error("EMPLOYEE_REGISTRATION_PREFLIGHT_TIMEZONE_INVALID");
  return employeeRegistrationClassifySchema_({ employeeValues: employeeValues,
    emptyColumns: emptyColumns, tailColumn: tailColumn, ledgerValues: ledgerValues,
    maxRows: maxRows, maxColumns: maxColumns, timeZone: timeZone });
}
