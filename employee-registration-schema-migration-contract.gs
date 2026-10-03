/* Phase 2B-2a-2a. Pure planning only; depends on the existing pure Safety Bundle.
 * Input is fresh inspected data, not the old preflight result or a caller plan.
 * Dates stay Dates; no employee data or target identity is included in output.
 * Evidence flags describe caller attestations, never proof of actual execution.
 * Snapshot v1:
 * - targetIdentityVerified: read-adapter identity attestation; no ID is returned.
 * - employees: rectangular header/data values, physical maxRows/maxColumns.
 * - columnInspections: full-physical-column inspection records in the existing
 *   preflight's {column,reusable,reasons} shape; required for every existing
 *   target whose header is absent. An absent physical AP needs no inspection
 *   here; the future write adapter must inspect it after expansion before write.
 * - ledger: null sheetId/values means verified missing, numeric sheetId plus
 *   rectangular values means verified present. [] means verified fully empty,
 *   not read failure. Dates in real operation rows must remain native Dates.
 * - emptyLedgerRecovery: legacy caller attestations retained for input compatibility.
 *   They never authorize repair of an empty ledger. A trusted recovery gate must
 *   independently match work record, target identity, sheet identity, creation time,
 *   fresh actual state, and human recovery approval; that gate is not implemented here.
 * - finalVerification: CONFIRMED/UNCONFIRMED; CONFIRMED alone cannot complete.
 * - externalChecks: six typed caller attestations used only for pending diagnostics.
 *   This planner never issues production execution permission. A later trusted gate
 *   must independently verify the two preflight limitations, external dependencies,
 *   write suspension, backup, production-version compatibility, fresh target state,
 *   and human final approval. COMPLETE describes schema plus caller verification,
 *   never execution authorization.
 * Malformed input throws a fixed error; semantically unsafe data yields a
 * manual-investigation result with no plan. Raw input is never returned.
 */
function planEmployeeRegistrationSchemaMigration_(snapshot) {
  const invalid = () => { throw new Error("EMPLOYEE_REGISTRATION_MIGRATION_SNAPSHOT_INVALID"); };
  function object(value, keys) {
    if (!value || Object.getPrototypeOf(value) !== Object.prototype ||
        Reflect.ownKeys(value).length !== keys.length) invalid();
    for (const key of keys) {
      const d = Object.getOwnPropertyDescriptor(value, key);
      if (!d || !Object.prototype.hasOwnProperty.call(d, "value") || !d.enumerable) invalid();
    }
  }
  function array(value) {
    if (!Array.isArray(value) || Object.getPrototypeOf(value) !== Array.prototype ||
        Reflect.ownKeys(value).length !== value.length + 1) invalid();
    for (let i = 0; i < value.length; i++) {
      const d = Object.getOwnPropertyDescriptor(value, String(i));
      if (!d || !Object.prototype.hasOwnProperty.call(d, "value") || !d.enumerable) invalid();
    }
  }
  function integer(value, minimum) {
    if (!Number.isSafeInteger(value) || value < minimum) invalid();
  }
  function bool(value) { if (typeof value !== "boolean") invalid(); }
  function grid(value) {
    array(value);
    for (const row of value) {
      array(row);
      for (const cell of row) {
        if (!["string", "boolean", "number"].includes(typeof cell) &&
            !(cell instanceof Date && Object.getPrototypeOf(cell) === Date.prototype &&
              Reflect.ownKeys(cell).length === 0 && Number.isFinite(cell.getTime()))) invalid();
        if (typeof cell === "number" && !Number.isFinite(cell)) invalid();
      }
    }
    if (value.length && value.some(row => row.length !== value[0].length)) invalid();
  }
  const checks = ["filterViewsReviewed", "crossSheetFormulasReviewed", "externalReferencesReviewed",
    "productionWritesStopped", "backupVerified", "productionVersionCompatible"];
  object(snapshot, ["contractVersion", "targetIdentityVerified", "timeZone", "employees",
    "columnInspections", "ledger", "emptyLedgerRecovery", "finalVerification", "externalChecks"]);
  if (snapshot.contractVersion !== 1) invalid();
  bool(snapshot.targetIdentityVerified);
  if (typeof snapshot.timeZone !== "string") invalid();
  if (!["CONFIRMED", "UNCONFIRMED"].includes(snapshot.finalVerification)) invalid();
  object(snapshot.externalChecks, checks);
  checks.forEach(key => bool(snapshot.externalChecks[key]));
  object(snapshot.employees, ["values", "maxRows", "maxColumns"]);
  const e = snapshot.employees;
  integer(e.maxRows, 1); integer(e.maxColumns, 1); grid(e.values);
  if (!e.values.length || !e.values[0].length || e.values.length > e.maxRows ||
      e.values[0].length > e.maxColumns) invalid();
  const headers = e.values[0], width = headers.length;
  array(snapshot.columnInspections);
  const inspectionMap = new Map(), reasons = ["VALUE", "FORMULA", "NOTE", "VALIDATION", "MERGE",
    "PROTECTION", "NAMED_RANGE", "CONDITIONAL_FORMAT", "FILTER"];
  for (const item of snapshot.columnInspections) {
    object(item, ["column", "reusable", "reasons"]);
    integer(item.column, 1); bool(item.reusable); array(item.reasons);
    if (item.column > e.maxColumns || inspectionMap.has(item.column) ||
        item.reasons.some(reason => !reasons.includes(reason)) ||
        new Set(item.reasons).size !== item.reasons.length ||
        item.reusable !== (item.reasons.length === 0)) invalid();
    inspectionMap.set(item.column, item);
  }
  object(snapshot.ledger, ["sheetId", "values"]);
  const ledger = snapshot.ledger;
  if (ledger.sheetId === null) { if (ledger.values !== null) invalid(); }
  else { integer(ledger.sheetId, 0); grid(ledger.values); }
  if (snapshot.emptyLedgerRecovery !== null) {
    object(snapshot.emptyLedgerRecovery, ["sheetId", "interruptedCreationVerified", "workRecordVerified"]);
    integer(snapshot.emptyLedgerRecovery.sheetId, 0);
    bool(snapshot.emptyLedgerRecovery.interruptedCreationVerified);
    bool(snapshot.emptyLedgerRecovery.workRecordVerified);
  }
  const issues = [], plan = [];
  const addIssue = code => { if (!issues.includes(code)) issues.push(code); };
  if (!snapshot.targetIdentityVerified) addIssue("TARGET_IDENTITY_UNVERIFIED");
  if (snapshot.timeZone !== "Asia/Tokyo") addIssue("TIMEZONE_UNEXPECTED");
  // Adding AP must not manufacture unnamed intermediate columns in a shorter schema.
  if (width < 41) addIssue("EMPLOYEE_HEADER_LAYOUT_UNEXPECTED");
  const headerInspection = employeeRegistrationPreflightHeaders_(headers);
  headerInspection.issues.filter(code => code !== "EMPLOYEE_EMPTY_HEADERS_MULTIPLE").forEach(addIssue);
  const positions = name => headers.flatMap((header, index) => header === name ? [index + 1] : []);
  const opPositions = positions("registration_operation_id");
  const legacyPositions = positions("initial_grant_check_target");
  let operation = "INVALID", legacy = "INVALID", ledgerState = "INVALID";
  function safeEmptyColumn(column) {
    const info = inspectionMap.get(column);
    return info && info.reusable && info.reasons.length === 0 &&
      e.values.slice(1).every(row => column > row.length || row[column - 1] === "");
  }
  if (opPositions.length === 1 && opPositions[0] === 35) operation = "APPLIED";
  else if (!opPositions.length && width >= 35 && headers[34] === "" && safeEmptyColumn(35))
    operation = "MISSING_SAFE";
  else addIssue("OPERATION_COLUMN_UNEXPECTED");
  if (legacyPositions.length === 1 && legacyPositions[0] === 42) {
    // This schema-only migration does not reinterpret existing legacy flags.
    // Blank/boolean values are valid; already established true/false flags are preserved.
    if (e.values.slice(1).every(row => row[41] === "" || typeof row[41] === "boolean")) legacy = "APPLIED";
    else addIssue("LEGACY_COLUMN_DATA_INVALID");
  } else if (!legacyPositions.length && width <= 42 &&
      (width < 42 || headers[41] === "") &&
      (e.maxColumns < 42 || safeEmptyColumn(42))) legacy = "MISSING_SAFE";
  else addIssue("LEGACY_COLUMN_UNEXPECTED");
  // The old preflight assumes at most one blank. This migration specifically permits
  // both fixed targets only when both are independently classified MISSING_SAFE.
  if (headerInspection.emptyColumns.length > 1 &&
      !(headerInspection.emptyColumns.length === 2 &&
        headerInspection.emptyColumns[0] === 35 && headerInspection.emptyColumns[1] === 42 &&
        operation === "MISSING_SAFE" && legacy === "MISSING_SAFE"))
    addIssue("EMPLOYEE_EMPTY_HEADERS_MULTIPLE");
  // No unnamed headers other than the two explicitly approved migration targets.
  if (headerInspection.emptyColumns.some(column =>
      !(column === 35 && operation === "MISSING_SAFE") &&
      !(column === 42 && legacy === "MISSING_SAFE"))) addIssue("EMPTY_HEADER_UNEXPECTED");
  let employeeProjection = null;
  if (operation !== "INVALID") {
    const virtual = e.values.map(row => row.slice());
    if (operation === "MISSING_SAFE") virtual[0][34] = "registration_operation_id";
    // Fill only a physically present, inspected, blank AP header in the copy.
    // Capacity-outside AP is absent from this grid and needs no virtual insertion.
    if (legacy === "MISSING_SAFE" && width === 42 && headers[41] === "")
      virtual[0][41] = "initial_grant_check_target";
    try {
      const formatter = value => {
        const parts = new Intl.DateTimeFormat("en", { timeZone: "Asia/Tokyo", year: "numeric",
          month: "2-digit", day: "2-digit" }).formatToParts(value);
        const part = name => parts.find(item => item.type === name).value;
        return part("year") + "-" + part("month") + "-" + part("day");
      };
      employeeProjection = employeeRegistrationProjectEmployees_(virtual, formatter, "registration_ready");
      const ids = employeeProjection.employeeRows.map(row => row.registration_operation_id.toLowerCase()).filter(Boolean);
      if (new Set(ids).size !== ids.length) throw new Error();
    } catch (ignored) { addIssue("EMPLOYEE_DATA_INVALID"); }
  }
  const expectedHeaders = EMPLOYEE_REGISTRATION_LEDGER_HEADERS_;
  let ledgerProjection = null;
  if (ledger.sheetId === null) ledgerState = "MISSING";
  else if (!ledger.values.length || ledger.values.every(row => row.every(cell => cell === ""))) {
    // No caller-provided flags can prove this is our interrupted creation.
    ledgerState = "EMPTY_UNPROVEN";
    addIssue("EMPTY_LEDGER_RECOVERY_UNPROVEN");
  } else {
    const lh = ledger.values[0];
    if (ledger.values.length === 1 && lh.length > 0 && lh.length < expectedHeaders.length &&
        lh.every((header, i) => header === expectedHeaders[i])) ledgerState = "PARTIAL_PREFIX";
    else if (lh.length === expectedHeaders.length && lh.every((header, i) => header === expectedHeaders[i])) {
      try { ledgerProjection = employeeRegistrationProjectOperationLedger_(ledger.values); ledgerState = "COMPLETE"; }
      catch (ignored) { addIssue("LEDGER_DATA_INVALID"); }
    } else addIssue("LEDGER_STRUCTURE_INVALID");
  }
  if (employeeProjection) {
    const hasOperationValues = employeeProjection.employeeRows.some(row => row.registration_operation_id !== "");
    if (ledgerState === "COMPLETE") {
      try { verifyEmployeeRegistrationLedgerEmployeeIntegrity_(ledgerProjection, employeeProjection); }
      catch (ignored) { addIssue("LEDGER_EMPLOYEE_INTEGRITY_INVALID"); }
    } else if (hasOperationValues) addIssue("EMPLOYEE_OPERATION_WITHOUT_LEDGER_RECORD");
  }
  const pendingChecks = checks.filter(key => !snapshot.externalChecks[key]);
  if (issues.length) return { contractVersion: 1, ok: false, state: "MANUAL_INVESTIGATION_REQUIRED",
    components: { operation: operation, legacy: legacy, ledger: ledgerState }, issues: issues,
    plan: [], pendingExternalChecks: pendingChecks };
  if (legacy === "MISSING_SAFE" && e.maxColumns < 42)
    plan.push({ type: "ENSURE_COLUMN_CAPACITY", sheet: "employees", minimum: 42 });
  if (operation === "MISSING_SAFE") plan.push({ type: "SET_HEADER", sheet: "employees",
    column: 35, header: "registration_operation_id" });
  if (legacy === "MISSING_SAFE") plan.push({ type: "SET_HEADER", sheet: "employees",
    column: 42, header: "initial_grant_check_target" });
  if (ledgerState === "MISSING") plan.push({ type: "CREATE_OPERATION_LEDGER",
    sheet: "employee_registration_operations", headers: expectedHeaders.slice() });
  if (ledgerState === "PARTIAL_PREFIX") {
    const completed = ledger.values[0].length;
    plan.push({ type: "COMPLETE_LEDGER_HEADERS", sheet: "employee_registration_operations",
      startColumn: completed + 1, headers: expectedHeaders.slice(completed) });
  }
  const complete = operation === "APPLIED" && legacy === "APPLIED" && ledgerState === "COMPLETE";
  const state = complete ? (snapshot.finalVerification === "CONFIRMED" ? "COMPLETE" : "VERIFICATION_REQUIRED") :
    (operation === "APPLIED" || legacy === "APPLIED" || ledgerState !== "MISSING" || e.maxColumns >= 42 ?
      "PARTIAL_MIGRATION" : "READY_FOR_MIGRATION");
  return { contractVersion: 1, ok: true, state: state,
    components: { operation: operation, legacy: legacy, ledger: ledgerState }, issues: [], plan: plan,
    pendingExternalChecks: pendingChecks };
}
