/* Phase 2B-2a-2b-2. Internal read-only entry; no public wrapper or top-level calls.
 * resolveConfiguration and readApprovedTargetRecord are trusted environment seams.
 * The latter MUST read an independently human-approved record, not echo config.
 * spreadsheetApp MUST be the trusted native service (native enums), or a test double.
 * Injected arbitrary code is not sandboxed by this adapter.
 * This adapter never creates/updates that record. No real IDs belong in source.
 * reusable covers the existing nine inspections ONLY; external checks stay false.
 * Returned snapshot/evidence are PRIVATE memory, never a public/log payload.
 * Two equal passes detect observed drift, not a transaction or execution approval.
 */
function readEmployeeRegistrationSchemaMigrationSnapshot_(environment) {
  const checks = ['filterViewsReviewed', 'crossSheetFormulasReviewed', 'externalReferencesReviewed',
    'productionWritesStopped', 'backupVerified', 'productionVersionCompatible'];
  const coverage = ['VALUE', 'FORMULA', 'NOTE', 'VALIDATION', 'MERGE', 'PROTECTION',
    'NAMED_RANGE', 'CONDITIONAL_FORMAT', 'FILTER'];
  const fail = () => { throw new Error('EMPLOYEE_REGISTRATION_MIGRATION_READ_FAILED'); };
  function exact(value, keys) {
    if (!value || Object.getPrototypeOf(value) !== Object.prototype ||
        Reflect.ownKeys(value).length !== keys.length) fail();
    for (const key of keys) {
      const d = Object.getOwnPropertyDescriptor(value, key);
      if (!d || !d.enumerable || !Object.prototype.hasOwnProperty.call(d, 'value')) fail();
    }
  }
  function integer(value, minimum) {
    if (!Number.isSafeInteger(value) || value < minimum) fail();
    return value;
  }
  function dense(value) {
    if (!isDenseEmployeeRegistrationArray_(value)) fail();
    return value;
  }
  function copyCell(value) {
    if (value instanceof Date && Object.getPrototypeOf(value) === Date.prototype &&
        Reflect.ownKeys(value).length === 0 && Number.isFinite(value.getTime())) return new Date(value.getTime());
    if (typeof value === 'string' || typeof value === 'boolean' ||
        (typeof value === 'number' && Number.isFinite(value))) return value;
    fail();
  }
  function grid(value, rows, columns, cell) {
    dense(value);
    if (value.length !== rows) fail();
    return value.map(row => {
      dense(row);
      if (row.length !== columns) fail();
      return row.map(cell);
    });
  }
  // Typed, deterministic internal comparison; never logged/persisted. Date != string.
  function canonical(value) {
    if (value instanceof Date) return ['date', value.getTime()];
    if (value === null) return ['null'];
    if (typeof value === 'number' && Object.is(value, -0)) return ['number', '-0'];
    if (Array.isArray(value)) return ['array', value.map(canonical)];
    if (typeof value === 'object') return ['object', Object.keys(value).sort().map(k => [k, canonical(value[k])])];
    return [typeof value, value];
  }
  function geometry(range) {
    return {row: integer(range.getRow(), 1), column: integer(range.getColumn(), 1),
      rows: integer(range.getNumRows(), 1), columns: integer(range.getNumColumns(), 1),
      sheetId: integer(range.getSheet().getSheetId(), 0)};
  }
  function criteria(value, relative = false) {
    if (value === null) return null;
    if (Array.isArray(value)) return dense(value).map(item => criteria(item, relative));
    if (relative) return {relativeDate: enumName(value,
      ['TODAY', 'TOMORROW', 'YESTERDAY', 'PAST_WEEK', 'PAST_MONTH', 'PAST_YEAR'], app.RelativeDate)};
    if (value && typeof value.getColumn === 'function') return geometry(value);
    return copyCell(value);
  }
  // Official DataValidationCriteria names; unknown/future criteria fail closed.
  const criteriaNames = ['DATE_AFTER', 'DATE_BEFORE', 'DATE_BETWEEN', 'DATE_EQUAL_TO',
    'DATE_IS_VALID_DATE', 'DATE_NOT_BETWEEN', 'DATE_ON_OR_AFTER', 'DATE_ON_OR_BEFORE',
    'NUMBER_BETWEEN', 'NUMBER_EQUAL_TO', 'NUMBER_GREATER_THAN', 'NUMBER_GREATER_THAN_OR_EQUAL_TO',
    'NUMBER_LESS_THAN', 'NUMBER_LESS_THAN_OR_EQUAL_TO', 'NUMBER_NOT_BETWEEN', 'NUMBER_NOT_EQUAL_TO',
    'TEXT_CONTAINS', 'TEXT_DOES_NOT_CONTAIN', 'TEXT_EQUAL_TO', 'TEXT_IS_VALID_EMAIL', 'TEXT_IS_VALID_URL',
    'VALUE_IN_LIST', 'VALUE_IN_RANGE', 'CUSTOM_FORMULA', 'CHECKBOX', 'DATE_AFTER_RELATIVE',
    'DATE_BEFORE_RELATIVE', 'DATE_EQUAL_TO_RELATIVE'];
  let app, rangeType, sheetType, criteriaTypes;
  function enumName(value, names, namespace) {
    // Reject bad raw types BEFORE coercion. Membership uses native enum identity.
    if (value === null || value === undefined ||
        !['string', 'object'].includes(typeof value) || !namespace) fail();
    const matches = names.filter(name => namespace[name] === value);
    if (matches.length !== 1 || String(value) !== matches[0]) fail();
    return matches[0];
  }
  function boolean(value) { if (typeof value !== 'boolean') fail(); return value; }
  function validation(value) {
    if (value === null) return null;
    if (!value || typeof value.getCriteriaType !== 'function') fail();
    const type = enumName(value.getCriteriaType(), criteriaNames, criteriaTypes);
    const allowInvalid = value.getAllowInvalid(), helpText = value.getHelpText();
    if (typeof allowInvalid !== 'boolean' || (helpText !== null && typeof helpText !== 'string')) fail();
    return {type: type, arguments: dense(value.getCriteriaValues()).map(item => criteria(item, type.endsWith('_RELATIVE'))),
      allowInvalid: allowInvalid, helpText: helpText};
  }
  // Explicit comparison scope: type, geometry, warning mode, domain edit, current
  // caller edit capability, explicit editor emails and unprotected geometries.
  // These private records do NOT prove group membership, target-audience policy,
  // description/name-binding changes or every protection-related external change.
  function protectionCollection(items) {
    dense(items);
    // Bound newly materialized protection evidence. Service calls return whole
    // collections; overflow stops before mapping, never truncates the evidence.
    if (items.length > 1000) fail();
    return items;
  }
  function protection(item, expectedType, maxRows, maxColumns, sheetId) {
    if (item.getProtectionType() !== expectedType) fail();
    const position = geometry(item.getRange());
    const bounded = part => part.sheetId === sheetId &&
      part.row + part.rows - 1 <= maxRows && part.column + part.columns - 1 <= maxColumns;
    if (!bounded(position)) fail();
    if (expectedType === sheetType && (position.row !== 1 || position.column !== 1 ||
        position.rows !== maxRows || position.columns !== maxColumns)) fail();
    const warningOnly = boolean(item.isWarningOnly());
    // getEditors/canDomainEdit may throw without protection-edit permission.
    // Missing permission is failure, never empty evidence or partial coverage.
    const domainEdit = boolean(item.canDomainEdit());
    const canEdit = boolean(item.canEdit());
    const editors = protectionCollection(item.getEditors()).map(user => {
      const email = user.getEmail();
      if (typeof email !== 'string' || email !== email.trim() || !/^[^\s@]+@[^\s@]+$/.test(email)) fail();
      return email;
    }).sort();
    if (new Set(editors).size !== editors.length) fail();
    const unprotected = protectionCollection(item.getUnprotectedRanges()).map(geometry);
    if (unprotected.some(part => !bounded(part)) || (expectedType === rangeType && unprotected.length)) fail();
    unprotected.sort((a, b) => a.row - b.row || a.column - b.column || a.rows - b.rows || a.columns - b.columns);
    return {type: expectedType === rangeType ? 'RANGE' : 'SHEET', range: position,
      warningOnly: warningOnly, domainEdit: domainEdit, canEdit: canEdit,
      editors: editors, unprotected: unprotected};
  }
  function protectionList(items, type, maxRows, maxColumns, sheetId) {
    return protectionCollection(items).map(item => protection(item, type, maxRows, maxColumns, sheetId))
      .sort((a, b) => {
        const left = JSON.stringify(canonical(a)), right = JSON.stringify(canonical(b));
        return left < right ? -1 : left > right ? 1 : 0;
      });
  }
  function inspectColumn(sheet, column, maxRows, maxColumns, sheetId) {
    const range = sheet.getRange(1, column, maxRows, 1);
    const position = geometry(range);
    if (position.row !== 1 || position.column !== column || position.rows !== maxRows ||
        position.columns !== 1 || position.sheetId !== sheetId) fail();
    const values = grid(range.getValues(), maxRows, 1, copyCell);
    const formulas = grid(range.getFormulas(), maxRows, 1, text);
    const notes = grid(range.getNotes(), maxRows, 1, text);
    const validations = grid(range.getDataValidations(), maxRows, 1, validation);
    const merges = dense(range.getMergedRanges()).map(geometry);
    const protections = protectionList(sheet.getProtections(rangeType), rangeType, maxRows, maxColumns, sheetId);
    const sheetProtections = protectionList(sheet.getProtections(sheetType), sheetType, maxRows, maxColumns, sheetId);
    const named = dense(sheet.getNamedRanges()).map(item => geometry(item.getRange()));
    const rules = dense(sheet.getConditionalFormatRules()).map(rule =>
      dense(rule.getRanges()).map(geometry));
    const filterObject = sheet.getFilter();
    if (filterObject === undefined) fail();
    const filter = filterObject === null ? null : geometry(filterObject.getRange());
    const allRanges = merges.concat(protections.map(item => item.range), named, ...rules, filter ? [filter] : []);
    if (allRanges.some(item => item.sheetId !== sheetId ||
        item.row + item.rows - 1 > maxRows || item.column + item.columns - 1 > maxColumns)) fail();
    const intersects = item => item.column <= column && column < item.column + item.columns;
    const reasons = [];
    if (values.some(row => row[0] !== '')) reasons.push('VALUE');
    if (formulas.some(row => row[0] !== '')) reasons.push('FORMULA');
    if (notes.some(row => row[0] !== '')) reasons.push('NOTE');
    if (validations.some(row => row[0] !== null)) reasons.push('VALIDATION');
    if (merges.some(intersects)) reasons.push('MERGE');
    if (sheetProtections.length || protections.some(item => intersects(item.range))) reasons.push('PROTECTION');
    if (named.some(intersects)) reasons.push('NAMED_RANGE');
    if (rules.some(items => items.some(intersects))) reasons.push('CONDITIONAL_FORMAT');
    if (filter && intersects(filter)) reasons.push('FILTER');
    return {inspection: {column: column, reusable: reasons.length === 0, reasons: reasons},
      evidence: {column: column, coverage: coverage.slice(), values: values, formulas: formulas,
        notes: notes, validations: validations, merges: merges, protections: protections,
        sheetProtections: sheetProtections, named: named, rules: rules, filter: filter}};
  }
  function text(value) { if (typeof value !== 'string') fail(); return value; }
  try {
    exact(environment, ['resolveConfiguration', 'readApprovedTargetRecord', 'spreadsheetApp', 'now']);
    if (typeof environment.resolveConfiguration !== 'function' ||
        typeof environment.readApprovedTargetRecord !== 'function' || typeof environment.now !== 'function') fail();
    app = environment.spreadsheetApp;
    if (!app || typeof app.openById !== 'function') fail();
    const types = app.ProtectionType;
    rangeType = types && types.RANGE; sheetType = types && types.SHEET;
    if (rangeType === sheetType || enumName(rangeType, ['RANGE', 'SHEET'], types) !== 'RANGE' ||
        enumName(sheetType, ['RANGE', 'SHEET'], types) !== 'SHEET') fail();
    criteriaTypes = app.DataValidationCriteria;
    const config = environment.resolveConfiguration();
    const approval = environment.readApprovedTargetRecord();
    exact(config, ['spreadsheetId']);
    exact(approval, ['spreadsheetId', 'employeesSheetId', 'ledgerSheetId', 'timeZone']);
    if (config === approval || typeof config.spreadsheetId !== 'string' ||
        !/^[A-Za-z0-9_-]+$/.test(config.spreadsheetId) ||
        config.spreadsheetId !== approval.spreadsheetId || approval.timeZone !== 'Asia/Tokyo') fail();
    integer(approval.employeesSheetId, 0);
    if (approval.ledgerSheetId !== null) integer(approval.ledgerSheetId, 0);
    if (approval.ledgerSheetId === approval.employeesSheetId) fail();
    // Copy approved primitives before reads; callbacks cannot mutate our expectation.
    const expected = {spreadsheetId: config.spreadsheetId, employeesSheetId: approval.employeesSheetId,
      ledgerSheetId: approval.ledgerSheetId, timeZone: approval.timeZone};
    const startedAt = integer(environment.now(), 0);
    function readPass() {
      const ss = app.openById(expected.spreadsheetId);
      if (ss.getId() !== expected.spreadsheetId || ss.getSpreadsheetTimeZone() !== expected.timeZone) fail();
      const sheets = dense(ss.getSheets());
      const inventory = sheets.map(sheet => ({name: text(sheet.getName()), id: integer(sheet.getSheetId(), 0)}));
      if (new Set(inventory.map(item => item.name)).size !== inventory.length ||
          new Set(inventory.map(item => item.id)).size !== inventory.length) fail();
      function readSheet(name, id) {
        const sheet = ss.getSheetByName(name);
        const matches = inventory.filter(item => item.name === name);
        if (sheet === null) {
          if (id !== null || matches.length) fail();
          return null;
        }
        if (!sheet || id === null || matches.length !== 1 || matches[0].id !== id ||
            sheet.getSheetId() !== id || sheet.getName() !== name) fail();
        const maxRows = integer(sheet.getMaxRows(), 1), maxColumns = integer(sheet.getMaxColumns(), 1);
        const lastRow = integer(sheet.getLastRow(), 0), lastColumn = integer(sheet.getLastColumn(), 0);
        if (lastRow > maxRows || lastColumn > maxColumns || (lastRow === 0) !== (lastColumn === 0)) fail();
        const rows = Math.max(1, lastRow), columns = Math.max(1, lastColumn);
        // Explicit resource ceiling: never silently truncate a read.
        if (rows * columns > 1000000 || maxRows > 1000000) fail();
        const range = sheet.getRange(1, 1, rows, columns);
        const position = geometry(range);
        if (position.row !== 1 || position.column !== 1 || position.rows !== rows ||
            position.columns !== columns || position.sheetId !== id) fail();
        const values = grid(range.getValues(), rows, columns, copyCell);
        const formulas = grid(range.getFormulas(), rows, columns, text);
        if (lastRow === 0 && (values.some(row => row.some(cell => cell !== '')) ||
            formulas.some(row => row.some(cell => cell !== '')))) fail();
        return {sheet: sheet, id: id, maxRows: maxRows, maxColumns: maxColumns,
          lastRow: lastRow, lastColumn: lastColumn, values: values, formulas: formulas};
      }
      const employees = readSheet('employees', expected.employeesSheetId);
      if (!employees || employees.lastColumn < 41 || employees.lastRow < 1) fail();
      const ledger = readSheet('employee_registration_operations', expected.ledgerSheetId);
      // Ledger formulas are not proof of an empty/prefix/static ledger. Fail closed.
      if (ledger && ledger.formulas.some(row => row.some(cell => cell !== ''))) fail();
      const inspections = [], columnEvidence = [];
      for (const column of [35, 42]) {
        if (column > employees.maxColumns) continue;
        // Inspect both targets, even applied ones, to retain observed metadata drift.
        const inspected = inspectColumn(employees.sheet, column, employees.maxRows, employees.maxColumns, employees.id);
        // A/B: validate EACH pass internally before C: comparing the two passes.
        // API blank cells are represented by ''; null/undefined are not substitutes.
        if (column <= employees.lastColumn) {
          for (let index = 0; index < employees.values.length; index++) {
            if (JSON.stringify(canonical(employees.values[index][column - 1])) !==
                  JSON.stringify(canonical(inspected.evidence.values[index][0])) ||
                employees.formulas[index][column - 1] !== inspected.evidence.formulas[index][0]) fail();
          }
        }
        columnEvidence.push(inspected.evidence);
        if (employees.values[0][column - 1] === '' || column > employees.lastColumn)
          inspections.push(inspected.inspection);
      }
      const snapshot = {contractVersion: 1, targetIdentityVerified: true, timeZone: expected.timeZone,
        employees: {values: employees.values, maxRows: employees.maxRows, maxColumns: employees.maxColumns},
        columnInspections: inspections, ledger: ledger ? {sheetId: ledger.id, values: ledger.values} :
          {sheetId: null, values: null}, emptyLedgerRecovery: null, finalVerification: 'UNCONFIRMED',
        externalChecks: Object.fromEntries(checks.map(key => [key, false]))};
      function evidence(sheet) {
        if (!sheet) return null;
        return {sheetId: sheet.id, maxRows: sheet.maxRows, maxColumns: sheet.maxColumns,
          lastRow: sheet.lastRow, lastColumn: sheet.lastColumn, formulas: sheet.formulas};
      }
      return {snapshot: snapshot, evidence: {spreadsheetId: expected.spreadsheetId, inventory: inventory,
        employees: evidence(employees), ledger: evidence(ledger), columns: columnEvidence}};
    }
    const first = readPass(), second = readPass();
    if (JSON.stringify(canonical(first)) !== JSON.stringify(canonical(second))) fail();
    const endedAt = integer(environment.now(), startedAt);
    // Validate raw data independently before handing a finalized snapshot to planner.
    // Virtual headers are confined to a validation copy, never the raw snapshot.
    const candidate = second.snapshot;
    const virtual = candidate.employees.values.map(row => row.map(copyCell));
    for (const column of [35, 42]) {
      if (column > candidate.employees.maxColumns) continue;
      if (virtual[0][column - 1] === '' || column > virtual[0].length) {
        const inspection = candidate.columnInspections.find(item => item.column === column);
        if (!inspection || !inspection.reusable || inspection.reasons.length) fail();
        if (column <= virtual[0].length)
          virtual[0][column - 1] = column === 35 ? 'registration_operation_id' : 'initial_grant_check_target';
      }
    }
    const formatter = value => {
      const parts = new Intl.DateTimeFormat('en', {timeZone: 'Asia/Tokyo', year: 'numeric',
        month: '2-digit', day: '2-digit'}).formatToParts(value);
      const part = name => parts.find(item => item.type === name).value;
      return part('year') + '-' + part('month') + '-' + part('day');
    };
    const projectedEmployees = employeeRegistrationProjectEmployees_(virtual, formatter, 'registration_ready');
    if (candidate.ledger.sheetId !== null) {
      const values = candidate.ledger.values;
      const blank = values.every(row => row.every(cell => cell === ''));
      const prefix = values.length === 1 && values[0].length < EMPLOYEE_REGISTRATION_LEDGER_HEADERS_.length &&
        values[0].every((header, i) => header === EMPLOYEE_REGISTRATION_LEDGER_HEADERS_[i]);
      if (!blank && !prefix) {
        const projectedLedger = employeeRegistrationProjectOperationLedger_(values);
        verifyEmployeeRegistrationLedgerEmployeeIntegrity_(projectedLedger, projectedEmployees);
      } else if (projectedEmployees.employeeRows.some(row => row.registration_operation_id !== '')) fail();
    } else if (projectedEmployees.employeeRows.some(row => row.registration_operation_id !== '')) fail();
    // Invoke only after full acquisition, data validation and drift checks. Never execute the plan.
    const result = planEmployeeRegistrationSchemaMigration_(second.snapshot);
    // A verified empty ledger remains a manual diagnostic, never recovery permission.
    // All other semantic anomalies stop acquisition; no snapshot escapes.
    if (result.issues.some(code => code !== 'EMPTY_LEDGER_RECOVERY_UNPROVEN')) fail();
    return {snapshot: second.snapshot,
      evidence: {startedAt: startedAt, endedAt: endedAt, observedEqual: true, details: second.evidence},
      diagnostics: {codes: ['OBSERVED_READS_EQUAL', 'EXTERNAL_CHECKS_UNCONFIRMED', 'NOT_TRANSACTIONAL'],
        issues: result.issues.slice()},
      safeSummary: {state: result.state, employeeReadRows: second.snapshot.employees.values.length,
        employeeReadColumns: second.snapshot.employees.values[0].length,
        inspectedColumns: second.snapshot.columnInspections.map(item => item.column),
        ledgerPresent: second.snapshot.ledger.sheetId !== null, externalChecksConfirmed: 0}};
  } catch (ignored) {
    // Do not propagate message, stack or cause from APIs, callbacks or raw data.
    throw new Error('EMPLOYEE_REGISTRATION_MIGRATION_READ_FAILED');
  }
}
