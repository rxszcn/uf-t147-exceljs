'use strict';

// Read-time fault-tolerance contract.
//
// Heterogeneous producers sometimes emit packages whose internal part names
// or relationship wiring do not follow the conventional layout. Reading such
// a package must neither silently drop data nor crash on an uncontrolled
// TypeError. Every tolerated anomaly is recorded here as a structured warning
// so callers can inspect programmatically what was degraded.
//
// Warning codes:
//   'unrecognised-entry'            a zip entry that is not part of the
//                                   supported spreadsheet layout
//   'entry-name-case'               a recognised part whose path casing does
//                                   not match the conventional part name
//   'missing-part'                  a part referenced by the package
//                                   metadata is absent from the container
//   'dangling-relationship'         a relationship Target cannot be resolved
//                                   to a parsed part
//
// All warnings share the shape {code, message, ...details}.

class ReadWarnings {
  constructor() {
    this.warnings = [];
  }

  add(code, message, details) {
    const warning = Object.assign({code, message}, details);
    this.warnings.push(warning);
    return warning;
  }

  unrecognisedEntry(entryName) {
    return this.add(
      'unrecognised-entry',
      `Skipped unrecognised package entry: ${entryName}`,
      {entryName}
    );
  }

  entryNameCase(entryName, canonicalName) {
    // The same fold can be observed twice (entry dispatch and workbook
    // reconciliation); keep a single warning per folded entry name.
    if (this.warnings.some(warning => warning.code === 'entry-name-case' && warning.entryName === entryName)) {
      return this.warnings.find(warning => warning.code === 'entry-name-case' && warning.entryName === entryName);
    }
    return this.add(
      'entry-name-case',
      `Part name "${entryName}" differs in case from conventional name "${canonicalName}"; matched case-insensitively`,
      {entryName, canonicalName}
    );
  }

  missingPart(partName, details) {
    return this.add('missing-part', `Missing package part: ${partName}`, Object.assign({partName}, details));
  }

  danglingRelationship(details) {
    const source = details && details.source ? `${details.source} ` : '';
    const target = details && details.target;
    return this.add(
      'dangling-relationship',
      `${source}relationship "${details && details.relId}" targets missing part: ${target}`,
      details
    );
  }

  get count() {
    return this.warnings.length;
  }

  toArray() {
    // return a shallow copy so callers cannot mutate internal state
    return this.warnings.slice();
  }

  [Symbol.iterator]() {
    return this.warnings[Symbol.iterator]();
  }
}

module.exports = ReadWarnings;
