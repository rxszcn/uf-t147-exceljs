'use strict';

const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

// Fault-tolerance contract for packages produced by heterogeneous systems
// whose internal part names or relationship wiring do not follow the
// conventional layout. Reading must neither silently drop data nor crash on
// an uncontrolled TypeError; every tolerated anomaly is exposed on
// `workbook.xlsx.readWarnings`.
describe('Malformed package fault tolerance', () => {
  const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="S1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;

  const workbookRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`;

  const worksheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>hello</t></is></c></row>
  </sheetData>
</worksheet>`;

  const rootRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

  const baseEntries = () => ({
    '_rels/.rels': rootRelsXml,
    'xl/workbook.xml': workbookXml,
    'xl/_rels/workbook.xml.rels': workbookRelsXml,
    'xl/worksheets/sheet1.xml': worksheetXml,
  });

  async function load(entries) {
    const zip = new JSZip();
    Object.entries(entries).forEach(([name, content]) => {
      zip.file(name, content);
    });
    const buffer = await zip.generateAsync({type: 'nodebuffer'});
    const workbook = new ExcelJS.Workbook();
    return workbook.xlsx.load(buffer);
  }

  describe('a conventionally laid out package', () => {
    it('reads every worksheet and collects no warnings', async () => {
      const workbook = await load(baseEntries());
      expect(workbook.worksheets).to.have.length(1);
      expect(workbook.getWorksheet(1).getCell('A1').value).to.equal('hello');
      expect(workbook.xlsx.readWarnings.count).to.equal(0);
      expect(workbook.xlsx.readWarnings.toArray()).to.deep.equal([]);
    });
  });

  describe('a worksheet part whose entry name casing is non-standard', () => {
    it('keeps the worksheet data and records an entry-name-case warning', async () => {
      const entries = baseEntries();
      delete entries['xl/worksheets/sheet1.xml'];
      entries['xl/Worksheets/sheet1.xml'] = worksheetXml;

      const workbook = await load(entries);
      expect(workbook.worksheets).to.have.length(1);
      expect(workbook.getWorksheet(1).name).to.equal('S1');
      expect(workbook.getWorksheet(1).getCell('A1').value).to.equal('hello');

      const warnings = workbook.xlsx.readWarnings.toArray();
      expect(warnings).to.have.length(1);
      expect(warnings[0].code).to.equal('entry-name-case');
      expect(warnings[0].entryName).to.equal('xl/Worksheets/sheet1.xml');
      expect(warnings[0].canonicalName).to.equal('xl/worksheets/sheet1.xml');
    });

    it('tolerates case deviations on fixed-name parts as well', async () => {
      const entries = baseEntries();
      delete entries['xl/workbook.xml'];
      entries['xl/WorkBook.xml'] = workbookXml;

      const workbook = await load(entries);
      expect(workbook.worksheets).to.have.length(1);
      const warnings = workbook.xlsx.readWarnings.toArray();
      expect(warnings.some(w => w.code === 'entry-name-case' && w.entryName === 'xl/WorkBook.xml')).to.be.true();
    });

    it('tolerates deviations in the file name itself, not just the directory', async () => {
      const entries = baseEntries();
      delete entries['xl/worksheets/sheet1.xml'];
      entries['xl/worksheets/Sheet1.XML'] = worksheetXml;

      const workbook = await load(entries);
      expect(workbook.worksheets).to.have.length(1);
      expect(workbook.getWorksheet(1).getCell('A1').value).to.equal('hello');
      const warnings = workbook.xlsx.readWarnings.toArray();
      expect(warnings.some(w => w.code === 'entry-name-case' && w.entryName === 'xl/worksheets/Sheet1.XML')).to
        .be.true();
    });
  });

  describe('a worksheet relationship dangling at a missing comments part', () => {
    const danglingCommentsRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>
</Relationships>`;

    it('resolves with the worksheet intact instead of throwing', async () => {
      const entries = baseEntries();
      entries['xl/worksheets/_rels/sheet1.xml.rels'] = danglingCommentsRels;

      const workbook = await load(entries);
      expect(workbook.worksheets).to.have.length(1);
      expect(workbook.getWorksheet(1).getCell('A1').value).to.equal('hello');

      const warnings = workbook.xlsx.readWarnings.toArray();
      expect(warnings).to.have.length(1);
      expect(warnings[0].code).to.equal('dangling-relationship');
      expect(warnings[0].target).to.equal('../comments1.xml');
      expect(warnings[0].reason).to.equal('comments-part-missing');
    });
  });

  describe('a worksheet relationship dangling at a missing VML drawing part', () => {
    const danglingVmlRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
</Relationships>`;

    it('resolves and records a dangling-relationship warning', async () => {
      const entries = baseEntries();
      entries['xl/worksheets/_rels/sheet1.xml.rels'] = danglingVmlRels;

      const workbook = await load(entries);
      expect(workbook.worksheets).to.have.length(1);
      const warnings = workbook.xlsx.readWarnings.toArray();
      expect(warnings.some(w => w.code === 'dangling-relationship' && w.reason === 'vml-drawing-part-missing')).to
        .be.true();
    });
  });

  describe('an unrecognised zip entry', () => {
    it('stays readable and records an unrecognised-entry warning', async () => {
      const entries = baseEntries();
      entries['xl/vendor-extension/payload.bin'] = 'opaque';

      const workbook = await load(entries);
      expect(workbook.worksheets).to.have.length(1);
      const warnings = workbook.xlsx.readWarnings.toArray();
      expect(warnings).to.have.length(1);
      expect(warnings[0].code).to.equal('unrecognised-entry');
      expect(warnings[0].entryName).to.equal('xl/vendor-extension/payload.bin');
    });

    it('stays silent for conventional package parts it does not consume', async () => {
      const entries = baseEntries();
      entries['[Content_Types].xml'] =
        '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>';
      entries['docProps/custom.xml'] = '<?xml version="1.0"?><Properties/>';

      const workbook = await load(entries);
      expect(workbook.xlsx.readWarnings.count).to.equal(0);
    });
  });

  describe('a workbook catalogue entry pointing at a missing worksheet part', () => {
    it('loads the remaining workbook and records a missing-part warning', async () => {
      const entries = baseEntries();
      delete entries['xl/worksheets/sheet1.xml'];

      const workbook = await load(entries);
      expect(workbook.worksheets).to.have.length(0);
      const warnings = workbook.xlsx.readWarnings.toArray();
      expect(warnings).to.have.length(1);
      expect(warnings[0].code).to.equal('missing-part');
      expect(warnings[0].partName).to.equal('xl/worksheets/sheet1.xml');
      expect(warnings[0].sheetName).to.equal('S1');
    });
  });

  describe('a workbook catalogue entry with a missing relationship', () => {
    it('skips the sheet and records a dangling-relationship warning', async () => {
      const entries = baseEntries();
      entries['xl/_rels/workbook.xml.rels'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;

      const workbook = await load(entries);
      expect(workbook.worksheets).to.have.length(0);
      const warnings = workbook.xlsx.readWarnings.toArray();
      expect(warnings).to.have.length(1);
      expect(warnings[0].code).to.equal('dangling-relationship');
      expect(warnings[0].relId).to.equal('rId1');
    });
  });

  describe('a package missing the workbook part (missing-bits contract)', () => {
    it('resolves with zero worksheets and no warnings', async () => {
      const entries = baseEntries();
      delete entries['xl/workbook.xml'];
      delete entries['xl/_rels/workbook.xml.rels'];

      const workbook = await load(entries);
      expect(workbook.worksheets).to.have.length(0);
      expect(workbook.xlsx.readWarnings.count).to.equal(0);
    });
  });

  describe('the warning collection', () => {
    it('is iterable and exposes a stable snapshot', async () => {
      const entries = baseEntries();
      entries['xl/vendor-extension/payload.bin'] = 'opaque';
      const workbook = await load(entries);

      const codes = [];
      for (const warning of workbook.xlsx.readWarnings) {
        codes.push(warning.code);
      }
      expect(codes).to.deep.equal(['unrecognised-entry']);

      const snapshot = workbook.xlsx.readWarnings.toArray();
      snapshot.push({code: 'tampered'});
      expect(workbook.xlsx.readWarnings.count).to.equal(1);
    });
  });
});
