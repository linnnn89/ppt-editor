import assert from 'node:assert/strict';
import { promises as fs } from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { randomUUID } from 'node:crypto';
import { readPackage } from '../src/ooxml.js';
import { indexPackage } from '../src/file-engine.js';
import { hash, inside, writeJson } from '../src/storage.js';

// An independent Windows JScript diagnostic, not an alternative production backend.
const source = path.resolve(process.argv[2] || 'work/native-verification/085d246a-bdae-44db-a95f-09a5a5d600d6/source.pptx');
const savedCheckpoint = path.join(path.dirname(source), 'native-1-0.pptx');
assert.ok(inside(process.cwd(), await fs.realpath(source)), 'Diagnostic input must be within the project');
const root = path.resolve('work/native-diagnosis', randomUUID());
await fs.mkdir(root, { recursive: true });
const original = await fs.readFile(source), checkpoint = await fs.readFile(savedCheckpoint);
await readPackage(original); await readPackage(checkpoint);
await fs.writeFile(path.join(root, 'working.pptx'), original);
await fs.writeFile(path.join(root, 'previous-checkpoint.pptx'), checkpoint);
const script = String.raw`
var root = WScript.Arguments.Item(0), app = null, pres = null;
function verify(value, message) { if (!value) throw new Error(message); }
function open(file) {
  var p = app.Presentations.Open(file, 0, 0, 0);
  verify(String(p.FullName).toLowerCase() === file.toLowerCase(), "Document identity mismatch");
  return p;
}
function closeCopy() { if (pres) { pres.Saved = -1; pres.Close(); pres = null; } }
try {
  app = new ActiveXObject("PowerPoint.Application");
  var beforeCount = app.Presentations.Count;
  pres = open(root + "\\working.pptx");
  var slide = pres.Slides.Item(1), heading = slide.Shapes.Item("Heading"), table = null;
  var expectedText = String(heading.TextFrame.TextRange.Text).replace("Original", "Revised");
  heading.TextFrame.TextRange.Text = expectedText;
  heading.TextFrame.TextRange.Font.Size = 28;
  heading.Left = 80;
  for (var i = 1; i <= slide.Shapes.Count; i++) if (slide.Shapes.Item(i).HasTable === -1) table = slide.Shapes.Item(i).Table;
  verify(table !== null, "Missing table");
  table.Cell(2, 2).Shape.TextFrame.TextRange.Text = "18";
  pres.SaveCopyAs(root + "\\result.pptx", 24);
  table = null; heading = null; slide = null;
  closeCopy();
  pres = open(root + "\\result.pptx");
  heading = pres.Slides.Item(1).Shapes.Item("Heading");
  verify(String(heading.TextFrame.TextRange.Text) === expectedText, "Text readback failed");
  verify(Number(heading.TextFrame.TextRange.Font.Size) === 28, "Font readback failed");
  verify(Number(heading.Left) === 80, "Geometry readback failed");
  heading = null;
  closeCopy();
  pres = open(root + "\\previous-checkpoint.pptx");
  verify(Number(pres.Slides.Count) === 1, "Checkpoint readback failed");
  closeCopy();
  verify(app.Presentations.Count === beforeCount, "Presentation count changed");
  WScript.Echo("source-edit-save-reopen:passed;previous-checkpoint-open:passed;document-count:unchanged");
} catch (error) {
  WScript.Echo("FAILED:" + (error.number || 0) + ":" + error.message);
  try { closeCopy(); } catch (cleanupError) { WScript.Echo("CLEANUP_FAILED:" + cleanupError.message); }
  WScript.Quit(1);
} finally { app = null; }
`;
const scriptFile = path.join(root, 'office-control.js');
await fs.writeFile(scriptFile, script);
const report = { root, source, sourceHash: hash(original), checkpointHash: hash(checkpoint), engine: 'Windows JScript diagnostic', structuralInput: 'passed', structuralCheckpoint: 'passed' };
try {
  report.nativeResult = execFileSync('cscript.exe', ['//nologo', '//B', '//T:30', scriptFile, root], { encoding: 'utf8', windowsHide: true, timeout: 35000 }).trim();
  const index = indexPackage(await readPackage(await fs.readFile(path.join(root, 'result.pptx'))));
  const heading = index.objects.find(o => o.name === 'Heading'), table = index.objects.find(o => o.kind === 'table');
  assert.equal(heading.text, 'Revised 中文 😀 text'); assert.equal(heading.geometry.left, 80);
  assert.equal(table.rows[1][1], '18');
  assert.equal(hash(await fs.readFile(source)), report.sourceHash);
  assert.equal(hash(await fs.readFile(savedCheckpoint)), report.checkpointHash);
  report.outputReadback = 'passed'; report.originalFilesUnchanged = true;
} catch (error) {
  report.error = { message: error.message, stdout: error.stdout?.toString(), stderr: error.stderr?.toString() };
  process.exitCode = 1;
} finally {
  await writeJson(path.join(root, 'ppt-report.json'), report);
  console.log(JSON.stringify(report, null, 2));
}
