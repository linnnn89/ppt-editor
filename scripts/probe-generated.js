import pptxgen from 'pptxgenjs';
import { createRequire } from 'node:module';
import { promises as fs } from 'node:fs';
import path from 'node:path';
import { buildDeck } from '../src/build.js';
import { readPackage, writePackage } from '../src/ooxml.js';
const root = path.resolve('work/generator-probe'); await fs.mkdir(root, { recursive: true });
const original = new pptxgen(); original.addSlide().addText('Original', { x: 1, y: 1, w: 4, h: 1 });
const raw = await original.write({ outputType: 'nodebuffer' });
const variants = { raw, repacked: await writePackage(await readPackage(raw)), generated: await buildDeck({ slides: [{ items: [{ type: 'text', text: 'Original', left: 72, top: 72, width: 300, height: 72 }] }] }) };
const rawParts = await readPackage(raw), generatedParts = await readPackage(variants.generated);
const restoreProperties = new Map(generatedParts);
for (const name of ['docProps/app.xml', 'docProps/core.xml']) restoreProperties.set(name, rawParts.get(name));
variants.restoredProperties = await writePackage(restoreProperties);
const restoreSlide = new Map(generatedParts); restoreSlide.set('ppt/slides/slide1.xml', rawParts.get('ppt/slides/slide1.xml'));
variants.restoredSlide = await writePackage(restoreSlide);
const w = createRequire(import.meta.url)('winax'), app = new w.Object('PowerPoint.Application');
try {
  for (const [name, bytes] of Object.entries(variants)) {
    const file = path.join(root, `${name}.pptx`); await fs.writeFile(file, bytes);
    try { const pres = app.Presentations.Open(file, 0, 0, 0); console.log(JSON.stringify({ name, opened: true, slides: Number(pres.Slides.Count) })); pres.Close(); w.release(pres); }
    catch (error) { console.log(JSON.stringify({ name, error: error.message })); }
  }
} finally { w.release(app); }
