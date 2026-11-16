import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { AlignmentType, BorderStyle, Document, ImageRun, PageOrientation, Packer, Paragraph, SectionType, Table, TableCell, TableRow, VerticalAlign, WidthType } from 'docx';
import { saveAs } from 'file-saver';
import { jsPDF } from 'jspdf';
import { ChevronLeft, ChevronRight, FileText, Grid2X2, Images, Info, Printer, RotateCcw, ShieldCheck, Upload, X } from 'lucide-react';

const PAGE = { width: 297, height: 210 };
const PHOTO = { width: 35, height: 46.5 };
const GRID = { cols: 7, rows: 4, columnGap: 3, rowGap: 1, margin: 10.5 };
const SLOTS_PER_PAGE = GRID.cols * GRID.rows;
const GRID_WIDTH = GRID.cols * PHOTO.width + (GRID.cols - 1) * GRID.columnGap;
const GRID_X = (PAGE.width - GRID_WIDTH) / 2;
const ACCEPTED = ['image/jpeg', 'image/png', 'image/webp'];
const mmToPx = mm => mm * 96 / 25.4;
const mmToTwips = mm => Math.round(mm * 1440 / 25.4);
const mmToEmuPx = mm => mm * 96 / 25.4;
const formatSize = bytes => bytes < 1024 * 1024 ? `${Math.max(1, Math.round(bytes / 1024))} KB` : `${(bytes / 1024 / 1024).toFixed(1)} MB`;

function buildPages(photos) {
  const count = Math.max(1, Math.ceil(photos.length / SLOTS_PER_PAGE));
  return Array.from({ length: count }, (_, pageIndex) => Array.from({ length: SLOTS_PER_PAGE }, (_, slotIndex) => {
    const index = pageIndex * SLOTS_PER_PAGE + slotIndex;
    const row = Math.floor(slotIndex / GRID.cols);
    const col = slotIndex % GRID.cols;
    return {
      photo: photos[index] || null,
      row,
      col,
      x: GRID_X + col * (PHOTO.width + GRID.columnGap),
      y: GRID.margin + row * (PHOTO.height + GRID.rowGap),
    };
  }));
}

function cropCanvas(photo, fit, width = 1050, height = 1395) {
  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext('2d');
  context.fillStyle = '#fff';
  context.fillRect(0, 0, width, height);
  const image = photo.image;
  const scale = fit === 'fill'
    ? Math.max(width / image.naturalWidth, height / image.naturalHeight)
    : Math.min(width / image.naturalWidth, height / image.naturalHeight);
  const drawWidth = image.naturalWidth * scale;
  const drawHeight = image.naturalHeight * scale;
  const x = (width - drawWidth) * photo.posX / 100;
  const y = (height - drawHeight) * photo.posY / 100;
  context.drawImage(image, x, y, drawWidth, drawHeight);
  return canvas;
}

function canvasBlob(canvas) {
  return new Promise((resolve, reject) => canvas.toBlob(blob => blob ? resolve(blob) : reject(new Error('Could not process this photo.')), 'image/jpeg', 0.97));
}

function Preview({ photos, pages, activePage, setActivePage, fit }) {
  const wrapRef = useRef(null);
  const [scale, setScale] = useState(1);
  const previewPage = pages[activePage] || pages[0];

  useEffect(() => {
    if (!wrapRef.current) return undefined;
    const resize = () => setScale(Math.min((wrapRef.current.clientWidth - 30) / mmToPx(PAGE.width), 1));
    resize();
    const observer = new ResizeObserver(resize);
    observer.observe(wrapRef.current);
    return () => observer.disconnect();
  }, []);

  return <div className="preview-stage" ref={wrapRef} style={{ height: mmToPx(PAGE.height) * scale + 26 }}>
    <div className="paper-shadow" style={{ width: mmToPx(PAGE.width) * scale, height: mmToPx(PAGE.height) * scale }}>
      <div className="preview-paper" style={{ width: mmToPx(PAGE.width), height: mmToPx(PAGE.height), transform: `scale(${scale})` }}>
        {previewPage.map((slot, index) => <div className={`preview-photo ${slot.photo ? '' : 'empty'}`} key={`${activePage}-${index}`} style={{ left: mmToPx(slot.x), top: mmToPx(slot.y), width: mmToPx(PHOTO.width), height: mmToPx(PHOTO.height) }}>
          {slot.photo && <img src={slot.photo.url} alt={slot.photo.file.name} style={{ objectFit: fit === 'fill' ? 'cover' : 'contain', objectPosition: `${slot.photo.posX}% ${slot.photo.posY}%` }} />}
          {!slot.photo && <span>{String(index + 1).padStart(2, '0')}</span>}
        </div>)}
      </div>
    </div>
  </div>;
}

function PrintPages({ pages, fit }) {
  return <div className="print-only" aria-hidden="true">
    {pages.map((page, pageIndex) => <section className="print-sheet" key={`print-${pageIndex}`}>
      {page.map((slot, index) => <div className="print-photo" key={`print-${pageIndex}-${index}`}>
        {slot.photo && <img src={slot.photo.url} alt="" style={{ objectFit: fit === 'fill' ? 'cover' : 'contain', objectPosition: `${slot.photo.posX}% ${slot.photo.posY}%` }} />}
      </div>)}
    </section>)}
  </div>;
}

export default function App() {
  const [photos, setPhotos] = useState([]);
  const [fit, setFit] = useState('fill');
  const [activePage, setActivePage] = useState(0);
  const [dragging, setDragging] = useState(false);
  const [busy, setBusy] = useState('');
  const [notice, setNotice] = useState('');
  const inputRef = useRef(null);
  const idRef = useRef(0);
  const photosRef = useRef(photos);
  photosRef.current = photos;

  const pages = useMemo(() => buildPages(photos), [photos]);
  const pageCount = pages.length;

  useEffect(() => () => photosRef.current.forEach(photo => URL.revokeObjectURL(photo.url)), []);
  useEffect(() => { if (activePage >= pageCount) setActivePage(Math.max(0, pageCount - 1)); }, [activePage, pageCount]);

  const addPhotos = useCallback(async files => {
    const incoming = Array.from(files || []);
    if (!incoming.length) return;
    const valid = [];
    const issues = [];
    for (const file of incoming) {
      if (!ACCEPTED.includes(file.type) && !/\.(jpe?g|png|webp)$/i.test(file.name)) {
        issues.push(`${file.name} is not a JPG, PNG, or WEBP photo.`);
        continue;
      }
      if (file.size > 40 * 1024 * 1024) {
        issues.push(`${file.name} is larger than 40 MB.`);
        continue;
      }
      const url = URL.createObjectURL(file);
      const image = new Image();
      image.src = url;
      try {
        await new Promise((resolve, reject) => { image.onload = resolve; image.onerror = reject; });
        valid.push({ id: ++idRef.current, file, url, image, posX: 50, posY: 50 });
      } catch {
        URL.revokeObjectURL(url);
        issues.push(`Could not read ${file.name}.`);
      }
    }
    if (valid.length) setPhotos(current => [...current, ...valid]);
    setNotice(issues.length ? issues.join(' ') : '');
    if (inputRef.current) inputRef.current.value = '';
  }, []);

  const removePhoto = id => setPhotos(current => {
    const removed = current.find(photo => photo.id === id);
    if (removed) URL.revokeObjectURL(removed.url);
    return current.filter(photo => photo.id !== id);
  });
  const updatePhoto = (id, patch) => setPhotos(current => current.map(photo => photo.id === id ? { ...photo, ...patch } : photo));
  const resetPhotos = () => {
    photos.forEach(photo => URL.revokeObjectURL(photo.url));
    setPhotos([]);
    setActivePage(0);
    setFit('fill');
    setNotice('');
  };

  const exportWord = async () => {
    if (!photos.length) return;
    setBusy('word');
    setNotice('');
    try {
      const sections = [];
      const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
      const columnWidths = [...Array(GRID.cols - 1).fill(mmToTwips(PHOTO.width + GRID.columnGap)), mmToTwips(PHOTO.width)];
      const tableWidth = mmToTwips(GRID_WIDTH);
      for (const page of pages) {
        const rows = [];
        for (let rowIndex = 0; rowIndex < GRID.rows; rowIndex++) {
          const cells = [];
          for (let colIndex = 0; colIndex < GRID.cols; colIndex++) {
            const slot = page[rowIndex * GRID.cols + colIndex];
            const children = [];
            if (slot.photo) {
              const blob = await canvasBlob(cropCanvas(slot.photo, fit));
              children.push(new Paragraph({
                alignment: AlignmentType.LEFT,
                spacing: { before: 0, after: 0, line: 1 },
                children: [new ImageRun({
                  data: new Uint8Array(await blob.arrayBuffer()),
                  transformation: { width: mmToEmuPx(PHOTO.width), height: mmToEmuPx(PHOTO.height) },
                })],
              }));
            } else {
              children.push(new Paragraph({ spacing: { before: 0, after: 0 }, children: [] }));
            }
            cells.push(new TableCell({
              width: { size: columnWidths[colIndex], type: WidthType.DXA },
              margins: { top: 0, bottom: 0, left: 0, right: 0 },
              verticalAlign: VerticalAlign.TOP,
              borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder },
              children,
            }));
          }
          const rowHeight = rowIndex === GRID.rows - 1 ? PHOTO.height : PHOTO.height + GRID.rowGap;
          rows.push(new TableRow({ height: { value: mmToTwips(rowHeight), rule: 'exact' }, cantSplit: true, children: cells }));
        }
        const table = new Table({
          width: { size: tableWidth, type: WidthType.DXA },
          columnWidths,
          layout: 'fixed',
          alignment: AlignmentType.CENTER,
          borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder, insideHorizontal: noBorder, insideVertical: noBorder },
          rows,
        });
        sections.push({
          properties: {
            type: SectionType.NEXT_PAGE,
            page: {
              size: { width: 16838, height: 11906, orientation: PageOrientation.LANDSCAPE },
              margin: { top: mmToTwips(GRID.margin), bottom: mmToTwips(GRID.margin), left: mmToTwips(GRID.margin), right: mmToTwips(GRID.margin), header: 720, footer: 720 },
            },
          },
          children: [table],
        });
      }
      const document = new Document({ sections });
      saveAs(await Packer.toBlob(document), 'psis-photo-booth-a4.docx');
    } catch (error) {
      console.error(error);
      setNotice(`Word export failed: ${error?.message || 'please try again.'}`);
    } finally {
      setBusy('');
    }
  };

  const exportPdf = async () => {
    if (!photos.length) return;
    setBusy('pdf');
    setNotice('');
    try {
      const pdf = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4', compress: true });
      for (let pageIndex = 0; pageIndex < pages.length; pageIndex++) {
        if (pageIndex) pdf.addPage('a4', 'landscape');
        for (const slot of pages[pageIndex]) {
          if (!slot.photo) continue;
          const canvas = cropCanvas(slot.photo, fit);
          pdf.addImage(canvas.toDataURL('image/jpeg', 0.97), 'JPEG', slot.x, slot.y, PHOTO.width, PHOTO.height, undefined, 'FAST');
        }
      }
      pdf.save('psis-photo-booth-a4.pdf');
    } catch (error) {
      console.error(error);
      setNotice(`PDF export failed: ${error?.message || 'please try again.'}`);
    } finally {
      setBusy('');
    }
  };

  return <div className="app-shell">
    <header className="topbar">
      <a className="brand" href="#top">
        <span className="brand-emblem"><img src="/logo-psis.webp" alt="Paññāsāstra International School seal" /></span>
        <span className="brand-wordmark"><strong>PSIS</strong><small>PHOTO BOOTH</small></span>
        <span className="brand-divider" />
        <span className="brand-caption">Paññāsāstra International School</span>
      </a>
      <div className="topbar-right"><span className="privacy"><ShieldCheck size={15} /> Photos stay on this device</span><span className="top-divider" /><span className="paper-pill">A4 · LANDSCAPE</span></div>
    </header>

    <main id="top">
      <section className="hero">
        <div>
          <div className="eyebrow"><span className="eyebrow-dot" /> STUDENT PHOTO SHEETS</div>
          <h1>Every student,<br className="desktop-break" /> <span className="hero-emphasis">print-ready.</span></h1>
          <p>Arrange class photos into a reference-matched A4 sheet.<br className="desktop-break" /> 28 portraits per page, at the right size.</p>
        </div>
        <div className="hero-art" aria-hidden="true"><div className="hero-art-back" /><div className="hero-sheet"><div className="mini-grid">{Array.from({ length: 28 }, (_, i) => <i key={i} />)}</div></div><span className="hero-chip">7 × 4 <small>CLASS SHEET</small></span><span className="hero-sparkle">✳</span><span className="hero-orbit orbit-one" /><span className="hero-orbit orbit-two" /></div>
      </section>

      <div className="workspace">
        <section className="left-column">
          <div className={`upload-card ${dragging ? 'is-dragging' : ''}`} onDragOver={event => { event.preventDefault(); setDragging(true); }} onDragLeave={event => { if (!event.currentTarget.contains(event.relatedTarget)) setDragging(false); }} onDrop={event => { event.preventDefault(); setDragging(false); addPhotos(event.dataTransfer.files); }}>
            <input ref={inputRef} type="file" accept="image/jpeg,image/png,image/webp,.jpg,.jpeg,.png,.webp" multiple hidden onChange={event => addPhotos(event.target.files)} />
            <div className="upload-symbol"><Upload size={19} /></div>
            <h2>Drop student photos here</h2>
            <p>or browse your device to add one or more photos</p>
            <button className="browse-button" onClick={() => inputRef.current?.click()}><Images size={15} /> Browse photos</button>
            <span className="file-note">JPG, JPEG, PNG, WEBP <span>·</span> up to 40 MB each</span>
          </div>

          <div className="gallery-card">
            <div className="section-heading"><div><span className="section-icon"><Images size={16} /></span><h2>Student photos <span className="count-pill">{photos.length}</span></h2></div>{photos.length > 0 && <button className="reset-button" onClick={resetPhotos}><RotateCcw size={13} /> Reset all photos</button>}</div>
            {!photos.length ? <div className="gallery-empty"><span className="empty-stack"><Images size={19} /></span><span>No photos added yet</span><small>Your photos will appear here before printing</small></div> : <div className="photo-list">{photos.map((photo, index) => <article className="photo-row" key={photo.id}>
              <img className="thumb" src={photo.url} alt={photo.file.name} />
              <div className="photo-info"><div className="photo-name" title={photo.file.name}>{String(index + 1).padStart(2, '0')} · {photo.file.name}</div><div className="photo-meta">{photo.image.naturalWidth} × {photo.image.naturalHeight} px <span>·</span> {formatSize(photo.file.size)}</div><div className="reposition-controls"><label>X</label><input aria-label={`Horizontal crop position for ${photo.file.name}`} type="range" min="0" max="100" value={photo.posX} onChange={event => updatePhoto(photo.id, { posX: +event.target.value })} /><label>Y</label><input aria-label={`Vertical crop position for ${photo.file.name}`} type="range" min="0" max="100" value={photo.posY} onChange={event => updatePhoto(photo.id, { posY: +event.target.value })} /></div></div>
              <button className="remove-photo" aria-label={`Remove ${photo.file.name}`} onClick={() => removePhoto(photo.id)}><X size={15} /></button>
            </article>)}</div>}
          </div>
          <div className="privacy-note"><ShieldCheck size={15} /><span><strong>Private by design.</strong> Photos are processed locally and never uploaded.</span></div>
        </section>

        <section className="right-column">
          <div className="settings-card">
            <div className="section-heading settings-heading"><div><span className="section-icon"><Grid2X2 size={16} /></span><h2>Sheet settings</h2></div><span className="reference-pill">MATCHES PSIS LAYOUT</span></div>
            <div className="setting-row"><div className="setting-label"><span>Paper</span><small>Standard A4 landscape</small></div><div className="static-choice"><span className="landscape-glyph" /><strong>A4 · Landscape</strong><span className="muted">297 × 210 mm</span></div></div>
            <div className="setting-row"><div className="setting-label"><span>Photo size</span><small>Exact print dimensions</small></div><div className="size-chips"><strong>3.5 × 4.65 cm</strong><span>W × H</span></div></div>
            <div className="setting-row"><div className="setting-label"><span>Page arrangement</span><small>Consistent with the reference</small></div><div className="grid-setting"><strong>7 across</strong><span>×</span><strong>4 down</strong></div></div>
            <div className="setting-row"><div className="setting-label"><span>Photo adjustment</span><small>Fit keeps the whole photo; Fill crops</small></div><div className="segmented"><button className={fit === 'fill' ? 'active' : ''} onClick={() => setFit('fill')}>Fill</button><button className={fit === 'fit' ? 'active' : ''} onClick={() => setFit('fit')}>Fit</button></div></div>
            <div className="capacity"><div className="capacity-main"><span className="capacity-icon"><Grid2X2 size={15} /></span><span><strong>28</strong><small> photo slots per page</small></span></div><span className="capacity-side">{photos.length} photos · {pageCount} {pageCount === 1 ? 'page' : 'pages'}</span></div>
          </div>

          <div className="preview-card">
            <div className="preview-header"><div><span className="section-icon preview-icon"><FileText size={15} /></span><h2>A4 layout preview</h2><span className="live-badge"><i /> LIVE</span></div><div className="page-switcher"><button aria-label="Previous page" disabled={activePage === 0} onClick={() => setActivePage(value => Math.max(0, value - 1))}><ChevronLeft size={15} /></button><span>Page <strong>{activePage + 1}</strong> of {pageCount}</span><button aria-label="Next page" disabled={activePage >= pageCount - 1} onClick={() => setActivePage(value => Math.min(pageCount - 1, value + 1))}><ChevronRight size={15} /></button></div></div>
            <Preview photos={photos} pages={pages} activePage={activePage} setActivePage={setActivePage} fit={fit} />
            <div className="preview-footer"><span><span className="orientation-mini" /> A4 · Landscape</span><span>{Math.min(photos.length - activePage * SLOTS_PER_PAGE, SLOTS_PER_PAGE) > 0 ? `${Math.min(photos.length - activePage * SLOTS_PER_PAGE, SLOTS_PER_PAGE)} of 28 photos` : '28 photo slots'}</span></div>
          </div>

          <div className="export-card">
            <div className="export-heading"><span>Ready to print?</span><small>Choose Actual Size / 100% on A4 landscape paper.</small></div>
            <button className="export-primary" disabled={!photos.length || !!busy} onClick={exportWord}>{busy === 'word' ? <span className="spinner" /> : <FileText size={16} />}<span>{busy === 'word' ? 'Preparing Word file…' : 'Download Word'}</span><span className="file-badge">.DOCX</span></button>
            <div className="export-secondary"><button disabled={!photos.length || !!busy} onClick={exportPdf}>{busy === 'pdf' ? <span className="spinner" /> : <FileText size={14} />} Download PDF</button><button disabled={!photos.length || !!busy} onClick={() => window.print()}><Printer size={14} /> Print now</button></div>
          </div>
        </section>
      </div>

      {notice && <div className="toast" role="status"><Info size={15} /><span>{notice}</span><button aria-label="Dismiss" onClick={() => setNotice('')}><X size={14} /></button></div>}
      <footer className="footer"><span>PSIS Photo Booth</span><span><ShieldCheck size={13} /> Print on A4 landscape at 100% Actual Size</span></footer>
    </main>
    <PrintPages pages={pages} fit={fit} />
  </div>;
}
