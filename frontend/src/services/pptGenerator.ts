import PptxGenJS from 'pptxgenjs';
import JSZip from 'jszip';
import { PMGroup, GateGroup, Project } from '../types';
import { VOIS_COLORS, FONTS, FONT_SIZES, PROJECTS_PER_SLIDE, LAYOUT } from '../config/constants';
import { truncate } from '../utils/nameUtils';

type SN = PptxGenJS.SHAPE_NAME;
const RECT = 'rect' as SN;
const LINE = 'line' as SN;
const h = (c: string) => (c.startsWith('#') ? c.slice(1) : c);

// ─── layout helpers ───────────────────────────────────────────────────────────

// Fixed font sizes — identical on every card, never shrink
const F_GATE = 7.5;
const F_PROJ = 7.0;

// Base row heights (natural size — may scale down proportionally when card is compressed)
const GATE_ROW_H  = 0.22;   // gate pill row
const PROJ_ROW_H  = 0.165;  // project bullet row
const HDR_H_TOTAL = 0.05 + 0.20 + 0.04 + 0.06; // topBar + name + divider + content-top-pad
const CARD_PAD_B  = 0.05;   // inner bottom padding

/** Total natural height to show ALL content at base row sizes */
function computeCardHeight(pm: PMGroup): number {
  let h = HDR_H_TOTAL + CARD_PAD_B;
  for (const gate of pm.gates) {
    h += GATE_ROW_H + gate.projects.length * PROJ_ROW_H;
  }
  return h;
}

// ─── chrome ───────────────────────────────────────────────────────────────────

function addHeader(slide: PptxGenJS.Slide, title: string, sub?: string): void {
  const W = LAYOUT.slideW, H = LAYOUT.headerH;

  slide.addShape(RECT, { x: 0, y: 0, w: W, h: H,
    fill: { color: h(VOIS_COLORS.headerBg) }, line: { color: h(VOIS_COLORS.headerBg) } });

  slide.addShape(RECT, { x: 0, y: 0, w: 0.07, h: H,
    fill: { color: h(VOIS_COLORS.accentStripe) }, line: { color: h(VOIS_COLORS.accentStripe) } });

  slide.addText(title, {
    x: 0.14, y: 0, w: W - 1.1, h: H,
    fontSize: FONT_SIZES.slideTitle, bold: true,
    color: h(VOIS_COLORS.white), fontFace: FONTS.title, valign: 'middle',
  });

  if (sub) {
    slide.addText(sub, {
      x: W - 0.95, y: 0, w: 0.80, h: H,
      fontSize: 6.5, color: 'AAAAAA', fontFace: FONTS.body,
      align: 'right', valign: 'middle', italic: true,
    });
  }

}

function addFooter(slide: PptxGenJS.Slide, slideNum: number): void {
  const fy = LAYOUT.footerY, W = LAYOUT.slideW, fH = LAYOUT.footerH;

  slide.addShape(RECT, { x: 0, y: fy, w: W, h: fH,
    fill: { color: h(VOIS_COLORS.footerBg) }, line: { color: h(VOIS_COLORS.footerBg) } });

  slide.addShape(RECT, { x: 0, y: fy, w: W, h: 0.018,
    fill: { color: h(VOIS_COLORS.primary) }, line: { color: h(VOIS_COLORS.primary) } });

  slide.addText('DWS EG PM Team', {
    x: 0.16, y: fy + 0.02, w: W * 0.35, h: fH - 0.04,
    fontSize: FONT_SIZES.footer, color: '888888',
    fontFace: FONTS.body, valign: 'middle',
  });

  // VOIS wordmark — centred in the footer
  slide.addText('VOIS', {
    x: (W - 0.59) / 2, y: fy, w: 0.59, h: fH,
    fontSize: 9.5, bold: true,
    color: h(VOIS_COLORS.primary), fontFace: FONTS.title,
    align: 'center', valign: 'middle',
  });

  slide.addText(`${slideNum}`, {
    x: W - 0.45, y: fy + 0.02, w: 0.32, h: fH - 0.04,
    fontSize: FONT_SIZES.slideNumber, bold: true,
    color: h(VOIS_COLORS.white), fontFace: FONTS.body,
    align: 'right', valign: 'middle',
  });
}

// ─── PM card ──────────────────────────────────────────────────────────────────

function drawPMCard(
  slide: PptxGenJS.Slide,
  pm: PMGroup,
  cx: number, cy: number,
  cardW: number, cardH: number,
): void {
  // shadow
  slide.addShape(RECT, {
    x: cx + 0.022, y: cy + 0.022, w: cardW, h: cardH,
    fill: { color: 'E2E2E2' }, line: { color: 'E2E2E2', pt: 0 },
  });
  // card body
  slide.addShape(RECT, {
    x: cx, y: cy, w: cardW, h: cardH,
    fill: { color: h(VOIS_COLORS.cardBg) },
    line: { color: h(VOIS_COLORS.cardBorder), pt: 0.75 },
  });
  // red top bar
  const TOP = 0.05;
  slide.addShape(RECT, {
    x: cx, y: cy, w: cardW, h: TOP,
    fill: { color: h(VOIS_COLORS.cardTopBar) }, line: { color: h(VOIS_COLORS.cardTopBar) },
  });

  const HDR_H = 0.20;
  const hdrY  = cy + TOP;
  const nameW = cardW - 0.58;

  slide.addText(truncate(pm.pmName, 38), {
    x: cx + 0.09, y: hdrY, w: nameW, h: HDR_H,
    fontSize: FONT_SIZES.cardTitle, bold: true,
    color: h(VOIS_COLORS.pmNameColor), fontFace: FONTS.title, valign: 'middle',
  });

  slide.addShape(RECT, {
    x: cx + cardW - 0.50, y: hdrY + 0.02, w: 0.44, h: HDR_H - 0.04,
    fill: { color: h(VOIS_COLORS.gateTag) },
    line: { color: h(VOIS_COLORS.cardBorder), pt: 0.5 },
  });
  slide.addText(`${pm.totalProjects}p`, {
    x: cx + cardW - 0.50, y: hdrY + 0.02, w: 0.44, h: HDR_H - 0.04,
    fontSize: 6.5, bold: true,
    color: h(VOIS_COLORS.gateTagText), fontFace: FONTS.body,
    align: 'center', valign: 'middle',
  });

  const divY = hdrY + HDR_H;
  slide.addShape(LINE, {
    x: cx + 0.07, y: divY, w: cardW - 0.14, h: 0,
    line: { color: h(VOIS_COLORS.cardBorder), pt: 0.5 },
  });

  // ── content area — all projects always shown, rows scale with card ────
  const contentStartY  = divY + 0.04;
  const textW          = cardW - 0.20;

  // Scale row heights proportionally so ALL content fills exactly this card's height
  const naturalContent = computeCardHeight(pm) - HDR_H_TOTAL - CARD_PAD_B;
  const actualContent  = cardH - HDR_H_TOTAL - CARD_PAD_B;
  const rowScale       = naturalContent > 0 ? actualContent / naturalContent : 1;

  const gateH = GATE_ROW_H * rowScale;
  const projH = PROJ_ROW_H * rowScale;

  let gy = contentStartY;
  for (const gate of pm.gates) {
    // Gate label pill
    slide.addShape(RECT, {
      x: cx + 0.07, y: gy + 0.01,
      w: cardW - 0.14, h: gateH - 0.02,
      fill: { color: h(VOIS_COLORS.gateTag) },
      line: { color: h(VOIS_COLORS.cardBorder), pt: 0.5 },
    });
    slide.addText(`Gate Approved: ${gate.gate} (${gate.projects.length})`, {
      x: cx + 0.10, y: gy + 0.01,
      w: cardW - 0.20, h: gateH - 0.02,
      fontSize: F_GATE, bold: true,
      color: h(VOIS_COLORS.gateTagText), fontFace: FONTS.body, valign: 'middle',
    });
    gy += gateH;

    // Project rows
    for (const proj of gate.projects) {
      slide.addText(`• ${proj.projectName} - ${proj.projectId}`, {
        x: cx + 0.12, y: gy,
        w: textW, h: projH,
        fontSize: F_PROJ,
        color: h(VOIS_COLORS.bodyText), fontFace: FONTS.body, valign: 'middle',
      });
      gy += projH;
    }
  }
}

// ─── summary slide ────────────────────────────────────────────────────────────

function addSummarySlide(pptx: PptxGenJS, pmGroups: PMGroup[], slideNum: number): void {
  const slide = pptx.addSlide();

  slide.addShape(RECT, { x: 0, y: 0, w: LAYOUT.slideW, h: LAYOUT.slideH,
    fill: { color: h(VOIS_COLORS.slideBg) }, line: { color: h(VOIS_COLORS.slideBg) } });

  addHeader(slide, 'Projects Summary - DWS EG PM Team');
  addFooter(slide, slideNum);

  const totalAll = pmGroups.reduce((s, g) => s + g.totalProjects, 0);

  const n     = pmGroups.length;
  const cols  = n <= 1 ? 1 : n <= 2 ? 2 : 3;
  const gapX  = 0.09;
  const gapY  = 0.07;
  const cardW = (LAYOUT.usableW - gapX * (cols - 1)) / cols;
  const availH = LAYOUT.cardsTotalH;

  // Compute natural (content-driven) height for every PM card
  const naturalH = pmGroups.map(pm => computeCardHeight(pm));

  // Greedy shortest-column-first packing
  const colItems: { pm: PMGroup; h: number }[][] = Array.from({ length: cols }, () => []);
  const colH: number[] = new Array(cols).fill(0);

  for (let i = 0; i < pmGroups.length; i++) {
    const shortest = colH.indexOf(Math.min(...colH));
    colItems[shortest].push({ pm: pmGroups[i], h: naturalH[i] });
    colH[shortest] += naturalH[i] + gapY;
  }

  // If the tallest column overflows, scale all heights down uniformly
  const maxColH = Math.max(...colH);
  const scale   = maxColH > availH ? availH / maxColH : 1;

  // Render
  for (let col = 0; col < cols; col++) {
    const cx = LAYOUT.marginL + col * (cardW + gapX);
    let cy    = LAYOUT.cardsStartY;

    for (const { pm, h } of colItems[col]) {
      const cardH = h * scale;
      drawPMCard(slide, pm, cx, cy, cardW, cardH);
      cy += cardH + gapY * scale;
    }
  }

  // Grand total bar
  const tbY = LAYOUT.cardsEndY + LAYOUT.GRAND_TOTAL_GAP;
  slide.addShape(RECT, {
    x: LAYOUT.marginL, y: tbY, w: LAYOUT.usableW, h: LAYOUT.GRAND_TOTAL_H,
    fill: { color: h(VOIS_COLORS.grandTotalBg) }, line: { color: h(VOIS_COLORS.grandTotalBg) },
  });
  slide.addText(
    `GRAND TOTAL   ${totalAll} Projects  ·  ${pmGroups.length} Project Manager${pmGroups.length !== 1 ? 's' : ''}`,
    {
      x: LAYOUT.marginL + 0.18, y: tbY, w: LAYOUT.usableW - 0.3, h: LAYOUT.GRAND_TOTAL_H,
      fontSize: FONT_SIZES.grandTotal, bold: true,
      color: h(VOIS_COLORS.grandTotalTxt), fontFace: FONTS.title, valign: 'middle',
    }
  );
}

// ─── PM detail slides ─────────────────────────────────────────────────────────

interface ProjectBlock { gate: GateGroup; project: Project; }

function flattenProjects(pm: PMGroup): ProjectBlock[] {
  const out: ProjectBlock[] = [];
  for (const gate of pm.gates)
    for (const proj of gate.projects)
      out.push({ gate, project: proj });
  return out;
}

// PM slide row geometry — every line uses the same BOX_H so there's no overlap
const BOX_H    = 0.22;    // height of every text box line
const BOX_STEP = 0.24;    // vertical step between line tops (BOX_H + 0.02 gap)
const PAD_TOP  = 0.02;    // top inner padding inside each project row
const PAD_BOT  = 0.08;    // bottom padding / row gap
// ROW_H = PAD_TOP + 3*BOX_STEP + BOX_H + PAD_BOT
//       = 0.02   + 0.72        + 0.22   + 0.08   = 1.04"
const ROW_H    = PAD_TOP + 3 * BOX_STEP + BOX_H + PAD_BOT;

function addPMSlide(
  pptx: PptxGenJS, pm: PMGroup,
  blocks: ProjectBlock[], pageIndex: number,
  totalPages: number, slideNum: number,
): void {
  const slide = pptx.addSlide();

  slide.addShape(RECT, { x: 0, y: 0, w: LAYOUT.slideW, h: LAYOUT.slideH,
    fill: { color: h(VOIS_COLORS.white) }, line: { color: h(VOIS_COLORS.white) } });

  const suffix = pageIndex > 0 ? ' — Cont.' : '';
  const label  = totalPages > 1 ? `Page ${pageIndex + 1} / ${totalPages}` : undefined;
  addHeader(slide, pm.pmName + suffix, label);
  addFooter(slide, slideNum);

  slide.addShape(RECT, {
    x: 0, y: LAYOUT.headerH, w: 0.05,
    h: LAYOUT.footerY - LAYOUT.headerH,
    fill: { color: h(VOIS_COLORS.sideStrip) }, line: { color: h(VOIS_COLORS.sideStrip) },
  });

  const sy = LAYOUT.contentStartY;

  // Total projects count only
  slide.addText(`Total Projects: ${pm.totalProjects}`, {
    x: LAYOUT.marginL + 0.10, y: sy, w: LAYOUT.usableW - 0.2, h: 0.21,
    fontSize: FONT_SIZES.summaryBar, bold: true,
    color: h(VOIS_COLORS.summaryBarText), fontFace: FONTS.body, valign: 'middle',
  });

  const listStartY = sy + 0.25;
  const listEndY   = LAYOUT.contentEndY;

  // Content starts after the thin red bar (0.05") so offset x by 0.10"
  const tx = LAYOUT.marginL + 0.10;
  const tw = LAYOUT.usableW - 0.10;

  for (let i = 0; i < blocks.length; i++) {
    const { project } = blocks[i];
    const by = listStartY + i * ROW_H;
    if (by + ROW_H > listEndY + 0.01) break;

    // ── Row background (alternating) ──────────────────────────────────
    slide.addShape(RECT, {
      x: LAYOUT.marginL, y: by, w: LAYOUT.usableW, h: ROW_H,
      fill: { color: i % 2 === 0 ? h(VOIS_COLORS.rowEven) : h(VOIS_COLORS.rowOdd) },
      line: { color: h(VOIS_COLORS.divider), pt: 0.25 },
    });

    // ── Red left bar (keep) ───────────────────────────────────────────
    slide.addShape(RECT, {
      x: LAYOUT.marginL, y: by, w: 0.05, h: ROW_H,
      fill: { color: h(VOIS_COLORS.primary) }, line: { color: h(VOIS_COLORS.primary) },
    });

    // ── Line 0: Project name (y = by + PAD_TOP) ───────────────────────
    const line0Y = by + PAD_TOP;
    slide.addText(`• ${project.projectName}`, {
      x: tx, y: line0Y, w: tw, h: BOX_H,
      fontSize: 9, bold: true,
      color: h(VOIS_COLORS.bodyText), fontFace: FONTS.title, valign: 'middle',
    });

    // ── Line 1: OIT Objective (y = by + PAD_TOP + 1*BOX_STEP) ─────────
    const line1Y = by + PAD_TOP + BOX_STEP;
    slide.addShape(RECT, {
      x: tx, y: line1Y, w: tw, h: BOX_H,
      fill: { color: 'FFF5F5' },
      line: { color: h(VOIS_COLORS.divider), pt: 0.3 },
    });
    slide.addText('OIT Objective:', {
      x: tx + 0.05, y: line1Y, w: 1.1, h: BOX_H,
      fontSize: 7, bold: true,
      color: h(VOIS_COLORS.summaryBarText), fontFace: FONTS.body, valign: 'middle',
    });
    // Blank editable input — h: 0.22, w: 6.6 (as requested)
    slide.addText('', {
      x: tx + 1.15, y: line1Y, w: 6.6, h: BOX_H,
      fontSize: 8,
      color: h(VOIS_COLORS.bodyText), fontFace: FONTS.body, valign: 'middle',
    });

    // ── Lines 2–3: blank bullet points (h: 0.22, w: 6.6) ─────────────
    for (let b = 0; b < 2; b++) {
      const lineY = by + PAD_TOP + (2 + b) * BOX_STEP;
      slide.addText('•', {
        x: tx + 0.05, y: lineY, w: 0.18, h: BOX_H,
        fontSize: 8,
        color: h(VOIS_COLORS.mutedText), fontFace: FONTS.body, valign: 'middle',
      });
      // Editable text box — h: 0.22, w: 6.6
      slide.addText('', {
        x: tx + 0.23, y: lineY, w: 6.6, h: BOX_H,
        fontSize: 8,
        color: h(VOIS_COLORS.bodyText), fontFace: FONTS.body, valign: 'middle',
      });
    }

    // ── Row separator ─────────────────────────────────────────────────
    slide.addShape(LINE, {
      x: LAYOUT.marginL, y: by + ROW_H - PAD_BOT, w: LAYOUT.usableW, h: 0,
      line: { color: h(VOIS_COLORS.divider), pt: 0.5 },
    });
  }
}

// ─── JSZip post-processing: lock all non-editable shapes ─────────────────────

const SHAPE_LOCKS = '<a:spLocks noSelect="1" noMove="1" noResize="1" noRot="1"/>';

/**
 * Returns true if the <p:sp> block contains at least one non-empty <a:t> element.
 * These shapes have pre-filled text and should be locked.
 */
function hasNonEmptyText(spBlock: string): boolean {
  const matches = spBlock.match(/<a:t>([^<]*)<\/a:t>/g);
  if (!matches) return false;
  return matches.some(m => m.replace(/<a:t>|<\/a:t>/g, '').trim().length > 0);
}

/**
 * Insert spLocks into a <p:cNvSpPr> element (handles self-closing and open-tag forms).
 */
function insertSpLocks(spBlock: string): string {
  // Already locked — skip
  if (spBlock.includes('<a:spLocks')) return spBlock;

  // Self-closing: <p:cNvSpPr ... />  →  <p:cNvSpPr ...><a:spLocks .../></p:cNvSpPr>
  const selfClose = /<p:cNvSpPr([^>]*?)\/>/;
  if (selfClose.test(spBlock)) {
    return spBlock.replace(selfClose, `<p:cNvSpPr$1>${SHAPE_LOCKS}</p:cNvSpPr>`);
  }

  // Open tag: <p:cNvSpPr ...>  →  <p:cNvSpPr ...><a:spLocks .../>
  const openTag = /(<p:cNvSpPr[^>]*>)/;
  if (openTag.test(spBlock)) {
    return spBlock.replace(openTag, `$1${SHAPE_LOCKS}`);
  }

  return spBlock;
}

/**
 * Post-process the raw PPTX blob:
 * - shapes with non-empty text  → locked (cannot be selected / moved / resized)
 * - shapes with no <p:txBody>   → locked (decorative rectangles, lines, etc.)
 * - shapes with only empty <a:t> → LEFT EDITABLE (user input boxes)
 */
async function lockNonEditableShapes(blob: Blob): Promise<Blob> {
  const zip = await JSZip.loadAsync(blob);

  const slideFiles = Object.keys(zip.files).filter(
    name => /^ppt\/slides\/slide\d+\.xml$/.test(name)
  );

  for (const slidePath of slideFiles) {
    let xml = await zip.files[slidePath].async('string');

    xml = xml.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, (spBlock) => {
      const hasTxBody = spBlock.includes('<p:txBody>');

      if (hasTxBody && !hasNonEmptyText(spBlock)) {
        // Text body present but all text is empty → editable placeholder, keep unlocked
        return spBlock;
      }

      // Everything else (pre-filled text, decorative shapes) → lock
      return insertSpLocks(spBlock);
    });

    zip.file(slidePath, xml);
  }

  return zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  });
}

// ─── public API ───────────────────────────────────────────────────────────────

export async function generatePPT(pmGroups: PMGroup[]): Promise<Blob> {
  const pptx    = new PptxGenJS();
  pptx.layout   = 'LAYOUT_16x9';
  pptx.author   = 'VOIS PMO';
  pptx.company  = 'VOIS';
  pptx.subject  = 'Portfolio Summary';
  pptx.title    = 'VOIS PMO Portfolio Report';

  let slideNum = 1;
  addSummarySlide(pptx, pmGroups, slideNum++);

  for (const pm of pmGroups) {
    const blocks = flattenProjects(pm);
    const chunks: ProjectBlock[][] = [];
    for (let i = 0; i < blocks.length; i += PROJECTS_PER_SLIDE)
      chunks.push(blocks.slice(i, i + PROJECTS_PER_SLIDE));
    if (!chunks.length) chunks.push([]);

    for (let p = 0; p < chunks.length; p++)
      addPMSlide(pptx, pm, chunks[p], p, chunks.length, slideNum++);
  }

  // Generate raw PPTX blob, then post-process to lock non-editable shapes
  const rawBlob = await pptx.write({ outputType: 'blob' }) as unknown as Blob;
  return lockNonEditableShapes(rawBlob);
}
