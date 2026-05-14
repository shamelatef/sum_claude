import PptxGenJS from 'pptxgenjs';
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

// Fixed row heights — predictable, consistent
const GATE_ROW_H = 0.22;   // gate pill row
const PROJ_ROW_H = 0.165;  // project bullet row
const MORE_ROW_H = 0.155;  // "+N more" row
const HDR_H_TOTAL = 0.05 + 0.20 + 0.04 + 0.06; // topBar + name + divider gap + content top pad
const CARD_PAD_B  = 0.05;  // inner bottom padding

/** Total natural height needed to show ALL content at fixed row sizes */
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

  slide.addText('VOIS', {
    x: W - 0.62, y: 0, w: 0.59, h: 0.55,
    fontSize: 9.5, bold: true,
    color: h(VOIS_COLORS.primary), fontFace: FONTS.title,
    align: 'center', valign: 'middle',
  });
}

function addFooter(slide: PptxGenJS.Slide, slideNum: number): void {
  const fy = LAYOUT.footerY, W = LAYOUT.slideW, fH = LAYOUT.footerH;

  slide.addShape(RECT, { x: 0, y: fy, w: W, h: fH,
    fill: { color: h(VOIS_COLORS.footerBg) }, line: { color: h(VOIS_COLORS.footerBg) } });

  slide.addShape(RECT, { x: 0, y: fy, w: W, h: 0.018,
    fill: { color: h(VOIS_COLORS.primary) }, line: { color: h(VOIS_COLORS.primary) } });

  slide.addText('DWS EG PM Team', {
    x: 0.16, y: fy + 0.02, w: W - 0.9, h: fH - 0.04,
    fontSize: FONT_SIZES.footer, color: '888888',
    fontFace: FONTS.body, valign: 'middle',
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

  // ── content area — fixed row heights, clean overflow with "+N more" ───
  const contentStartY = divY + 0.04;
  const cardBottom    = cy + cardH - CARD_PAD_B;
  const textW         = cardW - 0.20;

  // Flatten all items so we can count how many fit before committing to render
  interface RI { kind: 'gate' | 'proj'; gate: GateGroup; proj?: Project }
  const allRows: RI[] = [];
  for (const gate of pm.gates) {
    allRows.push({ kind: 'gate', gate });
    for (const proj of gate.projects) allRows.push({ kind: 'proj', gate, proj });
  }

  // Determine how many rows fit, reserving space for "+N more" when needed
  let usedH = contentStartY - cy;  // offset from card top to content start
  let fitCount = 0;
  for (let i = 0; i < allRows.length; i++) {
    const rowH = allRows[i].kind === 'gate' ? GATE_ROW_H : PROJ_ROW_H;
    const remaining = allRows.length - i - 1;
    const needMore  = remaining > 0 ? MORE_ROW_H : 0;
    if (usedH + rowH + needMore + CARD_PAD_B > cardH + 0.005) break;
    usedH += rowH;
    fitCount++;
  }

  const hidden    = allRows.length - fitCount;
  const showRows  = allRows.slice(0, fitCount);

  // Never render a gate label as the very last shown item (orphan guard)
  if (showRows.length > 0 && showRows[showRows.length - 1].kind === 'gate') {
    showRows.pop();
  }
  const finalHidden = allRows.length - showRows.length;

  let gy = contentStartY;
  for (const row of showRows) {
    if (row.kind === 'gate') {
      slide.addShape(RECT, {
        x: cx + 0.07, y: gy + 0.01,
        w: cardW - 0.14, h: GATE_ROW_H - 0.02,
        fill: { color: h(VOIS_COLORS.gateTag) },
        line: { color: h(VOIS_COLORS.cardBorder), pt: 0.5 },
      });
      slide.addText(`Gate Approved: ${row.gate.gate} (${row.gate.projects.length})`, {
        x: cx + 0.10, y: gy + 0.01,
        w: cardW - 0.20, h: GATE_ROW_H - 0.02,
        fontSize: F_GATE, bold: true,
        color: h(VOIS_COLORS.gateTagText), fontFace: FONTS.body, valign: 'middle',
      });
      gy += GATE_ROW_H;
    } else {
      const txt = `• ${row.proj!.projectName} - ${row.proj!.projectId}`;
      slide.addText(txt, {
        x: cx + 0.12, y: gy,
        w: textW, h: PROJ_ROW_H,
        fontSize: F_PROJ,
        color: h(VOIS_COLORS.bodyText), fontFace: FONTS.body,
        valign: 'middle',
      });
      gy += PROJ_ROW_H;
    }
  }

  // "+N more" indicator if content was clipped
  if (finalHidden > 0) {
    slide.addText(`+${finalHidden} more`, {
      x: cx + 0.12, y: gy,
      w: textW, h: MORE_ROW_H,
      fontSize: 6, italic: true,
      color: h(VOIS_COLORS.mutedText), fontFace: FONTS.body, valign: 'middle',
    });
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
  slide.addShape(RECT, {
    x: LAYOUT.marginL, y: sy, w: LAYOUT.usableW, h: 0.21,
    fill: { color: h(VOIS_COLORS.summaryBarBg) },
    line: { color: h(VOIS_COLORS.summaryBarBorder), pt: 0.5 },
  });
  slide.addText(
    `PM: ${pm.pmName}   ·   Total: ${pm.totalProjects} projects   ·   Slide: ${blocks.length} shown`,
    {
      x: LAYOUT.marginL + 0.10, y: sy, w: LAYOUT.usableW - 0.2, h: 0.21,
      fontSize: FONT_SIZES.summaryBar,
      color: h(VOIS_COLORS.summaryBarText), fontFace: FONTS.body, valign: 'middle',
    }
  );

  const listStartY = sy + 0.25;
  const listEndY   = LAYOUT.contentEndY;
  const listH      = listEndY - listStartY;
  const rowH       = blocks.length > 0 ? Math.min(listH / blocks.length, 0.68) : 0.68;

  const GATE_W = 1.55;
  let currentGate = '';

  for (let i = 0; i < blocks.length; i++) {
    const { gate, project } = blocks[i];
    const by = listStartY + i * rowH;
    if (by + rowH > listEndY + 0.04) break;

    const isNewGate = gate.gate !== currentGate;
    currentGate = gate.gate;

    slide.addShape(RECT, {
      x: LAYOUT.marginL, y: by, w: LAYOUT.usableW, h: rowH,
      fill: { color: i % 2 === 0 ? h(VOIS_COLORS.rowEven) : h(VOIS_COLORS.rowOdd) },
      line: { color: h(VOIS_COLORS.divider), pt: 0.25 },
    });

    slide.addShape(RECT, {
      x: LAYOUT.marginL, y: by, w: 0.022, h: rowH,
      fill: { color: h(VOIS_COLORS.primary) }, line: { color: h(VOIS_COLORS.primary) },
    });

    if (isNewGate) {
      slide.addShape(RECT, {
        x: LAYOUT.marginL + 0.022, y: by,
        w: GATE_W, h: rowH * 0.40,
        fill: { color: h(VOIS_COLORS.gateSolid) }, line: { color: h(VOIS_COLORS.gateSolid) },
      });
      slide.addText(`Gate Approved: ${gate.gate} (${gate.projects.length})`, {
        x: LAYOUT.marginL + 0.022, y: by,
        w: GATE_W, h: rowH * 0.40,
        fontSize: FONT_SIZES.gateLabel, bold: true,
        color: h(VOIS_COLORS.gateSolidTxt), fontFace: FONTS.body,
        align: 'center', valign: 'middle',
      });
    }

    const tx = LAYOUT.marginL + GATE_W + 0.10;
    const tw = LAYOUT.usableW - GATE_W - 0.12;
    const lh = rowH / 3.2;

    slide.addText(truncate(project.projectName, 90), {
      x: tx, y: by + lh * 0.06, w: tw, h: lh,
      fontSize: FONT_SIZES.pmName, bold: true,
      color: h(VOIS_COLORS.bodyText), fontFace: FONTS.title, valign: 'middle',
    });
    slide.addText(`ID: ${project.projectId}`, {
      x: tx, y: by + lh * 1.08, w: tw * 0.46, h: lh,
      fontSize: FONT_SIZES.projectDetail,
      color: h(VOIS_COLORS.mutedText), fontFace: FONTS.body, valign: 'middle',
    });
    slide.addShape(RECT, {
      x: tx + tw * 0.48, y: by + lh * 1.12,
      w: 0.58, h: lh * 0.72,
      fill: { color: h(VOIS_COLORS.gateTag) },
      line: { color: h(VOIS_COLORS.cardBorder), pt: 0.3 },
    });
    slide.addText(truncate(project.gateApproved, 20), {
      x: tx + tw * 0.48, y: by + lh * 1.12,
      w: 0.58, h: lh * 0.72,
      fontSize: FONT_SIZES.gateLabel, bold: true,
      color: h(VOIS_COLORS.gateTagText), fontFace: FONTS.body,
      align: 'center', valign: 'middle',
    });

    slide.addShape(LINE, {
      x: LAYOUT.marginL, y: by + rowH - 0.005, w: LAYOUT.usableW, h: 0,
      line: { color: h(VOIS_COLORS.divider), pt: 0.35 },
    });
  }
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

  return await pptx.write({ outputType: 'blob' }) as unknown as Blob;
}
