import PptxGenJS from 'pptxgenjs';
import { PMGroup, GateGroup, Project } from '../types';
import { VOIS_COLORS, FONTS, FONT_SIZES, PROJECTS_PER_SLIDE, LAYOUT } from '../config/constants';
import { truncate } from '../utils/nameUtils';

type SN = PptxGenJS.SHAPE_NAME;
const RECT = 'rect' as SN;
const LINE = 'line' as SN;
const h = (c: string) => (c.startsWith('#') ? c.slice(1) : c);

function distributeCards(count: number): { cols: number; rows: number } {
  if (count <= 1) return { cols: 1, rows: 1 };
  if (count <= 2) return { cols: 2, rows: 1 };
  if (count <= 4) return { cols: 2, rows: Math.ceil(count / 2) };
  if (count <= 6) return { cols: 3, rows: Math.ceil(count / 3) };
  return { cols: 3, rows: Math.ceil(count / 3) };
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

interface ContentItem { kind: 'gate' | 'proj' | 'more'; text: string; }

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

  // ── content area — dynamic row heights so ALL items fit ──────────────
  const contentStartY = divY + 0.04;
  const contentEndY   = cy + cardH - 0.04;
  const contentH      = contentEndY - contentStartY;

  const MAX_GATE_H = 0.22,  MIN_GATE_H = 0.11;
  const MAX_PROJ_H = 0.175, MIN_PROJ_H = 0.09;

  const items: ContentItem[] = [];
  for (const gate of pm.gates) {
    items.push({ kind: 'gate', text: `Gate Approved: ${gate.gate} (${gate.projects.length})` });
    for (const proj of gate.projects) {
      items.push({ kind: 'proj', text: `${proj.projectName} - ${proj.projectId}` });
    }
  }

  // Scale row heights down proportionally so everything fits
  const gateCount = items.filter(i => i.kind === 'gate').length;
  const projCount  = items.filter(i => i.kind === 'proj').length;
  const neededH    = gateCount * MAX_GATE_H + projCount * MAX_PROJ_H;

  let GATE_H: number, PROJ_H: number;
  if (neededH <= contentH) {
    GATE_H = MAX_GATE_H;
    PROJ_H = MAX_PROJ_H;
  } else {
    const scale = contentH / neededH;
    GATE_H = Math.max(MIN_GATE_H, MAX_GATE_H * scale);
    PROJ_H = Math.max(MIN_PROJ_H, MAX_PROJ_H * scale);
  }

  // Font sizes track row height
  const GATE_FONT = Math.max(5.5, GATE_H * 34);
  const PROJ_FONT = Math.max(5.0, PROJ_H * 40);

  // Fit as many items as possible; overflow → "+N more"
  let usedH = 0;
  const visible: ContentItem[] = [];
  for (let i = 0; i < items.length; i++) {
    const rowH = items[i].kind === 'gate' ? GATE_H : PROJ_H;
    if (usedH + rowH > contentH + 0.005) {
      visible.pop();
      const hidden = items.length - visible.length;
      visible.push({ kind: 'more', text: `+${hidden} more` });
      break;
    }
    visible.push(items[i]);
    usedH += rowH;
  }

  let gy = contentStartY;
  for (const item of visible) {
    if (item.kind === 'gate') {
      slide.addShape(RECT, {
        x: cx + 0.07, y: gy + 0.01,
        w: cardW - 0.14, h: GATE_H - 0.02,
        fill: { color: h(VOIS_COLORS.gateTag) },
        line: { color: h(VOIS_COLORS.cardBorder), pt: 0.5 },
      });
      slide.addText(truncate(item.text, 40), {
        x: cx + 0.10, y: gy + 0.01,
        w: cardW - 0.20, h: GATE_H - 0.02,
        fontSize: GATE_FONT, bold: true,
        color: h(VOIS_COLORS.gateTagText), fontFace: FONTS.body, valign: 'middle',
      });
      gy += GATE_H;
    } else if (item.kind === 'proj') {
      slide.addText(`• ${truncate(item.text, 55)}`, {
        x: cx + 0.12, y: gy,
        w: cardW - 0.20, h: PROJ_H,
        fontSize: PROJ_FONT,
        color: h(VOIS_COLORS.bodyText), fontFace: FONTS.body, valign: 'middle',
      });
      gy += PROJ_H;
    } else {
      slide.addText(item.text, {
        x: cx + 0.12, y: gy,
        w: cardW - 0.20, h: PROJ_H,
        fontSize: Math.max(5.0, PROJ_FONT - 0.5), italic: true,
        color: h(VOIS_COLORS.mutedText), fontFace: FONTS.body, valign: 'middle',
      });
      gy += PROJ_H;
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

  const n    = pmGroups.length;
  const grid = distributeCards(n);
  const cols = grid.cols, rows = grid.rows;
  const gapX = 0.09, gapY = 0.07;

  const cardW = (LAYOUT.usableW - gapX * (cols - 1)) / cols;
  const cardH = (LAYOUT.cardsTotalH - gapY * (rows - 1)) / rows;

  for (let i = 0; i < pmGroups.length; i++) {
    const pm  = pmGroups[i];
    const col = i % cols;
    const row = Math.floor(i / cols);
    const cx  = LAYOUT.marginL + col * (cardW + gapX);
    const cy  = LAYOUT.cardsStartY + row * (cardH + gapY);
    drawPMCard(slide, pm, cx, cy, cardW, cardH);
  }

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
