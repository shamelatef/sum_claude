import { SlideConfig } from '../types';

// PptxGenJS LAYOUT_WIDE exact dimensions
export const SLIDE_CONFIG: SlideConfig = {
  width: 10,
  height: 5.625,
  margin: {
    top: 0.0,
    bottom: 0.0,
    left: 0.28,
    right: 0.28,
  },
  header: { height: 0.55 },
  footer: { height: 0.35 },
};

// VOIS Red / Black identity palette
export const VOIS_COLORS = {
  primary:        'E60000',   // VOIS / Vodafone Red
  primaryDark:    'B30000',   // Darker red
  black:          '111111',   // Near-black
  darkGray:       '1A1A1A',
  mutedText:      '6B7280',

  headerBg:       '111111',   // Black header
  footerBg:       '1A1A1A',   // Dark footer
  accentStripe:   'E60000',   // Red left stripe

  slideBg:        'F7F7F7',   // Slide background
  white:          'FFFFFF',

  // Cards
  cardBg:         'FFFFFF',
  cardBorder:     'FECACA',
  cardTopBar:     'E60000',   // Red top bar
  pmNameColor:    'B30000',   // Dark red PM name

  // Gate pills (inside summary cards)
  gateTag:        'FEE2E2',
  gateTagText:    'B91C1C',

  // Gate label bar on PM detail slides
  gateSolid:      '111111',
  gateSolidTxt:   'FFFFFF',

  // Grand total bar
  grandTotalBg:   'E60000',
  grandTotalTxt:  'FFFFFF',

  // PM detail slides
  bodyText:       '1F2937',
  sideStrip:      'E60000',
  rowEven:        'FFF5F5',
  rowOdd:         'FFFFFF',
  divider:        'FEE2E2',
  summaryBarBg:   'FFF5F5',
  summaryBarBorder:'FECACA',
  summaryBarText: 'B91C1C',
};

export const FONTS = {
  title: 'Calibri',
  body:  'Calibri',
};

export const FONT_SIZES = {
  slideTitle:   20,
  cardTitle:    11,
  cardProj:      8,
  gateLabel:     7.5,
  projectDetail: 7.5,
  footer:        6,
  slideNumber:   7,
  grandTotal:    10,
  summaryBar:    7.5,
  pmName:        11,
};

export const PROJECTS_PER_SLIDE = 7;

// ─── computed layout constants ───────────────────────────────────────────────
// All sizes in inches.

const W  = SLIDE_CONFIG.width;
const H  = SLIDE_CONFIG.height;
const HDR = SLIDE_CONFIG.header.height;
const FTR = SLIDE_CONFIG.footer.height;
const ML  = SLIDE_CONFIG.margin.left;
const MR  = SLIDE_CONFIG.margin.right;

export const LAYOUT = {
  slideW:  W,
  slideH:  H,
  headerH: HDR,
  footerH: FTR,
  marginL: ML,
  marginR: MR,
  footerY:       H - FTR,
  usableW:       W - ML - MR,
  contentStartY: HDR + 0.08,          // just below header
  contentEndY:   H - FTR - 0.04,      // just above footer

  // Summary slide card area
  GRAND_TOTAL_H:   0.28,
  GRAND_TOTAL_GAP: 0.06,
  get cardsStartY()  { return HDR + 0.08; },
  get cardsEndY()    { return this.contentEndY - this.GRAND_TOTAL_H - this.GRAND_TOTAL_GAP; },
  get cardsTotalH()  { return this.cardsEndY - this.cardsStartY; },
};
