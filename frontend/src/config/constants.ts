export const SLIDE_CONFIG = {
  width: 10,
  height: 5.625,
  margin: { top: 0.0, bottom: 0.0, left: 0.28, right: 0.28 },
  header: { height: 0.55 },
  footer: { height: 0.35 },
};

export const VOIS_COLORS = {
  primary:        'E60000',
  primaryDark:    'B30000',
  black:          '111111',
  darkGray:       '1A1A1A',
  mutedText:      '6B7280',
  headerBg:       '111111',
  footerBg:       '1A1A1A',
  accentStripe:   'E60000',
  slideBg:        'F7F7F7',
  white:          'FFFFFF',
  cardBg:         'FFFFFF',
  cardBorder:     'FECACA',
  cardTopBar:     'E60000',
  pmNameColor:    'B30000',
  gateTag:        'FEE2E2',
  gateTagText:    'B91C1C',
  gateSolid:      '111111',
  gateSolidTxt:   'FFFFFF',
  grandTotalBg:   'E60000',
  grandTotalTxt:  'FFFFFF',
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

export const PROJECTS_PER_SLIDE = 4;

const W   = SLIDE_CONFIG.width;
const H   = SLIDE_CONFIG.height;
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
  contentStartY: HDR + 0.08,
  contentEndY:   H - FTR - 0.04,
  GRAND_TOTAL_H:   0.28,
  GRAND_TOTAL_GAP: 0.06,
  get cardsStartY()  { return HDR + 0.08; },
  get cardsEndY()    { return this.contentEndY - this.GRAND_TOTAL_H - this.GRAND_TOTAL_GAP; },
  get cardsTotalH()  { return this.cardsEndY - this.cardsStartY; },
};
