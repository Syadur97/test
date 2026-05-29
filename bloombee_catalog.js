const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  LevelFormat, PageBreak, VerticalAlign
} = require('docx');
const fs = require('fs');

// ── Brand colours ──────────────────────────────────────────────────────────
const C = {
  brand:     "8B1A4A",   // deep rose
  gold:      "B8860B",   // dark gold
  tier1:     "D6EAF8",   // mini  – soft blue
  tier2:     "E8F8E8",   // regular – soft green
  tier3:     "FDF2E9",   // premium – warm ivory
  hdrMini:   "2471A3",
  hdrReg:    "1E8449",
  hdrPrem:   "784212",
  hdrText:   "FFFFFF",
  tableHdr:  "4A1942",
  rowAlt:    "F9F0F4",
  border:    "D5B0C8",
};

// ── Helpers ─────────────────────────────────────────────────────────────────
const sp = (bef, aft) => ({ spacing: { before: bef, after: aft } });
const cellBorder = (color = C.border) => {
  const b = { style: BorderStyle.SINGLE, size: 1, color };
  return { top: b, bottom: b, left: b, right: b };
};

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, font: "Arial", size: 36, bold: true, color: C.brand })],
    ...sp(360, 200),
    alignment: AlignmentType.CENTER,
  });
}
function heading2(text, color = C.brand) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, font: "Arial", size: 28, bold: true, color })],
    ...sp(300, 120),
  });
}
function heading3(text, color = "333333") {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    children: [new TextRun({ text, font: "Arial", size: 22, bold: true, color })],
    ...sp(200, 80),
  });
}
function body(text, bold = false, italic = false, color = "222222") {
  return new Paragraph({
    children: [new TextRun({ text, font: "Arial", size: 22, bold, italic, color })],
    ...sp(60, 60),
  });
}
function labelVal(label, value) {
  return new Paragraph({
    children: [
      new TextRun({ text: `${label}: `, font: "Arial", size: 22, bold: true, color: C.brand }),
      new TextRun({ text: value, font: "Arial", size: 22, color: "333333" }),
    ],
    ...sp(60, 60),
  });
}
function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    children: [new TextRun({ text, font: "Arial", size: 22, color: "333333" })],
    ...sp(40, 40),
  });
}
function divider(color = C.border) {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color, space: 1 } },
    ...sp(160, 160),
    children: [],
  });
}
function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}
function centeredBold(text, size = 26, color = C.brand) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text, font: "Arial", size, bold: true, color })],
    ...sp(100, 100),
  });
}
function tagLine(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text, font: "Arial", size: 22, italic: true, color: "666666" })],
    ...sp(60, 180),
  });
}

// ── Tier Table ──────────────────────────────────────────────────────────────
function tierTable(rows) {
  // rows: [{label, value}]
  const W = 9360;
  const c1 = 2800, c2 = W - c1;
  return new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [c1, c2],
    rows: rows.map((r, i) => new TableRow({
      children: [
        new TableCell({
          borders: cellBorder(),
          width: { size: c1, type: WidthType.DXA },
          shading: { fill: i === 0 ? C.tableHdr : (i % 2 === 0 ? C.rowAlt : "FFFFFF"), type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({
            children: [new TextRun({ text: r.label, font: "Arial", size: 20,
              bold: i === 0, color: i === 0 ? "FFFFFF" : C.brand })],
          })],
        }),
        new TableCell({
          borders: cellBorder(),
          width: { size: c2, type: WidthType.DXA },
          shading: { fill: i === 0 ? C.tableHdr : (i % 2 === 0 ? C.rowAlt : "FFFFFF"), type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({
            children: [new TextRun({ text: r.value, font: "Arial", size: 20,
              bold: i === 0, color: i === 0 ? "FFFFFF" : "333333" })],
          })],
        }),
      ],
    })),
  });
}

// ── Per-tier section ─────────────────────────────────────────────────────────
function tierSection({ tierName, bangla, tagline, price, priceNote, items, extras, bgColor, hdrColor }) {
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      shading: { fill: hdrColor, type: ShadingType.CLEAR },
      ...sp(120, 0),
      children: [
        new TextRun({ text: `  ${tierName}  `, font: "Arial", size: 28, bold: true, color: "FFFFFF" }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      shading: { fill: bgColor, type: ShadingType.CLEAR },
      ...sp(0, 0),
      children: [
        new TextRun({ text: bangla, font: "Arial", size: 22, italic: true, color: hdrColor }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      shading: { fill: bgColor, type: ShadingType.CLEAR },
      ...sp(0, 80),
      children: [new TextRun({ text: tagline, font: "Arial", size: 20, italic: true, color: "555555" })],
    }),
    tierTable([
      { label: "Field", value: "Detail" },
      { label: "Price", value: price },
      { label: "Price Note", value: priceNote },
      { label: "Target", value: "Couples, families gifting anniversary" },
    ]),
    new Paragraph({ ...sp(120, 60), children: [] }),
    heading3("What's Inside the Box", hdrColor),
    ...items.map(bullet),
    ...(extras && extras.length ? [
      heading3("Special Extras", hdrColor),
      ...extras.map(bullet),
    ] : []),
    new Paragraph({ ...sp(80, 80), children: [] }),
  ];
}

// ── Meta Info Section ─────────────────────────────────────────────────────────
function metaSection(boxName, meta) {
  const W = 9360;
  const c1 = 2000, c2 = W - c1;
  const rows = [
    { label: "Field", value: "Content" },
    { label: "SEO Title", value: meta.title },
    { label: "Meta Description", value: meta.description },
    { label: "Keywords", value: meta.keywords },
  ];
  return [
    heading3("Meta Information (SEO)", C.gold),
    new Table({
      width: { size: W, type: WidthType.DXA },
      columnWidths: [c1, c2],
      rows: rows.map((r, i) => new TableRow({
        children: [
          new TableCell({
            borders: cellBorder(C.gold),
            width: { size: c1, type: WidthType.DXA },
            shading: { fill: i === 0 ? C.gold : (i % 2 === 0 ? "FEF9EE" : "FFFFFF"), type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [new Paragraph({
              children: [new TextRun({ text: r.label, font: "Arial", size: 20,
                bold: true, color: i === 0 ? "FFFFFF" : C.gold })],
            })],
          }),
          new TableCell({
            borders: cellBorder(C.gold),
            width: { size: c2, type: WidthType.DXA },
            shading: { fill: i === 0 ? C.gold : (i % 2 === 0 ? "FEF9EE" : "FFFFFF"), type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [new Paragraph({
              children: [new TextRun({ text: r.value, font: "Arial", size: 19,
                bold: i === 0, color: i === 0 ? "FFFFFF" : "333333" })],
            })],
          }),
        ],
      })),
    }),
    new Paragraph({ ...sp(100, 100), children: [] }),
  ];
}

// ═══════════════════════════════════════════════════════════════════════════
// DATA
// ═══════════════════════════════════════════════════════════════════════════

const boxes = [
  {
    id: 1,
    name: "Ek Bonde Bhalobashar Pore",
    tagline: "One Bond, One Love — Forever Cherished",
    bangla_subtitle: "একটি বন্ধনে, অনেক ভালোবাসা",
    description: "A heartfelt anniversary gift collection celebrating the unbreakable bond of two souls in love — rooted in the beauty of Bangladeshi traditions and tender modern moments.",
    tiers: [
      {
        tierName: "Mini — Kotha Chhilo",
        bangla: "কথা ছিলো | The Promise",
        tagline: "A small box that carries big feelings",
        price: "BDT 750 – 999",
        priceNote: "Budget-friendly, widely available items",
        hdrColor: C.hdrMini,
        bgColor: C.tier1,
        items: [
          "Handwritten love-note card (locally printed, Bashundhara paper)",
          "Rosewater face mist – 100ml (Meena Bazar / Shajgoj available brands)",
          "Rose-scented tea-light candle set (2 pcs) – local craft",
          "Couple's mini notepad with pen (personalised cover print)",
          "Assorted chocolate bar (Bôn Bon / Cadbury available locally)",
          "Organza ribbon bow with kraft gift box packaging",
        ],
        extras: null,
      },
      {
        tierName: "Regular — Shomoy Theke Shomoy",
        bangla: "সময় থেকে সময় | Through Every Season",
        tagline: "Curated warmth for the couple who has grown together",
        price: "BDT 1,500 – 1,999",
        priceNote: "Mid-range quality items, reliable local brands",
        hdrColor: C.hdrReg,
        bgColor: C.tier2,
        items: [
          "Premium greeting card with envelope (Shopno Niloy / Grameenphone Card designs)",
          "Couple's ceramic mug set (2 pcs) — printed with names/date",
          "Lavender-scented soy candle in glass jar — medium size",
          "Face serum or moisturiser duo (Garnier/Simple/Pond's — widely stocked)",
          "Silk-touch artificial rose arrangement in mini vase",
          "Dark chocolate assortment box (Ferrero / local premium brand)",
          "Floral printed photo frame for couple's memory",
          "Custom love coupon booklet (printed locally)",
          "Satin ribbon, dried flower filler, premium box with magnetic lid",
        ],
        extras: null,
      },
      {
        tierName: "Premium — Shobar Shera Upohar",
        bangla: "সবার সেরা উপহার | The Gift Above All",
        tagline: "Luxury, personalisation, and lasting memories — no compromise",
        price: "BDT 2,000 and above",
        priceNote: "Fully customisable; price increases with selections",
        hdrColor: C.hdrPrem,
        bgColor: C.tier3,
        items: [
          "Custom embossed name/date hardcover journal — couple's edition",
          "Premium perfume duo (men + women) — Engage / Fogg Gold / Imported (Daraz/local)",
          "Spa & bath gift set: body lotion, scrub, bath salts, loofa (Korean / local brand)",
          "Personalised photo album or photo book (printed via local photo studio)",
          "Luxury scented candle — Amber or Oud fragrance (large jar)",
          "Premium Belgian chocolate box (20+ pcs, Daraz imported)",
          "Artificial preserved rose in glass dome or acrylic box",
          "Couple name necklace or bracelet (silver-plated, local jewellery vendor)",
          "Printed couple portrait / caricature (digital art, locally printed A4)",
          "Custom anniversary message bottle with scroll",
          "Luxury rigid gift box with velvet interior, satin ribbon, and dried petals",
        ],
        extras: [
          "Optional: Add-on personalised video message QR card",
          "Optional: Balloon bouquet delivery coordination",
          "Optional: Premium wrapping with wax seal and custom sticker",
          "Optional: Cake voucher from local bakery (Cakesburg, Maron, etc.)",
        ],
      },
    ],
    meta: {
      mini: {
        title: "Ek Bonde Bhalobashar Pore – Mini Anniversary Gift Box | Under BDT 1000 | Bangladesh",
        description: "Celebrate your anniversary with the Kotha Chhilo Mini Gift Box — a beautifully curated, affordable set of heartfelt gifts including candles, chocolates, and a personal love card. Perfect for budget-conscious gifting in Bangladesh. Free delivery available.",
        keywords: "anniversary gift box Bangladesh, mini gift box BDT 1000, couple gift Bangladesh, romantic gift Dhaka, small anniversary gift, affordable gift box BD",
      },
      regular: {
        title: "Ek Bonde Bhalobashar Pore – Regular Anniversary Gift Box | BDT 1500–1999 | Bangladesh",
        description: "The Shomoy Theke Shomoy Regular Gift Box brings together ceramic mugs, scented candles, skincare, and chocolates in a premium magnetic-lid box. Ideal mid-range anniversary gift for couples in Bangladesh.",
        keywords: "anniversary gift box Bangladesh, regular couple gift, BDT 2000 anniversary gift, romantic gift box Dhaka, ceramic mug couple gift, skincare gift Bangladesh",
      },
      premium: {
        title: "Ek Bonde Bhalobashar Pore – Premium Anniversary Gift Box | Custom BDT 2000+ | Bangladesh",
        description: "Luxury, personalised anniversary gifting with the Shobar Shera Upohar Premium Box — perfume duo, spa sets, preserved roses, jewellery, and full customisation. The ultimate anniversary gift in Bangladesh.",
        keywords: "luxury anniversary gift Bangladesh, premium couple gift box, personalised gift Bangladesh, anniversary perfume gift, custom name gift BD, 2000 taka anniversary gift",
      },
    },
  },
  {
    id: 2,
    name: "Priya O Priyotar Golpo",
    tagline: "The Story of the Beloved — Told in Gifts",
    bangla_subtitle: "প্রিয় ও প্রিয়তার গল্প",
    description: "Inspired by Bangla literature's timeless romance, this gift collection narrates the story of two people in love through carefully chosen objects — each item a chapter in their shared journey.",
    tiers: [
      {
        tierName: "Mini — Prothom Adhyay",
        bangla: "প্রথম অধ্যায় | The First Chapter",
        tagline: "Simple beginnings, genuine feelings",
        price: "BDT 700 – 999",
        priceNote: "Affordable starter anniversary gift",
        hdrColor: C.hdrMini,
        bgColor: C.tier1,
        items: [
          "Illustrated anniversary greeting card with envelope",
          "Couple's bookmark set (2 pcs) — personalised with names",
          "Jasmine or rose-scented roll-on perfume (local brand, Aarong)",
          "Mini rose quartz crystal stone (love symbol, available in Dhaka craft shops)",
          "Handmade chocolate (2 pcs) — local confectioner",
          "Jute-wrapped small box with dried flowers",
        ],
        extras: null,
      },
      {
        tierName: "Regular — Madhyo Adhyay",
        bangla: "মধ্য অধ্যায় | The Middle Story",
        tagline: "A deeper gift, for a love that has grown deeper",
        price: "BDT 1,400 – 1,999",
        priceNote: "Balanced quality and price",
        hdrColor: C.hdrReg,
        bgColor: C.tier2,
        items: [
          "Premium hardcover couple's journal — matching set",
          "Rose-gold scented diffuser with rattan sticks (local or Daraz import)",
          "Moisturising hand cream duo — his and hers (Dove/Vaseline/Neutrogena)",
          "Silk flower bouquet with decorative vase",
          "Custom photo magnet set (4 pcs) — printed locally",
          "Premium dry fruit and nut gift box (almonds, cashews, dates — local grocery)",
          "Anniversary countdown calendar card",
          "Premium kraft paper box with tissue and ribbon",
        ],
        extras: null,
      },
      {
        tierName: "Premium — Shesh Adhyay Nei",
        bangla: "শেষ অধ্যায় নেই | A Story Without End",
        tagline: "Because their love story never ends — neither should the gifting",
        price: "BDT 2,000 and above",
        priceNote: "Fully customisable luxury bundle",
        hdrColor: C.hdrPrem,
        bgColor: C.tier3,
        items: [
          "Custom engraved wooden keepsake box with their names and anniversary date",
          "Luxury perfume set — branded or semi-imported (Rasasi/Swiss Arabian available BD)",
          "Premium skincare gift set — moisturiser, toner, serum trio (Korean or POND'S Gold)",
          "Artificial preserved rose bouquet in ribbon wrap",
          "Couple's matching silver bracelet with name tag",
          "Custom 3D printed miniature couple figurine (available via Bangladeshi 3D shops)",
          "LED fairy light photo display string for memories",
          "Premium date + nut + chocolate hamper in a wooden tray",
          "Personalised video message QR code printed card",
          "Luxury gift box — rigid, velvet-lined, with ribbon pull",
        ],
        extras: [
          "Optional: Couple caricature portrait (digital, local artist)",
          "Optional: Saree or panjabi add-on (Aarong / Kay Kraft coordination)",
          "Optional: Dinner reservation assistance voucher",
          "Optional: Custom song or poem written by local writer",
        ],
      },
    ],
    meta: {
      mini: {
        title: "Priya O Priyotar Golpo – Mini Anniversary Gift Box | BDT under 1000 | Bangladesh",
        description: "Begin the story with the Prothom Adhyay Mini Gift Box — a charming, affordable anniversary collection with jasmine perfume, rose quartz, chocolates, and handmade touches. Perfect first anniversary or budget gift in Bangladesh.",
        keywords: "mini anniversary gift Bangladesh, romantic gift box 1000 taka, couple gift Dhaka, affordable anniversary BD, jasmine perfume gift Bangladesh, gift box under 1000",
      },
      regular: {
        title: "Priya O Priyotar Golpo – Regular Anniversary Gift Box | BDT 1500–1999 | Bangladesh",
        description: "The Madhyo Adhyay Regular Gift Box combines a couple's journal, diffuser, dry fruit hamper, and silk flowers in a beautifully packaged mid-range anniversary gift — perfect for celebrating love in Bangladesh.",
        keywords: "regular anniversary gift Bangladesh, couple journal gift, scented diffuser gift BD, mid-range anniversary box, dry fruit hamper gift Bangladesh, romantic gift 2000 taka",
      },
      premium: {
        title: "Priya O Priyotar Golpo – Premium Anniversary Gift Box | Customised BDT 2000+ | Bangladesh",
        description: "The Shesh Adhyay Nei Premium Gift Box tells a love story that never ends — engraved keepsake boxes, luxury perfumes, K-beauty skincare, couple bracelets, and full personalisation. The ultimate anniversary gift in Bangladesh.",
        keywords: "premium anniversary gift Bangladesh, luxury couple gift BD, custom anniversary box, engraved gift Bangladesh, K-beauty gift BD, 3D couple figurine Bangladesh",
      },
    },
  },
  {
    id: 3,
    name: "Duijoner Akash",
    tagline: "Two Hearts, One Sky — Infinite Together",
    bangla_subtitle: "দুইজনের আকাশ",
    description: "Inspired by the boundless sky that two people share, this collection brings together gifts as wide and warm as love itself — aspirational, culturally grounded, and beautifully presented for Bangladeshi couples.",
    tiers: [
      {
        tierName: "Mini — Chhotto Akash",
        bangla: "ছোটো আকাশ | A Pocket of Sky",
        tagline: "Small, sincere, and full of love",
        price: "BDT 650 – 999",
        priceNote: "Entry-level gifting, everyday items elevated",
        hdrColor: C.hdrMini,
        bgColor: C.tier1,
        items: [
          "Printed anniversary card with Bangla verse inside",
          "Mini candle in tin — rose or vanilla scent (local craft brand)",
          "Lip balm duo (his & hers) — Vaseline / local brand",
          "Couple's keychain set — heart-shaped, locally sourced",
          "Gourmet date and nut pouch (premium packaging)",
          "Small pastel gift box with crinkle paper filler",
        ],
        extras: null,
      },
      {
        tierName: "Regular — Khola Akash",
        bangla: "খোলা আকাশ | Open Skies",
        tagline: "Room to breathe, room to love — a gift that expands",
        price: "BDT 1,200 – 1,999",
        priceNote: "Mid-range, quality-focused selection",
        hdrColor: C.hdrReg,
        bgColor: C.tier2,
        items: [
          "Customised couple's cushion cover (printed, local vendor)",
          "Body butter and lotion gift set — large size (The Body Shop / local)",
          "Scented pillar candle with candleholder (local décor shops)",
          "Mini couple's memory jar with 30 written notes",
          "Assorted premium biscuit and snack gift box",
          "Artificial eucalyptus wreath or floral ring for décor",
          "Custom anniversary photo calendar (desk or wall)",
          "Premium ribbon-tied box with crinkle paper",
        ],
        extras: null,
      },
      {
        tierName: "Premium — Ononto Akash",
        bangla: "অনন্ত আকাশ | Infinite Sky",
        tagline: "A gift as limitless as love — fully personalised, unforgettably luxurious",
        price: "BDT 2,000 and above",
        priceNote: "Fully customisable; curated on request",
        hdrColor: C.hdrPrem,
        bgColor: C.tier3,
        items: [
          "Luxury couple's bathrobe set — embroidered with names (local tailor / Aarong)",
          "Premium oud or floral perfume — imported mid-range (Daraz sourced)",
          "Complete skincare ritual set: cleanser, serum, eye cream, moisturiser",
          "Swarovski-inspired crystal love ornament / couple showpiece",
          "Custom neon sign — couple's name or date (local neon vendors, Dhaka)",
          "Gourmet hamper: imported chocolates, cheese crackers, premium tea",
          "Custom star map of their anniversary date and location (printed locally)",
          "Personalised couple portrait in watercolour style (local artist)",
          "Engraved wooden photo frame with USB music box",
          "Premium gift trunk box — wooden or rigid, branded Bloom Bee",
        ],
        extras: [
          "Optional: Live flower arrangement from local florist",
          "Optional: Couple spa voucher (local wellness centre)",
          "Optional: Anniversary jewellery — ring / necklace (gold-plated)",
          "Optional: Drone photo session coordination add-on",
        ],
      },
    ],
    meta: {
      mini: {
        title: "Duijoner Akash – Mini Anniversary Gift Box | Under BDT 1000 | Bangladesh",
        description: "The Chhotto Akash Mini Box is a heartfelt, pocket-friendly anniversary gift — keychain couple set, scented candles, lip balm, and dates in a charming pastel box. Ideal for quick, meaningful gifting in Bangladesh.",
        keywords: "mini couple gift Bangladesh, anniversary keychain gift, small gift box BD, 1000 taka gift Bangladesh, romantic mini gift, couple gift under 1000 taka",
      },
      regular: {
        title: "Duijoner Akash – Regular Anniversary Gift Box | BDT 1200–1999 | Bangladesh",
        description: "The Khola Akash Regular Box brings joy through a customised couple's cushion, body butter, memory jar, and snack hamper — beautifully wrapped and ready to gift for any anniversary in Bangladesh.",
        keywords: "regular anniversary gift BD, couple cushion gift Bangladesh, memory jar anniversary, body butter gift set BD, snack hamper Bangladesh, romantic gift 2000 BDT",
      },
      premium: {
        title: "Duijoner Akash – Premium Anniversary Gift Box | Custom BDT 2000+ | Bangladesh",
        description: "The Ononto Akash Premium Box is a luxurious anniversary experience — couple bathrobes, oud perfume, neon signs, custom star maps, and a gourmet hamper. Fully personalised, unforgettably delivered across Bangladesh.",
        keywords: "premium anniversary Bangladesh, luxury anniversary hamper, couple robe gift BD, custom star map Bangladesh, neon sign gift Bangladesh, anniversary gift above 2000 BDT",
      },
    },
  },
  {
    id: 4,
    name: "Tumi Amar Bhalobasha",
    tagline: "You Are My Love — A Gift That Says Everything",
    bangla_subtitle: "তুমি আমার ভালোবাসা",
    description: "A direct, warm declaration of love — this collection translates the deepest Bangla sentiment into tangible gifts. Perfect for anniversary moments that deserve to be felt, not just celebrated.",
    tiers: [
      {
        tierName: "Mini — Chhuye Dekho",
        bangla: "ছুঁয়ে দেখো | Touch and Feel",
        tagline: "Tangible love in its simplest form",
        price: "BDT 700 – 999",
        priceNote: "Everyday items with emotional weight",
        hdrColor: C.hdrMini,
        bgColor: C.tier1,
        items: [
          "Handwritten anniversary card — customised message",
          "Couple's matching bookmark (leather-look, local craft)",
          "Rose water facial spray — 50ml (Meena Bazar / Shajgoj)",
          "Jasmine incense stick set in decorative holder",
          "Mini love jar with folded paper hearts (50 reasons I love you)",
          "Festive tissue and ribbon gift wrap",
        ],
        extras: null,
      },
      {
        tierName: "Regular — Kache Thako",
        bangla: "কাছে থাকো | Stay Close",
        tagline: "For couples who choose closeness every day",
        price: "BDT 1,300 – 1,999",
        priceNote: "Balanced, meaningful mid-range items",
        hdrColor: C.hdrReg,
        bgColor: C.tier2,
        items: [
          "Personalised couple's phone case set (custom print, local vendor)",
          "Aloe vera and rose gel moisturiser gift set",
          "Scented soy wax candle in floral jar",
          "Couple's love lock with two keys (decorative, locally available)",
          "Premium mixed dry fruit and chocolate gift platter",
          "Floral-printed mini photo album (24 photos)",
          "Anniversary wristband pair — engraved message",
          "Ribbon-tied premium gift box",
        ],
        extras: null,
      },
      {
        tierName: "Premium — Jibon Bhore",
        bangla: "জীবন ভরে | For a Lifetime",
        tagline: "A lifetime commitment deserves a lifetime-quality gift",
        price: "BDT 2,000 and above",
        priceNote: "Fully customisable premium bundle",
        hdrColor: C.hdrPrem,
        bgColor: C.tier3,
        items: [
          "Custom couple's silk-print scarf or panjabi-saree combo (Aarong / Kay Kraft referral)",
          "Luxury perfume set — imported or premium local (Rasasi, Khadlaj, Fogg Gold)",
          "Premium men's grooming kit: razor, cream, aftershave, cologne",
          "Premium women's skincare set: serum, mask, cream, toner",
          "Handmade preserved flower arrangement in acrylic box",
          "Personalised leather wallet or purse with name embossing",
          "Couple's digital portrait — illustrated by local artist",
          "Premium chocolate tower or tiered gift basket",
          "Custom music box with engraved message on lid",
          "Rigid luxury gift box with satin interior, branded ribbon",
        ],
        extras: [
          "Optional: Monogrammed linen pillow pair",
          "Optional: Custom anniversary photo book (hardcover, 40 pages)",
          "Optional: Couple's cooking class voucher coordination",
          "Optional: Premium gift-wrapping with wax seal and personalised sticker",
        ],
      },
    ],
    meta: {
      mini: {
        title: "Tumi Amar Bhalobasha – Mini Anniversary Gift Box | BDT under 1000 | Bangladesh",
        description: "Say it simply and sincerely — the Chhuye Dekho Mini Gift Box features a love jar, rose water, jasmine incense, and a heartfelt card. An affordable yet deeply personal anniversary gift in Bangladesh.",
        keywords: "mini anniversary gift Bangladesh, love jar gift BD, jasmine incense gift, small romantic gift Bangladesh, anniversary card gift set, BDT 1000 gift box",
      },
      regular: {
        title: "Tumi Amar Bhalobasha – Regular Anniversary Gift Box | BDT 1300–1999 | Bangladesh",
        description: "The Kache Thako Regular Box includes couple love locks, custom phone cases, moisturising gift sets, and a floral mini album — a warm, meaningful mid-range anniversary gift across Bangladesh.",
        keywords: "couple love lock gift Bangladesh, anniversary phone case gift, mid-range gift box BD, romantic anniversary Bangladesh, aloe moisturiser gift, 2000 BDT anniversary gift",
      },
      premium: {
        title: "Tumi Amar Bhalobasha – Premium Anniversary Gift Box | Custom BDT 2000+ | Bangladesh",
        description: "For a love that deserves the finest — the Jibon Bhore Premium Box features grooming kits, skincare sets, preserved roses, leather gifts, custom portraits, and full personalisation. The ultimate anniversary in Bangladesh.",
        keywords: "premium anniversary Bangladesh, luxury grooming gift BD, leather wallet gift Bangladesh, custom couple portrait BD, premium skincare gift set, anniversary above 2000 taka",
      },
    },
  },
  {
    id: 5,
    name: "Ekti Shundor Preetir Utsob",
    tagline: "A Beautiful Festival of Love — Celebrate Every Year",
    bangla_subtitle: "একটি সুন্দর প্রীতির উৎসব",
    description: "Anniversary is a festival — and every festival deserves celebration. Rooted in Bangladesh's vibrant gift culture and festive spirit, this collection turns ordinary moments into extraordinary memories.",
    tiers: [
      {
        tierName: "Mini — Utsober Chhotto Bhethor",
        bangla: "উৎসবের ছোটো ভেতর | A Small Festival Inside",
        tagline: "Celebrate in a small way, feel it in a big way",
        price: "BDT 650 – 999",
        priceNote: "Festival-inspired affordable gifting",
        hdrColor: C.hdrMini,
        bgColor: C.tier1,
        items: [
          "Festive printed anniversary card with gold foil accent",
          "Miniature terracotta clay diya set (2 pcs) with tea-light",
          "Rose water and glycerine skin toner — 100ml (Meena Bazar)",
          "Bengali sweet sampler — mishti doi / sandesh (local sweets shop coordination)",
          "Couple's mini keychain with date engraving",
          "Handmade jute gift bag with ribbon",
        ],
        extras: null,
      },
      {
        tierName: "Regular — Utsob Jome Uthuk",
        bangla: "উৎসব জমে উঠুক | Let the Celebration Rise",
        tagline: "A full festive experience, gift after gift",
        price: "BDT 1,300 – 1,999",
        priceNote: "Culturally rich, quality-forward selection",
        hdrColor: C.hdrReg,
        bgColor: C.tier2,
        items: [
          "Embroidered couple's handkerchief set (Bangladeshi fabric)",
          "Premium Oud-based room spray — 100ml (Daraz / local)",
          "Luxury face mask sheet set (Korean brand or local beauty)",
          "Sweet and savoury gourmet gift basket — local brands",
          "Decorative tealight candle holder with 6 candles",
          "Custom couple's fridge magnet with photo print",
          "Anniversary countdown candle (numbered, locally available)",
          "Premium basket-style gift packaging with bow",
        ],
        extras: null,
      },
      {
        tierName: "Premium — Boro Utsob Boro Bhalobasha",
        bangla: "বড় উৎসব, বড় ভালোবাসা | Big Festival, Big Love",
        tagline: "An anniversary worth remembering — forever",
        price: "BDT 2,000 and above",
        priceNote: "Fully customisable luxury celebration bundle",
        hdrColor: C.hdrPrem,
        bgColor: C.tier3,
        items: [
          "Handwoven Jamdani or muslin fabric gift (Aarong sourced) — cultural heritage touch",
          "Luxury gift hamper: dates, saffron, dry fruits, premium tea, imported chocolates",
          "Premium couple's spa kit: bath salts, essential oils, face mask, body scrub",
          "Custom LED light-up photo frame — acrylic with USB",
          "Couple's embroidered throw blanket with names (local vendor)",
          "Luxury scented candle — large oud or amber (imported via Daraz)",
          "Personalised 'Year in Moments' photobook (60 pages, hardcover)",
          "Crystal anniversary keepsake: champagne glass style ornament",
          "Name-engraved jewellery box (velvet interior)",
          "Grand gift chest — wooden or rigid luxury box, Bloom Bee branded",
        ],
        extras: [
          "Optional: Bengali sweet hamper from premium sweets brand (Fakruddin, Mishti Hub)",
          "Optional: Couple saree + panjabi coordination (local boutique)",
          "Optional: Video montage / slideshow creation service",
          "Optional: Hotel or resort anniversary night coordination",
        ],
      },
    ],
    meta: {
      mini: {
        title: "Ekti Shundor Preetir Utsob – Mini Anniversary Gift Box | BDT under 1000 | Bangladesh",
        description: "Celebrate with tradition — the Utsober Chhotto Bhethor Mini Box features terracotta diyas, Bengali sweets, rose toner, and festive packaging. A culturally warm, affordable anniversary gift across Bangladesh.",
        keywords: "mini anniversary gift Bangladesh, cultural gift box BD, terracotta diya gift, Bengali sweet anniversary gift, affordable festive gift, BDT 1000 gift box",
      },
      regular: {
        title: "Ekti Shundor Preetir Utsob – Regular Anniversary Gift Box | BDT 1300–1999 | Bangladesh",
        description: "The Utsob Jome Uthuk Regular Box blends Bangladeshi embroidery, Korean beauty, oud sprays, and gourmet snacks into a festive anniversary package. A rich mid-range celebration gift in Bangladesh.",
        keywords: "festive anniversary gift Bangladesh, embroidered handkerchief gift, oud room spray gift BD, Korean beauty gift Bangladesh, mid-range anniversary box, gourmet basket Bangladesh",
      },
      premium: {
        title: "Ekti Shundor Preetir Utsob – Premium Anniversary Gift Box | Custom BDT 2000+ | Bangladesh",
        description: "The Boro Utsob Premium Box celebrates love in grand style — Jamdani fabric, a luxury spa kit, LED photo frames, saffron hamper, and full personalisation. Bangladesh's most culturally rich anniversary gift.",
        keywords: "premium anniversary Bangladesh, Jamdani gift BD, saffron hamper anniversary, luxury spa gift Bangladesh, LED photo frame gift, anniversary celebration box BD",
      },
    },
  },
  {
    id: 6,
    name: "Shongshar Shajano Upoharer Baksho",
    tagline: "The Gift That Adorns a Home Built with Love",
    bangla_subtitle: "সংসার সাজানো উপহারের বাক্স",
    description: "A tribute to the couple who built a life together — this collection honours the home, the hearth, and the beautiful everyday rituals that make a marriage last. Inspired by Bangladeshi household warmth and aspiration.",
    tiers: [
      {
        tierName: "Mini — Shuru Hoye Ache",
        bangla: "শুরু হয়ে আছে | It Has Already Begun",
        tagline: "The smallest home is built with the biggest love",
        price: "BDT 700 – 999",
        priceNote: "Home-inspired gifts at accessible prices",
        hdrColor: C.hdrMini,
        bgColor: C.tier1,
        items: [
          "Couple's matching tea/coffee mug (2 pcs) — simple print (local)",
          "Jasmine or lavender scented sachet (2 pcs) — for wardrobe freshness",
          "Mini decorative photo frame — wooden (local craft market)",
          "Handwritten recipe card set — blank, for their shared recipes",
          "Small aromatic candle in tin — local brand",
          "Kraft gift box with tissue and twine",
        ],
        extras: null,
      },
      {
        tierName: "Regular — Ghore Phiri",
        bangla: "ঘরে ফিরি | Coming Home",
        tagline: "Every anniversary is a return — to love, to home",
        price: "BDT 1,200 – 1,999",
        priceNote: "Home comfort gifts, quality-forward",
        hdrColor: C.hdrReg,
        bgColor: C.tier2,
        items: [
          "Matching linen couple's apron set (local fabric/tailor)",
          "Aromatherapy diffuser — compact plug-in (Daraz, widely available)",
          "Premium tea gift set — assorted loose-leaf (Kazi & Kazi / local)",
          "Couple's name doormat or floormat — custom print (local vendor)",
          "Decorative potpourri bowl with dried petals",
          "Personalised couple's wall art — canvas print (local print shop)",
          "Gourmet snack basket — locally branded",
          "Premium white gift box with gold ribbon",
        ],
        extras: null,
      },
      {
        tierName: "Premium — Shopno Shongshar",
        bangla: "স্বপ্ন সংসার | The Dream Home",
        tagline: "For the couple who dreamed it together and built it together",
        price: "BDT 2,000 and above",
        priceNote: "Fully customisable luxury home & lifestyle bundle",
        hdrColor: C.hdrPrem,
        bgColor: C.tier3,
        items: [
          "Premium couple's matching bathrobe — embroidered names (local tailor / Aarong)",
          "Luxury scented candle — large statement piece (Amber/Sandalwood/Oud)",
          "Artisan ceramic couple's dinner set (2 bowls, 2 mugs, 2 plates) — local pottery",
          "Personalised cutting board — engraved couple's name + date (local woodworker)",
          "Complete kitchen herb garden kit — seeds, pots, soil (local nursery)",
          "Premium bedding gift set — cotton throw with couple's monogram",
          "Handwoven Nakshi Kantha cushion cover set — authentic BD craft",
          "Custom 'Our Home Rules' printed wall plaque",
          "Anniversary scrapbook — 'Our Life Together' (blank, premium quality)",
          "Grand wooden gift chest — padded interior, Bloom Bee branded",
        ],
        extras: [
          "Optional: Custom home portrait illustration (their house, painted)",
          "Optional: Indoor plant arrangement with ceramic pot",
          "Optional: Premium cookware add-on (non-stick set / stovetop)",
          "Optional: Couple's personalised linen bedsheet set",
        ],
      },
    ],
    meta: {
      mini: {
        title: "Shongshar Shajano Upoharer Baksho – Mini Anniversary Gift Box | BDT under 1000 | Bangladesh",
        description: "The Shuru Hoye Ache Mini Box celebrates the home you built — couple mugs, a photo frame, aromatic sachets, and a recipe card set in a charming kraft box. A budget-friendly anniversary gift in Bangladesh.",
        keywords: "couple mug gift Bangladesh, mini home gift BD, affordable anniversary box, kitchen anniversary gift BD, aromatic sachet gift, BDT 1000 couple gift",
      },
      regular: {
        title: "Shongshar Shajano Upoharer Baksho – Regular Anniversary Gift Box | BDT 1200–1999 | Bangladesh",
        description: "The Ghore Phiri Regular Box brings couple aprons, a tea gift set, aromatherapy diffuser, and personalised wall art in a home-inspired anniversary gift. Warm, meaningful, and mid-range priced in Bangladesh.",
        keywords: "couple apron gift Bangladesh, anniversary home gift, tea set anniversary BD, aromatherapy diffuser gift, wall art anniversary Bangladesh, mid-range anniversary 2000 BDT",
      },
      premium: {
        title: "Shongshar Shajano Upoharer Baksho – Premium Anniversary Gift Box | Custom BDT 2000+ | Bangladesh",
        description: "The Shopno Shongshar Premium Box honours the life built together — Nakshi Kantha cushions, artisan ceramics, engraved cutting boards, couple bathrobes, and full personalisation. Bangladesh's most heartfelt home anniversary gift.",
        keywords: "premium home anniversary gift Bangladesh, Nakshi Kantha gift, artisan ceramic set BD, engraved cutting board gift, bathrobe couple gift, luxury anniversary above 2000 BDT",
      },
    },
  },
];

// ═══════════════════════════════════════════════════════════════════════════
// BUILD DOCUMENT
// ═══════════════════════════════════════════════════════════════════════════

const children = [];

// ── Cover page ───────────────────────────────────────────────────────────────
children.push(
  new Paragraph({ ...sp(720, 0), children: [] }),
  centeredBold("🌸  BLOOM BEE  🌸", 48, C.brand),
  centeredBold("Anniversary Gift Box Collection", 36, C.gold),
  tagLine("Six Elegantly Named Collections — Each in Three Tiers"),
  tagLine("Mini (BDT ≤ 1,000)  |  Regular (BDT ≤ 2,000)  |  Premium (BDT 2,000+)"),
  divider(C.brand),
  new Paragraph({ ...sp(120, 0), children: [] }),
  body("This document presents Bloom Bee's flagship Anniversary Gift Box Catalogue for the Bangladesh market. Each of the six uniquely named collections offers three curated price tiers, culturally appropriate product selections, and complete SEO-ready meta information for your online shop.", false, true),
  new Paragraph({ ...sp(200, 0), children: [] }),
  body("All items listed are sourced or sourcelable in the Bangladesh market (Dhaka, Chittagong, Sylhet) through major retail channels including Meena Bazar, Shajgoj, Aarong, Daraz, and local craft vendors.", false, true, "555555"),
  pageBreak(),
);

// ── Table of Contents ─────────────────────────────────────────────────────
children.push(
  heading1("Table of Contents"),
  ...boxes.map((b, i) => new Paragraph({
    children: [
      new TextRun({ text: `${i + 1}.  `, font: "Arial", size: 22, bold: true, color: C.brand }),
      new TextRun({ text: b.name, font: "Arial", size: 22, bold: true, color: "222222" }),
      new TextRun({ text: `  —  "${b.tagline}"`, font: "Arial", size: 20, italic: true, color: "666666" }),
    ],
    ...sp(100, 100),
  })),
  pageBreak(),
);

// ── Per-box sections ──────────────────────────────────────────────────────
boxes.forEach((box, idx) => {
  // Box header
  children.push(
    heading1(`${idx + 1}. ${box.name}`),
    centeredBold(`"${box.tagline}"`, 24, C.gold),
    tagLine(box.bangla_subtitle),
    body(box.description, false, true, "444444"),
    divider(),
  );

  // Three tiers
  box.tiers.forEach((tier, ti) => {
    children.push(...tierSection(tier));
    if (ti < 2) children.push(divider(C.border));
  });

  children.push(divider(C.gold));

  // Meta info for all 3 tiers
  children.push(heading2("SEO & Shop Meta Information", C.gold));
  [
    ["Mini Package", box.meta.mini],
    ["Regular Package", box.meta.regular],
    ["Premium Package", box.meta.premium],
  ].forEach(([label, meta]) => {
    children.push(heading3(label, C.brand), ...metaSection(box.name, meta));
  });

  // Page break between boxes (not after last)
  if (idx < boxes.length - 1) children.push(pageBreak());
});

// ── Final notes ───────────────────────────────────────────────────────────
children.push(
  pageBreak(),
  heading1("Packaging & Sourcing Notes"),
  heading2("Box Packaging Standards", C.brand),
  bullet("Mini Tier: Kraft paper box / organza bag / jute bag — printed Bloom Bee sticker seal"),
  bullet("Regular Tier: Rigid gift box with magnetic closure — satin ribbon, tissue filler"),
  bullet("Premium Tier: Luxury wooden chest or rigid velvet-lined gift chest — branded Bloom Bee ribbon, wax seal, personalised tag"),
  new Paragraph({ ...sp(120, 60), children: [] }),
  heading2("Recommended Sourcing Channels", C.brand),
  bullet("Skincare & Beauty: Meena Bazar, Shajgoj, Daraz, Lavish (local)"),
  bullet("Candles & Diffusers: Local craft vendors, Daraz imports, Décor shops Dhanmondi/Gulshan"),
  bullet("Chocolates & Snacks: Cadbury, Bôn Bon, imported via Daraz, local sweet shops"),
  bullet("Ceramics & Crafts: Aarong, local pottery shops, Banani craft market"),
  bullet("Personalised Printing: Local print shops (photo books, frames, cushions, mugs, canvas)"),
  bullet("Jewellery: Local silver-plated vendors, Gulshan-2 / Bashundhara City shops"),
  bullet("Fabric & Textile: Aarong, Kay Kraft, local boutique tailors, Jamdani vendors"),
  bullet("3D Printing / Custom: Dhaka-based 3D printing vendors (check Facebook Marketplace)"),
  new Paragraph({ ...sp(200, 60), children: [] }),
  centeredBold("🌸  Bloom Bee — Gifting with Heart, Rooted in Bangladesh  🌸", 24, C.brand),
  tagLine("For custom orders, wholesale enquiries, or partnership: hello@bloombee.com.bd"),
);

// ═══════════════════════════════════════════════════════════════════════════
// ASSEMBLE
// ═══════════════════════════════════════════════════════════════════════════

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        }],
      },
    ],
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Arial", color: C.brand },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0, alignment: AlignmentType.CENTER } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: C.brand },
        paragraph: { spacing: { before: 300, after: 120 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "Arial", color: "333333" },
        paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 2 } },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      },
    },
    children,
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("BloomBee_Anniversary_Gift_Boxes.docx", buf);
  console.log("Done!");
});
