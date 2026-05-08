/**
 * generate_cv.js
 * ──────────────
 * Erzeugt einen professionellen Lebenslauf im Zetsche-Stil:
 *   - Zweispaltig: linke Sidebar (farbig) + rechte Hauptspalte (weiß)
 *   - Oben in der Sidebar: Platz für Profilbild
 *   - Farbpalette vollständig konfigurierbar via JSON
 *
 * Aufruf:
 *   node generate_cv.js cv_data.json [palette.json] output.docx
 */

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, HeadingLevel, LevelFormat, UnderlineType, HeightRule,
} = require('docx');
const fs   = require('fs');
const path = require('path');

// ─────────────────────────────────────────────
// Farbpaletten
// ─────────────────────────────────────────────

const PALETTES = {
  // Original Zetsche-Grün
  green: {
    sidebarBg:      '0A5C46',
    sidebarText:    'FFFFFF',
    sidebarMuted:   '6FCFAD',
    sidebarAccent:  '2EB88A',
    sidebarSubhead: 'D6F0E8',
    mainBg:         'FFFFFF',
    mainText:       '1F2937',
    mainMuted:      '6B7280',
    mainAccent:     '0A5C46',
    mainHeading:    '164D3A',
    dotFill:        '2EB88A',
    dotEmpty:       '6FCFAD',
    divider:        '2EB88A',
    tagBg:          'D6F0E8',
    tagText:        '0A5C46',
  },
  // Klassisches Navy-Blau
  navy: {
    sidebarBg:      '1E3A5F',
    sidebarText:    'FFFFFF',
    sidebarMuted:   '90B4D4',
    sidebarAccent:  '4A90D9',
    sidebarSubhead: 'D0E4F7',
    mainBg:         'FFFFFF',
    mainText:       '1F2937',
    mainMuted:      '6B7280',
    mainAccent:     '1E3A5F',
    mainHeading:    '0D2035',
    dotFill:        '4A90D9',
    dotEmpty:       '90B4D4',
    divider:        '4A90D9',
    tagBg:          'D0E4F7',
    tagText:        '1E3A5F',
  },
  // Warmes Bordeaux
  bordeaux: {
    sidebarBg:      '6B1A2A',
    sidebarText:    'FFFFFF',
    sidebarMuted:   'D4909C',
    sidebarAccent:  'C0435A',
    sidebarSubhead: 'F5D5DB',
    mainBg:         'FFFFFF',
    mainText:       '1F2937',
    mainMuted:      '6B7280',
    mainAccent:     '6B1A2A',
    mainHeading:    '4A0E1A',
    dotFill:        'C0435A',
    dotEmpty:       'D4909C',
    divider:        'C0435A',
    tagBg:          'F5D5DB',
    tagText:        '6B1A2A',
  },
  // Schiefergrau (minimalistisch)
  slate: {
    sidebarBg:      '2D3748',
    sidebarText:    'FFFFFF',
    sidebarMuted:   'A0AEC0',
    sidebarAccent:  '68D391',
    sidebarSubhead: 'E2E8F0',
    mainBg:         'FFFFFF',
    mainText:       '1F2937',
    mainMuted:      '6B7280',
    mainAccent:     '2D3748',
    mainHeading:    '1A202C',
    dotFill:        '68D391',
    dotEmpty:       'A0AEC0',
    divider:        '68D391',
    tagBg:          'E2E8F0',
    tagText:        '2D3748',
  },
  // Warmes Cognac/Terrakotta
  cognac: {
    sidebarBg:      '7B3F1E',
    sidebarText:    'FFFFFF',
    sidebarMuted:   'D4A574',
    sidebarAccent:  'E07B39',
    sidebarSubhead: 'F5E6D3',
    mainBg:         'FFFFFF',
    mainText:       '1F2937',
    mainMuted:      '6B7280',
    mainAccent:     '7B3F1E',
    mainHeading:    '5A2D0C',
    dotFill:        'E07B39',
    dotEmpty:       'D4A574',
    divider:        'E07B39',
    tagBg:          'F5E6D3',
    tagText:        '7B3F1E',
  },
};

// ─────────────────────────────────────────────
// Maße (DXA: 1440 = 1 inch, A4)
// ─────────────────────────────────────────────

const PAGE_W    = 11906;  // A4 Breite
const PAGE_H    = 16838;  // A4 Höhe
const MARGIN    = 0;      // Kein äußerer Rand (Tabelle füllt alles)

const SIDEBAR_W = 3200;   // Linke Spalte
const MAIN_W    = PAGE_W - SIDEBAR_W;  // Rechte Spalte

const PHOTO_SIZE_EMU = 1080000;  // ~3cm in EMU (914400 = 1 inch)
const PHOTO_AREA_PT  = 130;      // Höhe des Foto-Bereichs in pt (als Leerabstand)

// Padding innerhalb der Zellen (DXA)
const SIDE_PAD_H = 280;  // horizontal
const SIDE_PAD_V = 200;  // vertikal
const MAIN_PAD_H = 360;
const MAIN_PAD_V = 240;

// ─────────────────────────────────────────────
// Hilfsfunktionen
// ─────────────────────────────────────────────

const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

function hex(c) { return c.replace('#', ''); }

/** Leerer Absatz mit konfigurierbarer Größe */
function spacer(sizeHalfPt = 0, color = null) {
  return new Paragraph({
    children: [new TextRun({ text: '', size: sizeHalfPt || 2, color: color })],
    spacing: { before: 0, after: 0 },
  });
}

/** Trennlinie als Paragraph-Border */
function dividerLine(palette) {
  return new Paragraph({
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 6, color: palette.sidebarAccent, space: 1 },
    },
    spacing: { before: 100, after: 100 },
    children: [new TextRun('')],
  });
}

function mainDivider(palette) {
  return new Paragraph({
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 4, color: palette.mainAccent, space: 1 },
    },
    spacing: { before: 80, after: 120 },
    children: [new TextRun('')],
  });
}

/** Skill-Punkte (5 Kreise, gefüllt je nach Level 1–5) */
function skillDots(level, palette) {
  const filled = '●';
  const empty  = '●';
  const dots = [];
  for (let i = 0; i < 5; i++) {
    dots.push(new TextRun({
      text: (i < level ? filled : empty) + ' ',
      color: i < level ? palette.dotFill : palette.dotEmpty,
      size: 14,
    }));
  }
  return dots;
}

/** Layout-Engine: Schriftgröße basierend auf Textlänge */
function adaptFontSize(text, maxLen, sizeDefault, sizeMin) {
  const len = (text || '').length;
  if (len <= maxLen) return sizeDefault;
  const t = Math.min((len - maxLen) / maxLen, 1);
  return Math.max(sizeDefault - t * (sizeDefault - sizeMin), sizeMin);
}

/** Spacing-Entscheidung basierend auf Content-Score */
function adaptSpacing(contentScore) {
  const LOW = 300, HIGH = 900;
  if (contentScore < LOW) return { before: 180, after: 120, sectionBefore: 240 };
  if (contentScore > HIGH) return { before: 60,  after: 40,  sectionBefore: 100 };
  const t = (contentScore - LOW) / (HIGH - LOW);
  return {
    before:        Math.round(180 - t * 120),
    after:         Math.round(120 - t * 80),
    sectionBefore: Math.round(240 - t * 140),
  };
}

/** Berechnet einen einfachen Content-Score */
function calcContentScore(data) {
  let score = 0;
  score += (data.name || '').length + (data.title || '').length + (data.summary || '').length;
  (data.experience || []).forEach(j => { score += 60 + (j.description || '').length; });
  (data.education  || []).forEach(() => { score += 40; });
  (data.skills     || []).forEach(s => { score += 30 + s.items.length * 5; });
  if (data.photo) score += 100;
  return score;
}

// ─────────────────────────────────────────────
// Sidebar-Inhalte
// ─────────────────────────────────────────────

function buildPhotoArea(data, palette, sp) {
  const children = [];

  if (data.photo && fs.existsSync(data.photo)) {
    // Bild einbetten (quadratisch zugeschnitten via Pillow, s. prepare_photo.py)
    const imgData = fs.readFileSync(data.photo);
    const ext = path.extname(data.photo).toLowerCase().replace('.', '');
    const imgType = ext === 'png' ? 'png' : 'jpeg';
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 160, after: 120 },
      children: [new ImageRun({
        data: imgData,
        transformation: { width: 120, height: 120 },  // pt
        type: imgType,
      })],
    }));
  } else {
    // Platzhalter-Fläche
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 0 },
      children: [new TextRun({ text: '', size: PHOTO_AREA_PT * 2 })],
    }));
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 80 },
      children: [new TextRun({
        text: '[ Profilbild ]',
        size: 16,
        color: palette.sidebarMuted,
        italics: true,
      })],
    }));
  }

  return children;
}

function buildSidebarName(data, palette, sp) {
  const nameParts = (data.name || '').split(' ');
  const lastName  = nameParts.pop() || '';
  const firstName = nameParts.join(' ');

  const nameFontSize = adaptFontSize(data.name, 18, 36, 26);

  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 60, after: 0 },
      children: [
        ...(firstName ? [new TextRun({
          text: firstName + ' ',
          size: nameFontSize,
          color: palette.sidebarText,
          font: 'Calibri',
        })] : []),
        new TextRun({
          text: lastName,
          size: nameFontSize,
          bold: true,
          color: palette.sidebarText,
          font: 'Calibri',
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 40, after: 100 },
      children: [new TextRun({
        text: data.title || '',
        size: 20,
        italics: true,
        color: palette.sidebarMuted,
        font: 'Calibri',
      })],
    }),
    dividerLine(palette),
  ];
}

function sidebarSectionHeading(title, palette) {
  return new Paragraph({
    spacing: { before: 200, after: 80 },
    children: [new TextRun({
      text: title.toUpperCase(),
      size: 17,
      bold: true,
      color: palette.sidebarSubhead,
      characterSpacing: 40,
      font: 'Calibri',
    })],
  });
}

function buildContactSection(data, palette) {
  const items = [
    { icon: '✉', value: data.email },
    { icon: '✆', value: data.phone },
    { icon: '⌂', value: data.location },
    { icon: 'in', value: data.linkedin },
  ].filter(i => i.value);

  return [
    sidebarSectionHeading('Kontakt', palette),
    ...items.map(item => new Paragraph({
      spacing: { before: 40, after: 40 },
      children: [
        new TextRun({ text: item.icon + '  ', size: 17, color: palette.sidebarAccent, font: 'Calibri' }),
        new TextRun({ text: item.value, size: 17, color: palette.sidebarText, font: 'Calibri' }),
      ],
    })),
  ];
}

function buildSkillsSection(data, palette) {
  if (!data.skills || data.skills.length === 0) return [];
  const result = [dividerLine(palette), sidebarSectionHeading('Kernkompetenzen', palette)];

  data.skills.forEach(cat => {
    if (cat.items && cat.items.length > 0 && typeof cat.items[0] === 'object') {
      // Mit Level-Angabe → Punkte
      result.push(new Paragraph({
        spacing: { before: 20, after: 4 },
        children: [new TextRun({
          text: cat.category,
          size: 17, bold: true, color: palette.sidebarSubhead, font: 'Calibri',
        })],
      }));
      cat.items.forEach(skill => {
        result.push(new Paragraph({
          spacing: { before: 30, after: 20 },
          children: [
            new TextRun({ text: (skill.name || skill) + '  ', size: 17, color: palette.sidebarText, font: 'Calibri' }),
            ...skillDots(skill.level || 4, palette),
          ],
        }));
      });
    } else {
      // Einfache Listen
      result.push(new Paragraph({
        spacing: { before: 50, after: 20 },
        children: [new TextRun({
          text: cat.category + ':',
          size: 17, bold: true, color: palette.sidebarSubhead, font: 'Calibri',
        })],
      }));
      result.push(new Paragraph({
        spacing: { before: 0, after: 40 },
        children: [new TextRun({
          text: cat.items.join('  ·  '),
          size: 16, color: palette.sidebarText, font: 'Calibri',
        })],
      }));
    }
  });

  return result;
}

function buildToolsSection(data, palette) {
  if (!data.tools || data.tools.length === 0) return [];
  return [
    dividerLine(palette),
    sidebarSectionHeading('Tools', palette),
    new Paragraph({
      spacing: { before: 60, after: 60 },
      children: data.tools.map((t, i) => new TextRun({
        text: t + (i < data.tools.length - 1 ? '  ' : ''),
        size: 17,
        bold: true,
        color: palette.sidebarAccent,
        font: 'Calibri',
      })),
    }),
  ];
}

function buildLanguagesSection(data, palette) {
  if (!data.languages || data.languages.length === 0) return [];
  return [
    dividerLine(palette),
    sidebarSectionHeading('Sprachen', palette),
    ...data.languages.map(lang => new Paragraph({
      spacing: { before: 40, after: 40 },
      children: [
        new TextRun({ text: lang.name + ' — ', size: 17, bold: true, color: palette.sidebarText, font: 'Calibri' }),
        new TextRun({ text: lang.level, size: 17, color: palette.sidebarMuted, font: 'Calibri' }),
      ],
    })),
  ];
}

function buildEducationSidebarSection(data, palette) {
  if (!data.education || data.education.length === 0) return [];
  return [
    dividerLine(palette),
    sidebarSectionHeading('Ausbildung', palette),
    ...data.education.map(edu => [
      new Paragraph({
        spacing: { before: 60, after: 20 },
        children: [new TextRun({ text: edu.degree, size: 18, bold: true, color: palette.sidebarText, font: 'Calibri' })],
      }),
      new Paragraph({
        spacing: { before: 0, after: 40 },
        children: [
          new TextRun({ text: edu.institution, size: 16, color: palette.sidebarMuted, font: 'Calibri' }),
          new TextRun({ text: '  ·  ' + edu.year, size: 16, color: palette.sidebarMuted, font: 'Calibri' }),
        ],
      }),
      ...(edu.details ? [new Paragraph({
        spacing: { before: 0, after: 40 },
        children: [new TextRun({ text: edu.details, size: 15, italics: true, color: palette.sidebarMuted, font: 'Calibri' })],
      })] : []),
    ]).flat(),
  ];
}

// ─────────────────────────────────────────────
// Hauptspalten-Inhalte
// ─────────────────────────────────────────────

function mainSectionHeading(title, palette) {
  return [
    new Paragraph({
      spacing: { before: 200, after: 40 },
      children: [new TextRun({
        text: title.toUpperCase(),
        size: 22,
        bold: true,
        color: palette.mainHeading,
        characterSpacing: 60,
        font: 'Calibri',
      })],
    }),
    new Paragraph({
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: palette.mainAccent, space: 1 } },
      spacing: { before: 0, after: 120 },
      children: [new TextRun('')],
    }),
  ];
}

function buildSummary(data, palette, sp) {
  if (!data.summary) return [];
  const summarySize = adaptFontSize(data.summary, 250, 20, 17);
  return [
    ...mainSectionHeading('Profil', palette),
    new Paragraph({
      spacing: { before: 0, after: sp.sectionBefore },
      children: [new TextRun({
        text: data.summary,
        size: summarySize,
        italics: true,
        color: palette.mainMuted,
        font: 'Calibri',
      })],
    }),
  ];
}

function buildExperience(data, palette, sp) {
  if (!data.experience || data.experience.length === 0) return [];
  const result = [...mainSectionHeading('Berufserfahrung', palette)];

  data.experience.forEach((job, idx) => {
    const descSize = adaptFontSize(job.description, 200, 19, 16);
    const isLast = idx === data.experience.length - 1;

    // Jobtitel
    result.push(new Paragraph({
      spacing: { before: idx === 0 ? 0 : sp.sectionBefore, after: 30 },
      children: [new TextRun({
        text: job.title,
        size: 22, bold: true, color: palette.mainText, font: 'Calibri',
      })],
    }));

    // Firma + Zeitraum
    const period = [job.start, job.end || 'heute'].filter(Boolean).join(' – ');
    result.push(new Paragraph({
      spacing: { before: 0, after: 60 },
      children: [
        new TextRun({ text: job.company, size: 19, bold: true, color: palette.mainAccent, font: 'Calibri' }),
        new TextRun({ text: '  ·  ', size: 19, color: palette.mainMuted, font: 'Calibri' }),
        new TextRun({ text: period, size: 19, italics: true, color: palette.mainMuted, font: 'Calibri' }),
        ...(job.location ? [new TextRun({ text: '  ·  ' + job.location, size: 19, color: palette.mainMuted, font: 'Calibri' })] : []),
      ],
    }));

    // KPI-Badges (falls vorhanden)
    if (job.kpis && job.kpis.length > 0) {
      result.push(new Paragraph({
        spacing: { before: 40, after: 60 },
        children: job.kpis.map(kpi => [
          new TextRun({ text: kpi.value + ' ', size: 24, bold: true, color: palette.mainAccent, font: 'Calibri' }),
          new TextRun({ text: kpi.label + '   ', size: 17, color: palette.mainMuted, font: 'Calibri' }),
        ]).flat(),
      }));
    }

    // Beschreibung / Bullet-Punkte
    if (job.description) {
      // Wenn mehrzeilig → als Bullets
      const lines = job.description.split('\n').map(l => l.trim()).filter(Boolean);
      if (lines.length > 1) {
        lines.forEach(line => {
          result.push(new Paragraph({
            spacing: { before: sp.before / 2, after: sp.after / 2 },
            indent: { left: 240, hanging: 200 },
            children: [
              new TextRun({ text: '▸  ', size: descSize, color: palette.mainAccent, font: 'Calibri' }),
              new TextRun({ text: line, size: descSize, color: palette.mainText, font: 'Calibri' }),
            ],
          }));
        });
      } else {
        result.push(new Paragraph({
          spacing: { before: 0, after: sp.after },
          children: [new TextRun({ text: lines[0], size: descSize, color: palette.mainText, font: 'Calibri' })],
        }));
      }
    }

    // Trennlinie zwischen Jobs (nicht nach dem letzten)
    if (!isLast) {
      result.push(new Paragraph({
        border: { bottom: { style: BorderStyle.DASHED, size: 3, color: palette.tagBg, space: 1 } },
        spacing: { before: 80, after: 0 },
        children: [new TextRun('')],
      }));
    }
  });

  return result;
}

function buildEducationMain(data, palette, sp) {
  // Ausbildung ist in der Sidebar – hier optional ein Zusatzblock
  if (!data.educationMain) return [];
  return [
    ...mainSectionHeading('Weiterbildung', palette),
    ...data.educationMain.map(item => new Paragraph({
      spacing: { before: sp.before, after: sp.after },
      children: [
        new TextRun({ text: item.title + '  ', size: 19, bold: true, color: palette.mainText, font: 'Calibri' }),
        new TextRun({ text: item.org + '  ·  ' + item.year, size: 18, color: palette.mainMuted, font: 'Calibri' }),
      ],
    })),
  ];
}

// ─────────────────────────────────────────────
// Haupt-Build-Funktion
// ─────────────────────────────────────────────

async function buildCV(data, palette, outputPath) {
  const score = calcContentScore(data);
  const sp    = adaptSpacing(score);

  console.log(`Content-Score: ${score} → ${score < 300 ? 'sparse' : score > 900 ? 'dense' : 'balanced'} layout`);
  console.log(`Abstände: before=${sp.before}, after=${sp.after}, sectionBefore=${sp.sectionBefore}`);

  // ── Sidebar-Inhalt ──
  const sidebarContent = [
    ...buildPhotoArea(data, palette, sp),
    ...buildSidebarName(data, palette, sp),
    ...buildContactSection(data, palette),
    ...buildSkillsSection(data, palette),
    ...buildToolsSection(data, palette),
    ...buildLanguagesSection(data, palette),
    ...buildEducationSidebarSection(data, palette),
    spacer(40, palette.sidebarBg),  // Abschluss-Puffer
  ];

  // ── Hauptspalten-Inhalt ──
  const mainContent = [
    spacer(60),
    ...buildSummary(data, palette, sp),
    ...buildExperience(data, palette, sp),
    ...buildEducationMain(data, palette, sp),
  ];

  // ── Tabelle (Layout-Träger) ──
  // Zeilenhöhe = exakt A4-Seitenhöhe (in Twips: DXA).
  // "exact" verhindert, dass Word die Zeile bei mehr Inhalt aufbläht
  // und eine Leerseite erzeugt. Der Inhalt wird ggf. abgeschnitten,
  // aber bei einem einseitigen CV passt alles auf eine Seite.
  const ROW_HEIGHT_DXA = PAGE_H; // 16838 Twips = A4-Höhe

  const layoutTable = new Table({
    width: { size: PAGE_W, type: WidthType.DXA },
    columnWidths: [SIDEBAR_W, MAIN_W],
    rows: [
      new TableRow({
        // Zeile exakt auf Seitenhöhe fixieren → Sidebar-Farbe füllt immer durch
        height: { value: ROW_HEIGHT_DXA, rule: HeightRule.EXACT },
        children: [
          // Linke Sidebar
          new TableCell({
            width: { size: SIDEBAR_W, type: WidthType.DXA },
            borders: noBorders,
            verticalAlign: VerticalAlign.TOP,
            shading: { fill: palette.sidebarBg, type: ShadingType.CLEAR },
            margins: { top: SIDE_PAD_V, bottom: SIDE_PAD_V, left: SIDE_PAD_H, right: SIDE_PAD_H },
            children: sidebarContent,
          }),
          // Rechte Hauptspalte
          new TableCell({
            width: { size: MAIN_W, type: WidthType.DXA },
            borders: noBorders,
            verticalAlign: VerticalAlign.TOP,
            shading: { fill: palette.mainBg, type: ShadingType.CLEAR },
            margins: { top: MAIN_PAD_V, bottom: MAIN_PAD_V, left: MAIN_PAD_H, right: MAIN_PAD_H },
            children: mainContent,
          }),
        ],
      }),
    ],
  });

  // ── Dokument ──
  const doc = new Document({
    styles: {
      default: {
        document: { run: { font: 'Calibri', size: 20, color: '1F2937' } },
      },
    },
    sections: [{
      properties: {
        page: {
          size: { width: PAGE_W, height: PAGE_H },
          margin: { top: 0, right: 0, bottom: 0, left: 0 },
        },
      },
      children: [layoutTable],
    }],
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buffer);
  console.log(`✓ CV gespeichert: ${outputPath}`);
  console.log(`  Palette: ${JSON.stringify({ sidebarBg: palette.sidebarBg, mainAccent: palette.mainAccent })}`);
}

// ─────────────────────────────────────────────
// CLI-Einstiegspunkt
// ─────────────────────────────────────────────

async function main() {
  const args = process.argv.slice(2);
  if (args.length < 2) {
    console.error('Verwendung: node generate_cv.js <cv_data.json> <output.docx> [palette_name_oder_custom.json]');
    process.exit(1);
  }

  const dataPath    = args[0];
  const outputPath  = args[1];
  const paletteArg  = args[2] || 'green';

  const data = JSON.parse(fs.readFileSync(dataPath, 'utf8'));

  let palette;
  if (PALETTES[paletteArg]) {
    palette = PALETTES[paletteArg];
    console.log(`Palette: ${paletteArg}`);
  } else if (fs.existsSync(paletteArg)) {
    palette = { ...PALETTES.green, ...JSON.parse(fs.readFileSync(paletteArg, 'utf8')) };
    console.log(`Custom-Palette: ${paletteArg}`);
  } else {
    console.warn(`Unbekannte Palette "${paletteArg}", verwende "green"`);
    palette = PALETTES.green;
  }

  await buildCV(data, palette, outputPath);
}

main().catch(err => { console.error(err); process.exit(1); });

// Paletten auch exportieren (für app.py)
module.exports = { PALETTES, buildCV };
