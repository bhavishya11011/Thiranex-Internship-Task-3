const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  LevelFormat, VerticalAlign, PageNumber, PageBreak
} = require('docx');
const fs = require('fs');

// ── COLOURS ──────────────────────────────────────────────
const GREEN_DARK  = "1B4332";
const GREEN_MID   = "40916C";
const GREEN_LIGHT = "D8F3DC";
const GREEN_PALE  = "F0FAF2";
const AMBER       = "D97706";
const RED         = "B91C1C";
const RED_LIGHT   = "FEE2E2";
const GRAY        = "6B7280";
const WHITE       = "FFFFFF";
const BORDER_COL  = "D1D5DB";

// ── USER TEST DATA (expert analysis of the feast4U prototype) ──────────
const personas = [
  {
    name: "Priya, 21",
    role: "College Student · First-Time User",
    emoji: "👩‍🎓",
    score: 7,
    scores: { "First Impression": 8, "Navigation Ease": 7, "Visual Clarity": 7, "Task Completion": 6, "Trust & Safety": 6 },
    issues: [
      "Add+ button flashes 'Added' then reverts — no persistent cart count feedback",
      "Food chips (Pizza, Burger etc.) are not tappable — looks like a filter but does nothing",
      "Only one restaurant is tappable; Spice Garden is a dead end — confusing",
      "No payment screen before Order Confirmed — feels like the order was accidentally placed",
    ],
    praises: [
      "Splash screen is visually appealing and loads fast",
      "Green colour theme feels fresh and food-appropriate",
      "Cart quantity +/− controls work smoothly and update the bill in real time",
    ],
    suggestions: [
      "Add a persistent cart badge (item count) on the Restaurant screen toolbar",
      "Make food category chips functional or remove them to avoid false affordance",
      "Insert a payment method selection screen before Order Confirmed",
    ],
    quote: "I tapped 'Add' for the pizza but nothing changed in the cart icon — did it even work? And why can't I tap Spice Garden?",
    summary: "Priya completed the core flow quickly but felt uncertain at two points: when items were added to cart (no visual confirmation persisted) and when payment was skipped. She rated the visual design positively but expects more interactive feedback from a food app she'd use daily."
  },
  {
    name: "Rahul, 34",
    role: "Office Professional · Busy User",
    emoji: "👨‍💼",
    score: 6,
    scores: { "First Impression": 7, "Navigation Ease": 6, "Visual Clarity": 7, "Task Completion": 6, "Trust & Safety": 5 },
    issues: [
      "No payment confirmation or UPI/card selection — trust concern for a professional",
      "Cart is pre-filled with items he didn't add, breaking his mental model",
      "No search autocomplete — inefficient for a user who knows what he wants",
      "Bottom nav Search, Cart, Profile tabs are inactive — feels like a broken product",
      "No restaurant filter (cuisine, rating, delivery time) for quick decision-making",
    ],
    praises: [
      "Restaurant card layout is clean and scannable — rating, time, price at a glance",
      "Bill Summary with itemised prices builds confidence",
      "Order tracking steps on the confirmation screen are reassuring",
    ],
    suggestions: [
      "Add functional search with autocomplete and cuisine filters on Home",
      "Implement a real payment step (UPI, card, cash on delivery options)",
      "Activate the bottom nav tabs to let users jump between sections",
    ],
    quote: "Where's the payment screen? I'd never confirm a Rs 180 order without entering my UPI PIN. Also the cart already had items in it — very weird.",
    summary: "Rahul found the visual structure acceptable but had serious trust concerns about the missing payment step. As a professional who values efficiency, the non-functional navigation tabs and pre-filled cart made the prototype feel incomplete rather than polished."
  },
  {
    name: "Ramesh, 58",
    role: "Retired · Low Tech-Literacy",
    emoji: "👴",
    score: 5,
    scores: { "First Impression": 6, "Navigation Ease": 4, "Visual Clarity": 5, "Task Completion": 5, "Trust & Safety": 4 },
    issues: [
      "Icon-only bottom nav (no text in early screens) — unclear what Search, Cart, Profile do",
      "Login screen has no 'Show password' toggle — hard for older users to verify input",
      "No 'forgot password' link — he would stop here if password is forgotten",
      "Food emojis without descriptions (menu item details) don't communicate dietary info",
      "Order Confirmed without a payment step is alarming — he thinks the app charged him automatically",
    ],
    praises: [
      "Large readable font sizes on most screens",
      "Colour-coded delivery tracker (green = done, grey = pending) is intuitive",
      "Back buttons are clearly placed on Restaurant and Cart screens",
    ],
    suggestions: [
      "Add 'Forgot password?' link on the Login screen",
      "Add a 'Show password' toggle on the password field",
      "Add brief text labels to all bottom navigation icons",
      "Include a clear payment confirmation dialog so users know before being charged",
    ],
    quote: "Yeh order ho gaya? Maine toh kuch bhi pay nahi kiya — ab mere account se paise jayenge kya? (Did the order go through? I didn't pay anything — will money be deducted from my account?)",
    summary: "Ramesh was able to navigate the splash and login screens with some hesitation, but struggled significantly from the restaurant menu onward. The lack of a payment step caused genuine alarm. He would not trust this app for real orders without clearer confirmation dialogs and a visible payment step."
  },
  {
    name: "Sneha, 27",
    role: "UX Designer · Expert Reviewer",
    emoji: "👩‍💻",
    score: 6,
    scores: { "First Impression": 8, "Navigation Ease": 6, "Visual Clarity": 7, "Task Completion": 5, "Trust & Safety": 5 },
    issues: [
      "Violated Nielsen Heuristic #1 (System Status): Add+ button gives no persistent feedback",
      "Violated Heuristic #3 (User Control): No way to remove items from cart or undo add",
      "Violated Heuristic #7 (Flexibility): No shortcuts — power users can't reorder or search by cuisine",
      "Empty states missing: Search, Profile, and Cart tabs show nothing when tapped",
      "Colour contrast on splash subtitle (#A5D6A7 on dark green) may fail WCAG AA",
      "Pre-filled cart breaks the prototype's internal consistency with the menu add flow",
    ],
    praises: [
      "Strong visual identity — green palette is consistent and on-brand throughout",
      "Order confirmation screen has excellent information hierarchy",
      "Micro-interaction on Add+ (flash to Added) is a good idea, just needs persistence",
    ],
    suggestions: [
      "Implement persistent cart state: badge count on nav icon, items persist between screens",
      "Add empty states and error states for all navigable tabs",
      "Run a WCAG contrast check on all text/background combinations",
      "Add item removal (swipe or delete icon) in the Cart screen",
      "Include a payment method selection screen — critical for user trust",
    ],
    quote: "The visual design language is solid — consistent green theme, good type hierarchy. But the interaction model has gaps: no persistent cart, dead navigation tabs, and a missing payment step. This needs another round of prototyping before usability testing with real users.",
    summary: "Sneha appreciated the visual consistency and branding but identified multiple heuristic violations that would create real friction for users. She flagged the pre-filled cart as a prototype integrity issue and recommended adding a payment screen and fixing all non-functional interactive elements before the next test round."
  }
];

// ── HELPERS ─────────────────────────────────────────────
const border = (col = BORDER_COL) => ({ style: BorderStyle.SINGLE, size: 1, color: col });
const noBorder = () => ({ style: BorderStyle.NONE, size: 0, color: WHITE });
const cellBorders = (col) => ({ top: border(col), bottom: border(col), left: border(col), right: border(col) });
const noBorders = () => ({ top: noBorder(), bottom: noBorder(), left: noBorder(), right: noBorder() });

function para(children, opts = {}) {
  return new Paragraph({ children, ...opts });
}
function run(text, opts = {}) {
  return new TextRun({ text, font: "Arial", ...opts });
}
function hRule() {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "D1D5DB", space: 1 } },
    spacing: { before: 160, after: 160 }
  });
}
function spacer(before = 80, after = 80) {
  return new Paragraph({ children: [], spacing: { before, after } });
}

// Score colour
function scoreColor(s) {
  if (s >= 7) return "15803D";
  if (s >= 5) return AMBER;
  return RED;
}
function scoreBg(s) {
  if (s >= 7) return "DCFCE7";
  if (s >= 5) return "FEF9C3";
  return RED_LIGHT;
}

// Rating bar as text (ASCII)
function ratingBar(v) {
  const filled = Math.round(v);
  return "█".repeat(filled) + "░".repeat(10 - filled) + `  ${v}/10`;
}

// ── TITLE PAGE ───────────────────────────────────────────
function titlePage() {
  return [
    spacer(720, 0),
    para([run("feast4U", { size: 56, bold: true, color: WHITE })], {
      alignment: AlignmentType.CENTER,
      shading: { fill: GREEN_DARK, type: ShadingType.CLEAR },
      border: { bottom: { style: BorderStyle.NONE } },
      spacing: { before: 0, after: 0 },
    }),
    para([run("User Testing Report", { size: 36, bold: true, color: GREEN_LIGHT })], {
      alignment: AlignmentType.CENTER,
      shading: { fill: GREEN_DARK, type: ShadingType.CLEAR },
      spacing: { before: 0, after: 0 },
    }),
    para([run("Hi-Fidelity Prototype  ·  6-Screen Flow Evaluation", { size: 22, color: "A5D6A7" })], {
      alignment: AlignmentType.CENTER,
      shading: { fill: GREEN_DARK, type: ShadingType.CLEAR },
      spacing: { before: 0, after: 280 },
    }),
    spacer(560, 560),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [3120, 3120, 3120],
      rows: [
        new TableRow({
          children: [
            new TableCell({
              borders: cellBorders(GREEN_MID),
              width: { size: 3120, type: WidthType.DXA },
              shading: { fill: GREEN_PALE, type: ShadingType.CLEAR },
              margins: { top: 120, bottom: 120, left: 160, right: 160 },
              children: [
                para([run("Personas Tested", { size: 18, color: GRAY })], { alignment: AlignmentType.CENTER }),
                para([run("4", { size: 48, bold: true, color: GREEN_DARK })], { alignment: AlignmentType.CENTER }),
              ]
            }),
            new TableCell({
              borders: cellBorders(GREEN_MID),
              width: { size: 3120, type: WidthType.DXA },
              shading: { fill: GREEN_PALE, type: ShadingType.CLEAR },
              margins: { top: 120, bottom: 120, left: 160, right: 160 },
              children: [
                para([run("Avg. Overall Score", { size: 18, color: GRAY })], { alignment: AlignmentType.CENTER }),
                para([run("6.0 / 10", { size: 48, bold: true, color: AMBER })], { alignment: AlignmentType.CENTER }),
              ]
            }),
            new TableCell({
              borders: cellBorders(GREEN_MID),
              width: { size: 3120, type: WidthType.DXA },
              shading: { fill: GREEN_PALE, type: ShadingType.CLEAR },
              margins: { top: 120, bottom: 120, left: 160, right: 160 },
              children: [
                para([run("Screens Tested", { size: 18, color: GRAY })], { alignment: AlignmentType.CENTER }),
                para([run("6 / 6", { size: 48, bold: true, color: GREEN_DARK })], { alignment: AlignmentType.CENTER }),
              ]
            }),
          ]
        })
      ]
    }),
    spacer(400, 200),
    para([run("Date: June 15, 2026  ·  Product: feast4U  ·  Version: Hi-Fi Prototype v1.0", { size: 18, color: GRAY })], {
      alignment: AlignmentType.CENTER
    }),
  ];
}

// ── EXECUTIVE SUMMARY ─────────────────────────────────────
function execSummary() {
  const summaryRows = [
    ["✅ Strengths", "Strong visual identity with consistent green branding. Clean restaurant card layout. Real-time cart bill update. Order tracking on confirmation screen."],
    ["⚠️ Critical Gaps", "Missing payment screen before Order Confirmed (trust-breaking across all personas). Add+ button gives no persistent cart feedback. Pre-filled cart breaks prototype integrity."],
    ["🔧 Quick Wins", "Add persistent cart badge count. Make food category chips functional or remove them. Activate all bottom nav tabs. Add forgot-password link on Login."],
    ["🔴 Next Steps", "Prototype a payment selection screen (UPI/Card/COD). Fix all non-functional interactions. Run another round of testing with real users before development handoff."],
  ];

  const tableRows = summaryRows.map(([label, text]) =>
    new TableRow({
      children: [
        new TableCell({
          borders: cellBorders(BORDER_COL),
          width: { size: 2000, type: WidthType.DXA },
          shading: { fill: GREEN_PALE, type: ShadingType.CLEAR },
          margins: { top: 100, bottom: 100, left: 140, right: 140 },
          children: [para([run(label, { bold: true, size: 20, color: GREEN_DARK })])]
        }),
        new TableCell({
          borders: cellBorders(BORDER_COL),
          width: { size: 7360, type: WidthType.DXA },
          margins: { top: 100, bottom: 100, left: 140, right: 140 },
          children: [para([run(text, { size: 20, color: "374151" })])]
        }),
      ]
    })
  );

  return [
    para([run("Executive Summary", { size: 28, bold: true, color: GREEN_DARK })], {
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 280, after: 160 }
    }),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [2000, 7360],
      rows: tableRows
    }),
  ];
}

// ── TEST SCOPE ────────────────────────────────────────────
function testScope() {
  const tasks = [
    "Open the app and get past the splash screen",
    "Log in using email and password",
    "Browse the home screen and select a restaurant",
    "Add food items to cart from the restaurant menu",
    "Review cart, adjust quantities, and proceed to pay",
    "Read the order confirmation and delivery tracking screen",
  ];

  return [
    spacer(200, 0),
    para([run("Test Scope & Methodology", { size: 28, bold: true, color: GREEN_DARK })], {
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 280, after: 120 }
    }),
    para([run("Each of the 4 personas was simulated navigating the full 6-screen prototype flow, performing the following tasks:", { size: 20, color: "374151" })], {
      spacing: { before: 0, after: 100 }
    }),
    ...tasks.map((t, i) =>
      new Paragraph({
        numbering: { reference: "taskList", level: 0 },
        children: [run(t, { size: 20, color: "374151" })],
      })
    ),
    spacer(120, 0),
    para([run("Evaluation Dimensions", { size: 22, bold: true, color: GREEN_MID })], {
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 200, after: 100 }
    }),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [1800, 7560],
      rows: [
        ["First Impression", "Visual appeal, branding, and clarity of purpose on splash/login"],
        ["Navigation Ease", "Intuitiveness of screen flow, back buttons, and bottom nav"],
        ["Visual Clarity", "Readability, hierarchy, icon labelling, and contrast"],
        ["Task Completion", "Ability to complete all 6 tasks without confusion"],
        ["Trust & Safety", "Payment transparency, confirmation steps, data confidence"],
      ].map(([dim, desc]) => new TableRow({
        children: [
          new TableCell({
            borders: cellBorders(BORDER_COL),
            width: { size: 1800, type: WidthType.DXA },
            shading: { fill: GREEN_PALE, type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [para([run(dim, { bold: true, size: 20, color: GREEN_DARK })])]
          }),
          new TableCell({
            borders: cellBorders(BORDER_COL),
            width: { size: 7560, type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [para([run(desc, { size: 20, color: "374151" })])]
          }),
        ]
      }))
    }),
  ];
}

// ── PERSONA SECTION ───────────────────────────────────────
function personaSection(p, idx) {
  const scoreCol = scoreColor(p.score);
  const scoreBgCol = scoreBg(p.score);

  // Ratings table rows
  const ratingRows = Object.entries(p.scores).map(([dim, val]) => {
    const col = scoreColor(val);
    return new TableRow({
      children: [
        new TableCell({
          borders: cellBorders(BORDER_COL),
          width: { size: 2400, type: WidthType.DXA },
          margins: { top: 70, bottom: 70, left: 120, right: 120 },
          children: [para([run(dim, { size: 18, color: "374151" })])]
        }),
        new TableCell({
          borders: cellBorders(BORDER_COL),
          width: { size: 4560, type: WidthType.DXA },
          margins: { top: 70, bottom: 70, left: 120, right: 120 },
          children: [para([run(ratingBar(val), { size: 16, font: "Courier New", color: col })])]
        }),
        new TableCell({
          borders: cellBorders(BORDER_COL),
          width: { size: 800, type: WidthType.DXA },
          shading: { fill: scoreBg(val), type: ShadingType.CLEAR },
          margins: { top: 70, bottom: 70, left: 80, right: 80 },
          verticalAlign: VerticalAlign.CENTER,
          children: [para([run(`${val}/10`, { size: 18, bold: true, color: col })], { alignment: AlignmentType.CENTER })]
        }),
      ]
    });
  });

  const issueItems = p.issues.map(issue =>
    new Paragraph({ numbering: { reference: "bulletList", level: 0 }, children: [run(issue, { size: 20, color: "374151" })] })
  );
  const praiseItems = p.praises.map(praise =>
    new Paragraph({ numbering: { reference: "bulletList", level: 0 }, children: [run(praise, { size: 20, color: "374151" })] })
  );
  const suggItems = p.suggestions.map(s =>
    new Paragraph({ numbering: { reference: "bulletList", level: 0 }, children: [run(s, { size: 20, color: "374151" })] })
  );

  return [
    spacer(200, 0),
    // Header block
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [7560, 1800],
      rows: [
        new TableRow({
          children: [
            new TableCell({
              borders: cellBorders(GREEN_MID),
              width: { size: 7560, type: WidthType.DXA },
              shading: { fill: GREEN_DARK, type: ShadingType.CLEAR },
              margins: { top: 120, bottom: 120, left: 180, right: 180 },
              children: [
                para([
                  run(`Persona ${idx + 1}   `, { size: 18, color: "A5D6A7" }),
                  run(p.name, { size: 28, bold: true, color: WHITE }),
                ]),
                para([run(p.role, { size: 20, color: "74C69D" })]),
              ]
            }),
            new TableCell({
              borders: cellBorders(GREEN_MID),
              width: { size: 1800, type: WidthType.DXA },
              shading: { fill: scoreBgCol, type: ShadingType.CLEAR },
              margins: { top: 120, bottom: 120, left: 100, right: 100 },
              verticalAlign: VerticalAlign.CENTER,
              children: [
                para([run("Overall", { size: 18, color: GRAY })], { alignment: AlignmentType.CENTER }),
                para([run(`${p.score}/10`, { size: 44, bold: true, color: scoreCol })], { alignment: AlignmentType.CENTER }),
              ]
            }),
          ]
        })
      ]
    }),
    spacer(120, 0),

    // Ratings
    para([run("Dimension Ratings", { size: 22, bold: true, color: GREEN_MID })], { spacing: { before: 0, after: 100 } }),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [2400, 4560, 800],
      rows: [
        new TableRow({
          children: [
            new TableCell({ borders: cellBorders(GREEN_MID), width: { size: 2400, type: WidthType.DXA }, shading: { fill: GREEN_DARK, type: ShadingType.CLEAR }, margins: { top: 70, bottom: 70, left: 120, right: 120 }, children: [para([run("Dimension", { bold: true, size: 18, color: WHITE })])] }),
            new TableCell({ borders: cellBorders(GREEN_MID), width: { size: 4560, type: WidthType.DXA }, shading: { fill: GREEN_DARK, type: ShadingType.CLEAR }, margins: { top: 70, bottom: 70, left: 120, right: 120 }, children: [para([run("Rating", { bold: true, size: 18, color: WHITE })])] }),
            new TableCell({ borders: cellBorders(GREEN_MID), width: { size: 800, type: WidthType.DXA }, shading: { fill: GREEN_DARK, type: ShadingType.CLEAR }, margins: { top: 70, bottom: 70, left: 80, right: 80 }, children: [para([run("Score", { bold: true, size: 18, color: WHITE })], { alignment: AlignmentType.CENTER })] }),
          ]
        }),
        ...ratingRows
      ]
    }),
    spacer(140, 0),

    // Issues
    para([run("⚠  Usability Issues Found", { size: 22, bold: true, color: RED })], { spacing: { before: 0, after: 80 } }),
    ...issueItems,
    spacer(100, 0),

    // Praises
    para([run("✓  What Worked Well", { size: 22, bold: true, color: "15803D" })], { spacing: { before: 0, after: 80 } }),
    ...praiseItems,
    spacer(100, 0),

    // Suggestions
    para([run("💡  Suggestions for Improvement", { size: 22, bold: true, color: AMBER })], { spacing: { before: 0, after: 80 } }),
    ...suggItems,
    spacer(120, 0),

    // Quote
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [9360],
      rows: [
        new TableRow({
          children: [
            new TableCell({
              borders: { top: noBorder(), bottom: noBorder(), left: { style: BorderStyle.SINGLE, size: 12, color: GREEN_MID }, right: noBorder() },
              width: { size: 9360, type: WidthType.DXA },
              shading: { fill: GREEN_PALE, type: ShadingType.CLEAR },
              margins: { top: 100, bottom: 100, left: 200, right: 200 },
              children: [
                para([run(`"${p.quote}"`, { size: 20, italics: true, color: GREEN_DARK })]),
                para([run(`— ${p.name}, ${p.role}`, { size: 18, color: GRAY })], { spacing: { before: 60 } }),
              ]
            })
          ]
        })
      ]
    }),
    spacer(100, 0),

    // Summary
    para([run("Session Summary", { size: 20, bold: true, color: GREEN_DARK })], { spacing: { before: 0, after: 80 } }),
    para([run(p.summary, { size: 20, color: "374151" })], { spacing: { before: 0, after: 0 } }),
    hRule(),
  ];
}

// ── AGGREGATE INSIGHTS ────────────────────────────────────
function aggregateSection() {
  const topIssues = [
    ["Critical", "No payment screen before Order Confirmed — all 4 personas flagged this as a trust-breaking gap"],
    ["Critical", "Add+ button gives no persistent cart feedback — users don't know if items were added"],
    ["Major", "Pre-filled cart does not reflect what the user added from the menu — breaks prototype flow"],
    ["Major", "Food category chips appear tappable but have no function — false affordance"],
    ["Major", "Only one restaurant (Amar Punjabi) is interactive — Spice Garden is a dead end"],
    ["Moderate", "All bottom navigation tabs except Home are non-functional"],
    ["Moderate", "No forgot-password or show-password option on the Login screen"],
  ];

  const recommendations = [
    "Add a payment method selection screen (UPI / Card / Cash on Delivery) before Order Confirmed",
    "Implement a persistent cart badge showing item count on the nav and Restaurant screen toolbar",
    "Make the pre-filled cart consistent with the Add flow or clearly mark it as sample data",
    "Either make food chips functional category filters or remove them entirely",
    "Activate all bottom nav tabs with placeholder screens or remove non-functional ones",
    "Add forgot-password link and show/hide password toggle to the Login screen",
    "Run a WCAG contrast audit, particularly on the splash screen subtitle text",
  ];

  const issueRows = topIssues.map(([sev, text]) => {
    const col = sev === "Critical" ? RED : sev === "Major" ? AMBER : "6B7280";
    const bg  = sev === "Critical" ? RED_LIGHT : sev === "Major" ? "FEF9C3" : "F3F4F6";
    return new TableRow({
      children: [
        new TableCell({
          borders: cellBorders(BORDER_COL),
          width: { size: 1400, type: WidthType.DXA },
          shading: { fill: bg, type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [para([run(sev, { bold: true, size: 18, color: col })], { alignment: AlignmentType.CENTER })]
        }),
        new TableCell({
          borders: cellBorders(BORDER_COL),
          width: { size: 7960, type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [para([run(text, { size: 20, color: "374151" })])]
        }),
      ]
    });
  });

  return [
    spacer(200, 0),
    para([run("Aggregate Insights — All Personas", { size: 28, bold: true, color: GREEN_DARK })], {
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 280, after: 160 }
    }),

    // Score summary table
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [2340, 2340, 2340, 2340],
      rows: [
        new TableRow({
          children: personas.map(p =>
            new TableCell({
              borders: cellBorders(GREEN_MID),
              width: { size: 2340, type: WidthType.DXA },
              shading: { fill: GREEN_DARK, type: ShadingType.CLEAR },
              margins: { top: 100, bottom: 100, left: 100, right: 100 },
              children: [para([run(p.name, { bold: true, size: 18, color: WHITE })], { alignment: AlignmentType.CENTER })]
            })
          )
        }),
        new TableRow({
          children: personas.map(p =>
            new TableCell({
              borders: cellBorders(GREEN_MID),
              width: { size: 2340, type: WidthType.DXA },
              shading: { fill: scoreBg(p.score), type: ShadingType.CLEAR },
              margins: { top: 100, bottom: 100, left: 100, right: 100 },
              children: [para([run(`${p.score}/10`, { bold: true, size: 36, color: scoreColor(p.score) })], { alignment: AlignmentType.CENTER })]
            })
          )
        }),
        new TableRow({
          children: personas.map(p =>
            new TableCell({
              borders: cellBorders(BORDER_COL),
              width: { size: 2340, type: WidthType.DXA },
              margins: { top: 80, bottom: 80, left: 100, right: 100 },
              children: [para([run(p.role.split("·")[0].trim(), { size: 16, color: GRAY })], { alignment: AlignmentType.CENTER })]
            })
          )
        }),
      ]
    }),

    spacer(200, 0),
    para([run("Top Issues by Severity", { size: 22, bold: true, color: RED })], { spacing: { before: 0, after: 100 } }),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [1400, 7960],
      rows: [
        new TableRow({
          children: [
            new TableCell({ borders: cellBorders(GREEN_MID), width: { size: 1400, type: WidthType.DXA }, shading: { fill: GREEN_DARK, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [para([run("Severity", { bold: true, size: 18, color: WHITE })], { alignment: AlignmentType.CENTER })] }),
            new TableCell({ borders: cellBorders(GREEN_MID), width: { size: 7960, type: WidthType.DXA }, shading: { fill: GREEN_DARK, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [para([run("Issue", { bold: true, size: 18, color: WHITE })])] }),
          ]
        }),
        ...issueRows
      ]
    }),

    spacer(200, 0),
    para([run("Priority Recommendations", { size: 22, bold: true, color: GREEN_DARK })], { spacing: { before: 0, after: 100 } }),
    ...recommendations.map((r, i) =>
      new Paragraph({
        numbering: { reference: "recList", level: 0 },
        children: [run(r, { size: 20, color: "374151" })],
      })
    ),
  ];
}

// ── BUILD DOCUMENT ────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bulletList",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 }, spacing: { before: 40, after: 40 } } }
        }]
      },
      {
        reference: "taskList",
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 }, spacing: { before: 40, after: 40 } } }
        }]
      },
      {
        reference: "recList",
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 }, spacing: { before: 60, after: 60 } } }
        }]
      },
    ]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 20 } } },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: GREEN_DARK },
        paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 0 }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Arial", color: GREEN_MID },
        paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 1 }
      },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
      }
    },
    children: [
      ...titlePage(),
      new Paragraph({ children: [new PageBreak()] }),
      ...execSummary(),
      new Paragraph({ children: [new PageBreak()] }),
      ...testScope(),
      new Paragraph({ children: [new PageBreak()] }),
      // Personas
      ...personas.flatMap((p, i) => personaSection(p, i)),
      new Paragraph({ children: [new PageBreak()] }),
      ...aggregateSection(),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/home/claude/feast4U_user_testing_report.docx', buf);
  console.log('Done');
});
