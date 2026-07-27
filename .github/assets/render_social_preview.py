"""Render the docx_plus GitHub social preview card (1280x640).

Regenerate with::

    uv run --with cairosvg python .github/assets/render_social_preview.py

Writes ``social-preview.svg`` (editable source) and ``social-preview.png``
(the artifact to upload under repo Settings -> Social preview) next to this
script. GitHub wants 1280x640 and under 1MB.

cairosvg is deliberately not a project dev dependency -- this runs by hand,
rarely. The card uses Segoe UI and Cascadia Mono, falling back to Arial and
Consolas; on a box without them the layout holds but the metrics shift.
"""

import pathlib

import cairosvg

OUT = pathlib.Path(__file__).parent
W, H = 1280, 640

BG, BG2 = "#0d1226", "#141b3a"
INK, MUTED = "#ffffff", "#94a3b8"
ACCENT, ACCENT2 = "#818cf8", "#38bdf8"
CHIP_BG, CHIP_INK = "#1e2547", "#c7d2fe"
GREEN, AMBER, PINK = "#4ade80", "#fbbf24", "#f472b6"

SANS = "Segoe UI, Arial, sans-serif"
MONO = "Cascadia Mono, Consolas, monospace"


def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


# Each code line is ONE <text> with <tspan> children, so font metrics -- not
# my arithmetic -- advance the pen. (text-anchor + tspan is what broke concept B;
# there is no anchor here, so tspans lay out correctly.)
CODE = [
    [("# python-docx stops at the surface.", MUTED)],
    [],
    [("fmt = ", INK), ("resolve_effective_formatting", ACCENT2), ("(", INK)],
    [("    p, include_provenance=", INK), ("True", PINK), (")", INK)],
    [],
    [("fmt.font_size", INK), ("            # 13.0", MUTED)],
    [("fmt.provenance[", INK), ("'font_size'", GREEN), ("]", INK)],
    [("# layer='paragraphStyle'", MUTED)],
    [("# style_id='Heading2'", MUTED)],
]


def code_block(x, y, size=21, leading=32):
    out = []
    for i, spans in enumerate(CODE):
        if not spans:
            continue
        parts = "".join(f'<tspan fill="{c}">{esc(t)}</tspan>' for t, c in spans)
        out.append(
            f'<text x="{x}" y="{y + i * leading}" font-family="{MONO}" '
            f'font-size="{size}" xml:space="preserve">{parts}</text>'
        )
    return "\n".join(out)


def chips(items, x, y, gap=12, pad=18, fs=20, char_w=9.5):
    out, cx = [], x
    for label in items:
        w = len(label) * char_w + pad * 2
        out.append(
            f'<rect x="{cx:.0f}" y="{y}" width="{w:.0f}" height="40" rx="20" '
            f'fill="{CHIP_BG}" stroke="{ACCENT}" stroke-opacity="0.25"/>'
            f'<text x="{cx + w / 2:.0f}" y="{y + 27}" font-family="{SANS}" '
            f'font-size="{fs}" fill="{CHIP_INK}" text-anchor="middle">{label}</text>'
        )
        cx += w + gap
    return "\n".join(out)


svg = f"""<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" viewBox="0 0 {W} {H}">
<defs>
  <linearGradient id="bg" x1="0" y1="0" x2="1" y2="1">
    <stop offset="0%" stop-color="{BG}"/><stop offset="100%" stop-color="{BG2}"/>
  </linearGradient>
  <linearGradient id="rule" x1="0" y1="0" x2="1" y2="0">
    <stop offset="0%" stop-color="{ACCENT}"/>
    <stop offset="100%" stop-color="{ACCENT2}" stop-opacity="0.1"/>
  </linearGradient>
  <radialGradient id="glow" cx="50%" cy="50%" r="50%">
    <stop offset="0%" stop-color="{ACCENT}" stop-opacity="0.20"/>
    <stop offset="55%" stop-color="{ACCENT}" stop-opacity="0.07"/>
    <stop offset="100%" stop-color="{ACCENT}" stop-opacity="0"/>
  </radialGradient>
</defs>
<rect width="{W}" height="{H}" fill="url(#bg)"/>
<circle cx="1180" cy="90" r="340" fill="url(#glow)"/>

<!-- left column -->
<text x="72" y="196" font-family="{SANS}" font-size="80" font-weight="700"
      fill="{INK}">docx<tspan fill="{ACCENT}">_plus</tspan></text>
<rect x="74" y="222" width="150" height="5" rx="2.5" fill="url(#rule)"/>

<text x="72" y="288" font-family="{SANS}" font-size="30" fill="{MUTED}">OOXML-level extensions for</text>
<text x="72" y="328" font-family="{SANS}" font-size="30" fill="{INK}">python-docx</text>

<text x="72" y="390" font-family="{SANS}" font-size="20" fill="{MUTED}">The style cascade, content controls, comments,</text>
<text x="72" y="418" font-family="{SANS}" font-size="20" fill="{MUTED}">and tracked changes python-docx can&#8217;t reach.</text>

{chips(["cascade", "controls", "comments", "revisions"], 72, 452)}

<text x="72" y="580" font-family="{MONO}" font-size="22" fill="{ACCENT2}">pip install docx-plus</text>
<text x="1208" y="580" font-family="{SANS}" font-size="20" fill="{MUTED}"
      text-anchor="end">MIT &#183; Python 3.10+</text>

<!-- code panel -->
<rect x="656" y="150" width="576" height="366" rx="16" fill="#080d1e" fill-opacity="0.8"
      stroke="{ACCENT}" stroke-opacity="0.22"/>
<circle cx="686" cy="180" r="6" fill="#ef4444" fill-opacity="0.7"/>
<circle cx="708" cy="180" r="6" fill="{AMBER}" fill-opacity="0.7"/>
<circle cx="730" cy="180" r="6" fill="{GREEN}" fill-opacity="0.7"/>
<line x1="656" y1="206" x2="1232" y2="206" stroke="{ACCENT}" stroke-opacity="0.14"/>
{code_block(686, 248)}
</svg>"""

(OUT / "social-preview.svg").write_text(svg, encoding="utf-8")
cairosvg.svg2png(
    bytestring=svg.encode("utf-8"),
    write_to=str(OUT / "social-preview.png"),
    output_width=W,
    output_height=H,
)
print("wrote social-preview.svg + .png", (OUT / "social-preview.png").stat().st_size, "bytes")
