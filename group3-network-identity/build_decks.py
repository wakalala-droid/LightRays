"""Builds the Group 3 DNS decks.

  python3 build_decks.py
    -> Group3_Network_Identity_DNS_Redesign.pptx  (styled like our APE deck)
    -> Group3_Network_Identity_DNS.pptx           (first version, now with cover + sponsor logos)
"""
import os
from PIL import Image
from pptx import Presentation
from pptx.util import Emu, Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

HERE = os.path.dirname(os.path.abspath(__file__))
A = lambda f: os.path.join(HERE, "assets", f)

# Layout works on a 2000 x 1125 px canvas, same as the APE deck.
PX = 6096  # EMU per px on a 13.333in wide slide
E = lambda v: Emu(int(v * PX))
FS = lambda v: Pt(v * 0.48)  # px font size -> pt

FONT_HEAD = "Segoe UI Black"
FONT = "Segoe UI"
FONT_SEMI = "Segoe UI Semibold"

def rgb(h): return RGBColor.from_string(h.lstrip("#"))

NAVY, NAVY2 = "#0F1D33", "#12284A"
INK, BODY, MUTED = "#0F172A", "#334155", "#64748B"
BG, CARD, LINE = "#F4F6FA", "#FBFCFE", "#DDE3EC"
BLUE, TEAL, ORANGE, PURPLE, RED = "#1D4ED8", "#0F766E", "#B45309", "#7C3AED", "#BE123C"
TINT = {BLUE: "#E3EAFB", TEAL: "#DDF0EE", ORANGE: "#FCE9D9", PURPLE: "#ECE5FD", RED: "#FBE1E7"}
CODE_TINT = {BLUE: "#E3EAFB", TEAL: "#0B5C55", ORANGE: "#FCE9D9", PURPLE: "#ECE5FD", RED: "#FBE1E7"}

MEMBERS = [("Wakalala Mulyokela", "Leader", BLUE), ("Ruth Chipompo", "", TEAL),
           ("Hampande Luyando", "", ORANGE), ("Loveness Chanda", "", PURPLE),
           ("Dumisani Zulu", "", RED)]
SPEAKER = {m[0]: (m[0] + (" · " + m[1] if m[1] else ""), m[2]) for m in MEMBERS}

LOGOS = [("logo_smartafrica.png", 258, 72), ("logo_mastercard.png", 525, 70),
         ("logo_growthafrica.png", 705, 52), ("logo_pigier.png", 1015, 72),
         ("logo_netone.png", 1232, 60), ("logo_kectil.png", 1545, 72)]


# ---------- primitives ----------
def rect(s, x, y, w, h, fill, line=None, radius=0, dash=False, shape=None):
    kind = shape or (MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE)
    r = s.shapes.add_shape(kind, E(x), E(y), E(w), E(h))
    if radius and kind == MSO_SHAPE.ROUNDED_RECTANGLE:
        r.adjustments[0] = min(0.5, radius / min(w, h))
    if fill: r.fill.solid(); r.fill.fore_color.rgb = rgb(fill)
    else: r.fill.background()
    if line:
        r.line.color.rgb = rgb(line); r.line.width = Pt(1.25)
        if dash: r.line.dash_style = 4
    else: r.line.fill.background()
    r.shadow.inherit = False
    return r

def text(s, x, y, w, h, paras, size=36, color=BODY, font=FONT, bold=False,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=1.12, after=0, box=None):
    """paras: str or list of paragraphs; a paragraph is a str or list of (text, opts) runs."""
    shp = box or s.shapes.add_textbox(E(x), E(y), E(w), E(h))
    tf = shp.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    if isinstance(paras, str): paras = [paras]
    for i, p in enumerate(paras):
        q = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        q.alignment = align; q.line_spacing = spacing; q.space_after = FS(after)
        runs = [(p, {})] if isinstance(p, str) else p
        for t, o in runs:
            r = q.add_run(); r.text = t
            r.font.size = FS(o.get("size", size)); r.font.name = o.get("font", font)
            r.font.bold = o.get("bold", bold or r.font.name in (FONT_HEAD, FONT_SEMI)); r.font.color.rgb = rgb(o.get("color", color))
    return shp

def pill(s, x, y, label, color, size=26, h=56, pad=24, fg="#FFFFFF", right=False):
    w = int(len(label) * size * 0.55) + pad * 2
    if right: x = x - w
    r = rect(s, x, y, w, h, color, radius=h / 2)
    text(s, 0, 0, 0, 0, label, size=size, color=fg, font=FONT_SEMI,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, box=r)
    return x + w

def picture(s, path, x, y, h=None, w=None):
    im = Image.open(path); ar = im.width / im.height
    if h and not w: w = h * ar
    if w and not h: h = w / ar
    return s.shapes.add_picture(path, E(x), E(y), E(w), E(h))

def footer(s, num, inch_scale=False):
    rect(s, 0, 1017, 2000, 98, "#FFFFFF")
    text(s, 30, 1048, 220, 30, "PROGRAMME SPONSORS", size=17, color="#374151", font=FONT_SEMI,
         anchor=MSO_ANCHOR.MIDDLE)
    for f, x, h in LOGOS:
        picture(s, A(f), x, 1066 - h / 2, h=h)
    if num is not None:
        text(s, 1880, 1040, 80, 44, f"{num:02d}", size=26, color="#374151", font=FONT_SEMI,
             align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    rect(s, 0, 1115, 2000, 10, NAVY)

def cover_slide(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(A("cover.png"), 0, 0, prs.slide_width, prs.slide_height)
    return s


# ---------- redesigned deck ----------
class Deck:
    def __init__(self):
        self.p = Presentation(); self.p.slide_width = E(2000); self.p.slide_height = E(1125)
        self.n = 0

    def slide(self, dark=False):
        s = self.p.slides.add_slide(self.p.slide_layouts[6]); self.n += 1
        rect(s, 0, 0, 2000, 1125, NAVY if dark else BG)
        if dark:
            g = rect(s, 0, 0, 2000, 1017, None)
            g.fill.gradient(); g.fill.gradient_angle = 135
            st = g.fill.gradient_stops; st[0].color.rgb = rgb(NAVY); st[1].color.rgb = rgb(NAVY2)
            for i, c in enumerate([BLUE, TEAL, ORANGE, PURPLE, RED]):
                rect(s, i * 400, 0, 400, 16, c)
        footer(s, self.n + 1)  # slide 1 is the programme cover
        return s

    def header(self, s, kicker, title, speaker, color, dark=False):
        text(s, 134, 110, 1200, 40, kicker.upper(), size=26, color="#93C5FD" if dark else color,
             font=FONT_SEMI)
        text(s, 134, 165, 1700, 100, title, size=66, color="#F8FAFC" if dark else INK, font=FONT_HEAD)
        if speaker:
            label, c = SPEAKER[speaker]
            pill(s, 1866, 100, "Speaker: " + label, c, right=True)

    def card(self, s, x, y, w, h, head, body, color=None, tint=False, head_color=INK, size=34,
             head_size=42, code=None, num=None):
        rect(s, x, y, w, h, TINT[color] if tint else CARD, line=color and tint and TINT[color] or LINE,
             radius=22)
        hx = x + 32
        if num is not None:
            c = rect(s, x + 30, y + 26, 58, 58, color, shape=MSO_SHAPE.OVAL)
            text(s, 0, 0, 0, 0, str(num), size=34, color="#FFFFFF", font=FONT_HEAD,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, box=c)
            hx = x + 106
        text(s, hx, y + 28, w - (hx - x) - 24, 60, head, size=head_size, color=head_color, font=FONT_HEAD)
        by = y + 98
        if code:
            r = rect(s, x + 32, by, w - 64, 60, CODE_TINT[color], radius=10)
            text(s, 0, 0, 0, 0, code, size=30, color=color, font="Consolas", bold=True,
                 anchor=MSO_ANCHOR.MIDDLE, box=r).text_frame.margin_left = E(16)
            by += 82
        text(s, x + 32, by, w - 64, h - (by - y) - 20, body, size=size, color=BODY, after=8)

    def shot(self, s, x, y, w, h, label, color):
        rect(s, x, y, w, h, "#FFFFFF", line=color, radius=18, dash=True)
        text(s, x + 30, y, w - 60, h, [[("[ Paste screenshot here ]", {"font": FONT_SEMI, "color": color})],
                                      label], size=28, color=MUTED, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, after=6)

    def save(self, path): self.p.save(path)


def build_redesign(path):
    d = Deck()
    cover_slide(d.p)

    # Title
    s = d.slide(dark=True)
    text(s, 134, 115, 1700, 40, "SAMA NALA PROGRAM · TASKS 2.1 & 2.2 · NETWORK FUNDAMENTALS",
         size=26, color="#93C5FD", font=FONT_SEMI)
    text(s, 134, 318, 1500, 50, "Group 3 · Tasks 2.1 & 2.2", size=38, color="#CBD5E1")
    text(s, 134, 395, 1750, 220, ["Your Network Identity;", "DNS in Action"], size=96,
         color="#F8FAFC", font=FONT_HEAD, spacing=1.0)
    text(s, 134, 640, 1700, 60, "Who are we on the network, and how do names become IP addresses?",
         size=38, color="#E2E8F0")
    x = 134
    for name, role, c in MEMBERS:
        x = pill(s, x, 855, name + (" · " + role if role else ""), c, size=28) + 18

    # Who presents what
    s = d.slide()
    d.header(s, "Our plan · about 10 minutes", "Who presents what", "Wakalala Mulyokela", BLUE)
    rows = [("Leader · opens and closes", "Introduction · Summary and questions", "2 min"),
            ("Task 2.1", "What is our IP address?", "2 min"),
            ("Task 2.1", "DNS, DHCP and the default gateway", "2 min"),
            ("Task 2.2", "How DNS converts names into IP addresses · DNS in action", "2 min"),
            ("Task 2.2", "Why DNS matters for cybersecurity investigations", "2 min")]
    for i, ((name, _, c), (sub, topic, t)) in enumerate(zip(MEMBERS, rows)):
        y = 292 + i * 137
        rect(s, 134, y, 1732, 120, CARD, line=LINE, radius=20)
        text(s, 168, y + 18, 460, 50, name, size=38, color=INK, font=FONT_HEAD)
        text(s, 168, y + 70, 460, 40, sub, size=26, color=c, font=FONT_SEMI)
        text(s, 650, y, 1040, 120, topic, size=36, color=BODY, anchor=MSO_ANCHOR.MIDDLE)
        r = rect(s, 1726, y + 34, 104, 52, TINT[c], radius=26)
        text(s, 0, 0, 0, 0, t, size=26, color=c, font=FONT_SEMI, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, box=r)

    # Q1 IP address
    s = d.slide()
    d.header(s, "Task 2.1 · Question 1 · IP address", "What is our IP address?", "Ruth Chipompo", TEAL)
    rect(s, 134, 292, 760, 670, TEAL, radius=24)
    text(s, 180, 335, 680, 40, "OUR LAPTOP'S NETWORK IDENTITY", size=26, color="#CCFBF1", font=FONT_SEMI)
    for i, (k, hint) in enumerate([("Private IPv4", "from ipconfig"), ("Subnet mask", "from ipconfig"),
                                   ("Default gateway", "from ipconfig"), ("Public IP", "from a “what is my IP” site")]):
        y = 400 + i * 132
        text(s, 180, y, 680, 40, [[(k, {"font": FONT_HEAD, "color": "#FFFFFF"}),
                                   ("   " + hint, {"size": 24, "color": "#99F6E4"})]], size=34)
        r = rect(s, 180, y + 50, 668, 62, "#0B5C55", radius=10)
        text(s, 0, 0, 0, 0, "____ . ____ . ____ . ____", size=32, color="#FFFFFF", font="Consolas",
             bold=True, anchor=MSO_ANCHOR.MIDDLE, box=r).text_frame.margin_left = E(20)
    d.card(s, 924, 292, 942, 205, "What an IP address is",
           "The number that identifies a device on a network, like a postal address. "
           "IPv4: 192.168.1.105 · IPv6: fe80::1c2a:3bff…", size=30)
    d.card(s, 924, 523, 942, 205, "Private vs public",
           "Private (10.x, 172.16–31.x, 192.168.x) works inside our network. The router's "
           "public IP is shared by all of us through NAT.", size=30)
    d.shot(s, 924, 754, 942, 208, "ipconfig  ·  blur the public IP before sharing", TEAL)

    # Q2 DHCP / DNS / gateway
    s = d.slide()
    d.header(s, "Task 2.1 · Question 2 · DNS, DHCP and gateway", "The three settings that connect us",
             "Hampande Luyando", ORANGE)
    cards = [("DHCP", "Automatically gives the laptop its IP, mask, gateway and DNS servers, for a lease. "
                      "DORA: Discover → Offer → Request → Acknowledge.", "DHCP Enabled: Yes"),
             ("Default gateway", "The router. Traffic for any address outside our network leaves "
                                 "through it: our exit to the Internet.", "Default Gateway: ________"),
             ("DNS", "The Internet's phone book. Turns names like google.com into the IP "
                     "addresses computers use.", "DNS Servers: ________")]
    for i, (h, b, c) in enumerate(cards):
        d.card(s, 134 + i * 587, 292, 558, 410, h, b, color=ORANGE, tint=(i == 0), code=c, size=30)
    d.shot(s, 134, 728, 850, 234, "ipconfig /all  ·  highlight DHCP, Gateway and DNS lines", ORANGE)
    d.shot(s, 1016, 728, 850, 234, "tracert google.com  ·  first hop = default gateway", ORANGE)

    # Q3 How DNS works
    s = d.slide()
    d.header(s, "Task 2.2 · Question 3 · Name resolution", "How DNS turns a name into an IP address",
             "Loveness Chanda", PURPLE)
    flow = [("Our laptop", "Checks its cache and hosts file"), ("Resolver", "Our DNS server, asked on UDP port 53"),
            ("Root server", "“Ask the .com servers”"), ("TLD (.com)", "“Ask google.com's name servers”"),
            ("Authoritative", "Returns the A record: the IP")]
    for i, (h, b) in enumerate(flow):
        x = 134 + i * 352
        hl = i == 4
        rect(s, x, 292, 300, 220, PURPLE if hl else CARD, line=None if hl else LINE, radius=22)
        text(s, x + 26, 318, 260, 50, h, size=38, color="#FFFFFF" if hl else INK, font=FONT_HEAD)
        text(s, x + 26, 378, 255, 130, b, size=29, color="#EDE9FE" if hl else BODY)
        if i < 4:
            rect(s, x + 308, 384, 36, 36, "#94A3B8", shape=MSO_SHAPE.RIGHT_ARROW)
    d.card(s, 134, 540, 850, 422, "Then what?",
           ["• The resolver sends the IP back to the laptop", "• Both cache it for its TTL (time to live)",
            "• The browser connects to that IP",
            "• “Non-authoritative answer” in nslookup = came from a cache"], size=30)
    d.card(s, 1016, 540, 850, 422, "Common record types",
           [[("A ", {"bold": True, "color": PURPLE}), ("IPv4 address", {})],
            [("AAAA ", {"bold": True, "color": PURPLE}), ("IPv6 address", {})],
            [("CNAME ", {"bold": True, "color": PURPLE}), ("alias to another name", {})],
            [("MX ", {"bold": True, "color": PURPLE}), ("mail server", {})],
            [("NS · TXT ", {"bold": True, "color": PURPLE}), ("name servers · text / SPF", {})]],
           color=PURPLE, tint=True, size=30)

    # DNS in action
    s = d.slide()
    d.header(s, "Task 2.2 · DNS in action · on our laptop", "Watching DNS work", "Loveness Chanda", PURPLE)
    cmds = [("nslookup google.com", "Name → IP: ____________"),
            ("ipconfig /displaydns", "Names our laptop already cached"),
            ("ipconfig /flushdns", "Clear the cache, then look up again"),
            ("ping google.com", "Resolve the name, then reach the IP"),
            ("nslookup -type=MX gmail.com", "Find Gmail's mail servers")]
    for i, (c, b) in enumerate(cmds):
        y = 292 + i * 136
        circ = rect(s, 134, y + 20, 66, 66, PURPLE, shape=MSO_SHAPE.OVAL)
        text(s, 0, 0, 0, 0, str(i + 1), size=34, color="#FFFFFF", font=FONT_HEAD,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, box=circ)
        r = rect(s, 228, y + 14, 560, 58, TINT[PURPLE], radius=10)
        text(s, 0, 0, 0, 0, c, size=29, color=PURPLE, font="Consolas", bold=True,
             anchor=MSO_ANCHOR.MIDDLE, box=r).text_frame.margin_left = E(16)
        text(s, 230, y + 80, 600, 40, b, size=27, color=BODY)
    d.shot(s, 870, 292, 996, 330, "nslookup google.com", PURPLE)
    d.shot(s, 870, 646, 486, 316, "ipconfig /displaydns or /flushdns", PURPLE)
    d.shot(s, 1380, 646, 486, 316, "ping or MX lookup", PURPLE)

    # Cybersecurity
    s = d.slide()
    d.header(s, "Task 2.2 · Question 4 · Cybersecurity", "Why DNS matters in investigations",
             "Dumisani Zulu", RED)
    sec = [("Activity timeline", "Nearly every connection starts with a lookup, so DNS logs show what a device contacted and when."),
           ("Malware and C2", "Infected machines call attacker domains: random-looking (DGA) names, many NXDOMAIN errors."),
           ("Phishing", "Spot look-alike and newly registered domains, like g00gle-login.com."),
           ("DNS tunnelling", "Stolen data hidden in long or frequent queries and TXT records."),
           ("Attacks on DNS", "Spoofing, cache poisoning, hijacking and DDoS. DNSSEC checks answers are genuine."),
           ("Defence and evidence", "Filtering and sinkholing block bad domains. Check the DNS cache and hosts file.")]
    for i, (h, b) in enumerate(sec):
        x = 134 + (i % 3) * 587; y = 292 + (i // 3) * 318
        d.card(s, x, y, 558, 292, h, b, color=RED, tint=(i == 5), size=30, head_size=40)
    text(s, 134, 938, 1700, 40, "Attribution too: WHOIS and passive DNS history link domains and IPs to the same attacker.",
         size=27, color=MUTED)

    # Summary
    s = d.slide(dark=True)
    d.header(s, "Summary", "Answering our questions", "Wakalala Mulyokela", BLUE, dark=True)
    r = rect(s, 134, 330, 1732, 330, "#1A3558", radius=24)
    text(s, 180, 360, 1640, 270,
         [[("IP address: ", {"bold": True, "color": "#FFFFFF"}), ("our identity on the network, private inside and public outside.", {})],
          [("DHCP ", {"bold": True, "color": "#FFFFFF"}), ("configures us, the ", {}), ("gateway ", {"bold": True, "color": "#FFFFFF"}),
           ("gets us out, ", {}), ("DNS ", {"bold": True, "color": "#FFFFFF"}), ("finds names.", {})],
          [("Resolution: ", {"bold": True, "color": "#FFFFFF"}), ("cache → resolver → root → TLD → authoritative.", {})],
          [("Security: ", {"bold": True, "color": "#FFFFFF"}), ("DNS logs are key evidence for spotting malware, phishing and tunnelling.", {})]],
         size=36, color="#E2E8F0", anchor=MSO_ANCHOR.MIDDLE, after=10)
    text(s, 134, 790, 1700, 100, "Thank you. Questions?", size=80, color="#F8FAFC", font=FONT_HEAD)
    x = 134
    for name, role, c in MEMBERS:
        x = pill(s, x, 905, name + (" · " + role if role else ""), c, size=28) + 18

    # Key terms
    s = d.slide()
    text(s, 134, 110, 1200, 40, "REFERENCE · FOR QUESTIONS AND REVISION", size=26, color=INK, font=FONT_SEMI)
    text(s, 134, 165, 1700, 100, "Key terms in simple words", size=66, color=INK, font=FONT_HEAD)
    pill(s, 1866, 100, "Backup slide · any member", INK, right=True)
    terms = [("IP address", "A device's address on a network", BLUE), ("NAT", "Many private IPs sharing one public IP", BLUE),
             ("DHCP", "Hands out IP settings automatically", TEAL), ("Gateway", "The router to other networks", ORANGE),
             ("DNS", "Turns names into IP addresses", ORANGE), ("Resolver", "The DNS server that looks names up for us", PURPLE),
             ("TTL", "How long an answer may be cached", PURPLE), ("NXDOMAIN", "“This name does not exist” reply", RED),
             ("DNSSEC", "Signatures proving DNS answers are real", RED)]
    for i, (h, b, c) in enumerate(terms):
        x = 134 + (i % 3) * 587; y = 292 + (i // 3) * 230
        d.card(s, x, y, 558, 204, h, b, head_color=c, size=32, head_size=40)

    d.save(path)


# ---------- first deck, now with cover and sponsor logos ----------
def build_v1(path):
    TEALv, NAVYv, GREY = rgb("#0E6E66"), rgb("#14243D"), rgb("#888888")
    p = Presentation(); p.slide_width = E(2000); p.slide_height = E(1125)
    B = p.slide_layouts[6]
    cover_slide(p)
    BOTTOM = Inches(6.55)  # content stops above the sponsor footer

    def bar(s, title):
        r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, p.slide_width, Inches(1.1))
        r.fill.solid(); r.fill.fore_color.rgb = TEALv; r.line.fill.background()
        tf = r.text_frame; tf.text = title; q = tf.paragraphs[0]
        q.font.size = Pt(30); q.font.bold = True; q.font.color.rgb = RGBColor(255, 255, 255)

    def bullets(s, items, x, y, w, h, size=18):
        tf = s.shapes.add_textbox(x, y, w, h).text_frame; tf.word_wrap = True
        for i, t in enumerate(items):
            q = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            sub = t.startswith("- ")
            q.text = ("    – " if sub else "• ") + t.lstrip("- ")
            q.font.size = Pt(size - 3 if sub else size); q.font.color.rgb = NAVYv; q.space_after = Pt(5)

    def shot(s, label, x, y, w, h):
        r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h); r.fill.solid()
        r.fill.fore_color.rgb = RGBColor(0xEE, 0xF4, 0xF3); r.line.color.rgb = TEALv; r.line.dash_style = 4
        tf = r.text_frame; tf.word_wrap = True; tf.text = "[ Paste screenshot here ]\n" + label
        for q in tf.paragraphs: q.font.size = Pt(14); q.font.color.rgb = GREY

    def slide(title, items, shots=(), size=18):
        s = p.slides.add_slide(B); bar(s, title)
        top = Inches(1.3); avail = BOTTOM - top
        if not shots:
            bullets(s, items, Inches(0.6), top, Inches(12.1), avail, size)
        else:
            bullets(s, items, Inches(0.5), top, Inches(6.0), avail, size)
            h = avail / len(shots)
            for i, l in enumerate(shots):
                shot(s, l, Inches(6.8), top + int(h * i), Inches(6.0), int(h - Inches(0.12)))
        footer(s, len(p.slides))
        return s

    s = p.slides.add_slide(B)
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, p.slide_width, p.slide_height)
    r.fill.solid(); r.fill.fore_color.rgb = TEALv; r.line.fill.background()
    tf = s.shapes.add_textbox(Inches(0.8), Inches(1.4), Inches(11.7), Inches(4.5)).text_frame; tf.word_wrap = True
    for i, (t, sz, b) in enumerate([("Your Network Identity & DNS in Action", 40, True),
                                     ("Group 3 — Tasks 2.1 & 2.2", 24, False), ("", 12, False),
                                     ("Wakalala Mulyokela (Leader) · Ruth Chipompo · Hampande Luyando · "
                                      "Loveness Chanda · Dumisani Zulu", 18, False)]):
        q = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        q.text = t; q.font.size = Pt(sz); q.font.bold = b; q.font.color.rgb = RGBColor(255, 255, 255)
    footer(s, 2)

    slide("1. What is our IP address?", [
        "IP address = a device's address on a network (like a postal address)",
        "IPv4: 4 numbers, e.g. 192.168.1.105 · IPv6: e.g. fe80::1c2a:…",
        "Private IP (inside our network): ____________", "- found with: ipconfig",
        "Subnet mask: ____________", "Public IP (what the Internet sees): ____________",
        "- many devices share one public IP via NAT",
        "Private ranges: 10.x.x.x · 172.16–31.x.x · 192.168.x.x"],
        ["ipconfig (01)", "Public IP website (03) — blur before sharing"])
    slide("2. DHCP, DNS & the Default Gateway", [
        "DHCP: automatically gives the laptop its IP, mask, gateway & DNS",
        "- DORA: Discover → Offer → Request → Acknowledge (leased for a time)",
        "- Evidence: \"DHCP Enabled: Yes\", \"DHCP Server\"",
        "Default gateway: the router — our exit to other networks/Internet",
        "- Ours: ____________ (first hop in tracert)",
        "DNS: the Internet's phone book — names → IP addresses", "- Our DNS servers: ____________"],
        ["ipconfig /all (02) — highlight DHCP, Gateway, DNS lines", "tracert google.com (09)"])
    slide("3. How DNS converts names into IP addresses", [
        "1. Check local cache & hosts file (ipconfig /displaydns)",
        "2. Ask recursive resolver (from DHCP) over UDP port 53",
        "3. Root server → \"ask the .com servers\"", "4. TLD (.com) server → \"ask google.com's name servers\"",
        "5. Authoritative server → A record = IP address", "6. Answer cached for its TTL; browser connects to the IP",
        "Records: A (IPv4) · AAAA (IPv6) · CNAME · MX (mail) · NS · TXT"],
        ["Diagram: Laptop → Resolver → Root → TLD → Authoritative"], size=17)
    slide("4. DNS in Action — our lookups", [
        "nslookup google.com → returned: ____________",
        "- \"Non-authoritative answer\" = came from a resolver's cache",
        "ipconfig /displaydns → names already cached", "ipconfig /flushdns → cache cleared, fresh lookup",
        "ping google.com → name resolved, then IP contacted", "nslookup -type=MX gmail.com → mail servers"],
        ["nslookup google.com (04)", "ipconfig /displaydns or /flushdns (06/07)", "ping / MX lookup (08/10)"])
    slide("5. Why DNS matters in cybersecurity investigations", [
        "Nearly every connection starts with a DNS lookup → DNS logs = activity timeline",
        "Malware & C2: lookups of attacker domains, random-looking (DGA) names, many NXDOMAIN errors",
        "Phishing: look-alike & newly registered domains (g00gle-login.com)",
        "DNS tunnelling: data stolen inside long/frequent queries or TXT records",
        "Attacks on DNS: spoofing / cache poisoning, hijacking, DDoS — DNSSEC helps",
        "Attribution: WHOIS & passive DNS link domains/IPs to attackers",
        "Defence: DNS filtering & sinkholing block known-bad domains",
        "Evidence on the PC: DNS cache & hosts file (malware may edit it)"], size=18)
    slide("Summary", [
        "IP address = our identity on the network (private vs public)",
        "DHCP configures us; the default gateway gets us out; DNS finds names",
        "DNS resolution: cache → resolver → root → TLD → authoritative",
        "DNS records are key evidence in cybersecurity investigations", "Questions?"], size=22)
    p.save(path)


if __name__ == "__main__":
    build_redesign(os.path.join(HERE, "Group3_Network_Identity_DNS_Redesign.pptx"))
    build_v1(os.path.join(HERE, "Group3_Network_Identity_DNS.pptx"))
