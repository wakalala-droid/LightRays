"""Builds the DNS decks for Tasks 2.1 & 2.2.

  python3 build_decks.py
    -> Group3_Network_Identity_DNS_Redesign.pptx  (Wakalala Mulyokela, styled like the APE deck)
    -> Group3_Network_Identity_DNS.pptx           (Loveness Chanda, first design + cover and sponsor logos)
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
            r.font.bold = o.get("bold", bold or r.font.name in (FONT_HEAD, FONT_SEMI))
            r.font.color.rgb = rgb(o.get("color", color))
    return shp

def pill(s, x, y, label, color, size=26, h=56, pad=24, fg="#FFFFFF"):
    w = int(len(label) * size * 0.55) + pad * 2
    r = rect(s, x, y, w, h, color, radius=h / 2)
    text(s, 0, 0, 0, 0, label, size=size, color=fg, font=FONT_SEMI,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, box=r)
    return x + w

def picture(s, path, x, y, h=None, w=None):
    im = Image.open(path); ar = im.width / im.height
    if h and not w: w = h * ar
    if w and not h: h = w / ar
    return s.shapes.add_picture(path, E(x), E(y), E(w), E(h))

def footer(s, num):
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


# ---------- redesigned deck (Wakalala Mulyokela) ----------
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

    def header(self, s, kicker, title, color, dark=False):
        text(s, 134, 110, 1700, 40, kicker.upper(), size=26, color="#93C5FD" if dark else color,
             font=FONT_SEMI)
        text(s, 134, 165, 1732, 100, title, size=66, color="#F8FAFC" if dark else INK, font=FONT_HEAD)

    def card(self, s, x, y, w, h, head, body, color=None, tint=False, head_color=INK, size=34,
             head_size=42, code=None):
        rect(s, x, y, w, h, TINT[color] if tint else CARD, line=color and tint and TINT[color] or LINE,
             radius=22)
        text(s, x + 32, y + 28, w - 56, 60, head, size=head_size, color=head_color, font=FONT_HEAD)
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
    text(s, 134, 640, 1700, 60, "Who is my laptop on the network, and how does it find websites by name?",
         size=38, color="#E2E8F0")
    pill(s, 134, 855, "Wakalala Mulyokela", BLUE, size=28)

    # Q1 IP address
    s = d.slide()
    d.header(s, "Task 2.1 · Question 1 · IP address", "What is my IP address?", TEAL)
    rect(s, 134, 292, 760, 670, TEAL, radius=24)
    text(s, 180, 335, 680, 40, "WHAT I FOUND ON MY LAPTOP", size=26, color="#CCFBF1", font=FONT_SEMI)
    for i, (k, hint) in enumerate([("Private IPv4", "from ipconfig"), ("Subnet mask", "from ipconfig"),
                                   ("Default gateway", "from ipconfig"), ("Public IP", "from a “what is my IP” site")]):
        y = 400 + i * 132
        text(s, 180, y, 680, 40, [[(k, {"font": FONT_HEAD, "color": "#FFFFFF"}),
                                   ("   " + hint, {"size": 24, "color": "#99F6E4"})]], size=34)
        r = rect(s, 180, y + 50, 668, 62, "#0B5C55", radius=10)
        text(s, 0, 0, 0, 0, "____ . ____ . ____ . ____", size=32, color="#FFFFFF", font="Consolas",
             bold=True, anchor=MSO_ANCHOR.MIDDLE, box=r).text_frame.margin_left = E(20)
    d.card(s, 924, 292, 942, 205, "So what is an IP address?",
           "My laptop's home address on the network, like 192.168.1.105 (IPv4) "
           "or a longer IPv6 one.", size=30)
    d.card(s, 924, 523, 942, 205, "I actually have two",
           "A private one from my router for inside my network, and a public one "
           "the Internet sees, shared by everyone on my Wi-Fi.", size=30)
    d.shot(s, 924, 754, 942, 208, "ipconfig  ·  I blurred my public IP", TEAL)

    # Q2 DHCP / DNS / gateway
    s = d.slide()
    d.header(s, "Task 2.1 · Question 2 · DNS, DHCP and gateway", "Three settings I never set myself", ORANGE)
    cards = [("DHCP", "The moment I join the Wi-Fi, DHCP hands my laptop an IP, a gateway and DNS servers. "
                      "It's a four-step chat: Discover, Offer, Request, Acknowledge.", "DHCP Enabled: Yes"),
             ("Default gateway", "This is my router. Anything not on my local network goes out "
                                 "through it. It's my door to the Internet.", "Default Gateway: ________"),
             ("DNS", "The Internet's phone book. I type google.com, and DNS finds the number "
                     "my laptop actually needs to connect.", "DNS Servers: ________")]
    for i, (h, b, c) in enumerate(cards):
        d.card(s, 134 + i * 587, 292, 558, 410, h, b, color=ORANGE, tint=(i == 0), code=c, size=30)
    d.shot(s, 134, 728, 850, 234, "ipconfig /all  ·  DHCP, Gateway and DNS lines highlighted", ORANGE)
    d.shot(s, 1016, 728, 850, 234, "tracert google.com  ·  the first hop is my gateway", ORANGE)

    # Q3 How DNS works
    s = d.slide()
    d.header(s, "Task 2.2 · Question 3 · Name resolution", "How a name becomes an IP address", PURPLE)
    flow = [("My laptop", "Do I already know it? Checks its cache"), ("Resolver", "My DNS server goes and asks for me"),
            ("Root server", "“Try the .com servers”"), ("TLD (.com)", "“Try google.com's own servers”"),
            ("Authoritative", "“Here's the IP address”")]
    for i, (h, b) in enumerate(flow):
        x = 134 + i * 352
        hl = i == 4
        rect(s, x, 292, 300, 220, PURPLE if hl else CARD, line=None if hl else LINE, radius=22)
        text(s, x + 26, 318, 260, 50, h, size=38, color="#FFFFFF" if hl else INK, font=FONT_HEAD)
        text(s, x + 26, 378, 255, 130, b, size=29, color="#EDE9FE" if hl else BODY)
        if i < 4:
            rect(s, x + 308, 384, 36, 36, "#94A3B8", shape=MSO_SHAPE.RIGHT_ARROW)
    d.card(s, 134, 540, 850, 422, "And then",
           ["• The answer comes back to my laptop",
            "• It gets remembered for a while (the TTL), so the next visit is instant",
            "• My browser connects to that IP",
            "• When nslookup says “Non-authoritative answer”, it came from a cache"], size=30)
    d.card(s, 1016, 540, 850, 422, "Records I came across",
           [[("A ", {"bold": True, "color": PURPLE}), ("IPv4 address", {})],
            [("AAAA ", {"bold": True, "color": PURPLE}), ("IPv6 address", {})],
            [("CNAME ", {"bold": True, "color": PURPLE}), ("a nickname for another name", {})],
            [("MX ", {"bold": True, "color": PURPLE}), ("where email goes", {})],
            [("NS · TXT ", {"bold": True, "color": PURPLE}), ("name servers · notes like SPF", {})]],
           color=PURPLE, tint=True, size=30)

    # DNS in action
    s = d.slide()
    d.header(s, "Task 2.2 · DNS in action", "What I tried on my laptop", PURPLE)
    cmds = [("nslookup google.com", "Name → IP: ____________"),
            ("ipconfig /displaydns", "Names my laptop already remembers"),
            ("ipconfig /flushdns", "Wiped the memory, then looked up again"),
            ("ping google.com", "Name found, then the IP answered"),
            ("nslookup -type=MX gmail.com", "Where Gmail's email goes")]
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
    d.header(s, "Task 2.2 · Question 4 · Cybersecurity", "Why DNS matters when something goes wrong", RED)
    sec = [("It's a diary", "Almost everything starts with a lookup, so DNS logs show what a device talked to and when."),
           ("It exposes malware", "Infected machines phone home. Weird random names and lots of failed lookups give them away."),
           ("It catches phishing", "Fake look-alike sites like g00gle-login.com, often registered just days before."),
           ("It can leak data", "Attackers can sneak stolen data out inside DNS queries. That's called tunnelling."),
           ("It gets attacked", "Fake answers (spoofing), hijacked settings, floods. DNSSEC helps prove answers are real."),
           ("It's evidence", "Blocking bad domains protects everyone, and my DNS cache and hosts file are worth checking.")]
    for i, (h, b) in enumerate(sec):
        x = 134 + (i % 3) * 587; y = 292 + (i // 3) * 318
        d.card(s, x, y, 558, 292, h, b, color=RED, tint=(i == 5), size=30, head_size=40)
    text(s, 134, 938, 1700, 40, "Investigators also use WHOIS and old DNS records to link domains back to the same attacker.",
         size=27, color=MUTED)

    # Summary
    s = d.slide(dark=True)
    d.header(s, "Summary", "What I took away", BLUE, dark=True)
    rect(s, 134, 330, 1732, 330, "#1A3558", radius=24)
    text(s, 180, 360, 1640, 270,
         [[("My IP address ", {"bold": True, "color": "#FFFFFF"}), ("is my laptop's identity: private at home, public to the world.", {})],
          [("DHCP ", {"bold": True, "color": "#FFFFFF"}), ("sets me up, the ", {}), ("gateway ", {"bold": True, "color": "#FFFFFF"}),
           ("lets me out, ", {}), ("DNS ", {"bold": True, "color": "#FFFFFF"}), ("finds the way.", {})],
          [("A lookup ", {"bold": True, "color": "#FFFFFF"}), ("goes cache → resolver → root → TLD → authoritative.", {})],
          [("For security, ", {"bold": True, "color": "#FFFFFF"}), ("DNS is one of the first places I'd look.", {})]],
         size=36, color="#E2E8F0", anchor=MSO_ANCHOR.MIDDLE, after=10)
    text(s, 134, 790, 1700, 100, "Thank you. Any questions?", size=80, color="#F8FAFC", font=FONT_HEAD)
    pill(s, 134, 905, "Wakalala Mulyokela", BLUE, size=28)

    # Key terms
    s = d.slide()
    text(s, 134, 110, 1200, 40, "EXTRA · IN CASE IT COMES UP", size=26, color=INK, font=FONT_SEMI)
    text(s, 134, 165, 1700, 100, "Key terms, in plain words", size=66, color=INK, font=FONT_HEAD)
    terms = [("IP address", "My device's address on a network", BLUE), ("NAT", "Lots of private IPs sharing one public IP", BLUE),
             ("DHCP", "Hands out network settings automatically", TEAL), ("Gateway", "The router that gets me out", ORANGE),
             ("DNS", "Turns names into IP addresses", ORANGE), ("Resolver", "The DNS server doing the asking for me", PURPLE),
             ("TTL", "How long an answer is remembered", PURPLE), ("NXDOMAIN", "“That name doesn't exist”", RED),
             ("DNSSEC", "Proof that a DNS answer is genuine", RED)]
    for i, (h, b, c) in enumerate(terms):
        x = 134 + (i % 3) * 587; y = 292 + (i // 3) * 230
        d.card(s, x, y, 558, 204, h, b, head_color=c, size=32, head_size=40)

    d.save(path)


# ---------- first design, with cover and sponsor logos (Loveness Chanda) ----------
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
                                     ("Loveness Chanda", 22, False)]):
        q = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        q.text = t; q.font.size = Pt(sz); q.font.bold = b; q.font.color.rgb = RGBColor(255, 255, 255)
    footer(s, 2)

    slide("1. What is my IP address?", [
        "An IP address is basically my laptop's home address on a network",
        "It looks like 192.168.1.105 (IPv4), or a longer IPv6 one like fe80::1c2a:…",
        "My private IP (inside my network): ____________", "- I found it by running ipconfig",
        "My subnet mask: ____________", "My public IP (what the Internet sees): ____________",
        "- everyone on my Wi-Fi shares this one, thanks to NAT",
        "Private addresses start with 10., 172.16–31. or 192.168."],
        ["ipconfig", "My public IP (blurred)"])
    slide("2. DHCP, DNS and the default gateway", [
        "DHCP: when I join the Wi-Fi, it hands my laptop an IP, gateway and DNS servers",
        "- four steps: Discover, Offer, Request, Acknowledge",
        "- ipconfig /all shows \"DHCP Enabled: Yes\" and the DHCP server",
        "Default gateway: my router, the way out to the Internet",
        "- mine is ____________ (it's the first hop in tracert)",
        "DNS: the Internet's phone book, turning names into IP addresses", "- my DNS servers: ____________"],
        ["ipconfig /all (DHCP, Gateway and DNS lines highlighted)", "tracert google.com"])
    slide("3. How DNS turns a name into an IP address", [
        "1. My laptop checks if it already knows the answer (its cache)",
        "2. If not, it asks my DNS server (the resolver)",
        "3. The resolver asks a root server, which says \"try the .com servers\"",
        "4. The .com server says \"try google.com's own servers\"",
        "5. Google's server gives back the IP address (the A record)",
        "6. My laptop remembers it for a while (the TTL) and connects",
        "Other records: AAAA (IPv6), CNAME (nickname), MX (email), NS, TXT"],
        ["Diagram: my laptop → resolver → root → .com → Google's server"], size=17)
    slide("4. DNS in action on my laptop", [
        "nslookup google.com gave me: ____________",
        "- \"Non-authoritative answer\" just means it came from a cache",
        "ipconfig /displaydns shows names my laptop already remembers",
        "ipconfig /flushdns wipes that memory, so the next lookup is fresh",
        "ping google.com finds the name, then talks to the IP",
        "nslookup -type=MX gmail.com shows where Gmail's email goes"],
        ["nslookup google.com", "ipconfig /displaydns or /flushdns", "ping or MX lookup"])
    slide("5. Why DNS matters in cybersecurity investigations", [
        "Almost everything starts with a lookup, so DNS logs are like a diary of what a device did",
        "Malware gives itself away by looking up strange, random names and failing a lot",
        "Phishing sites use look-alike names like g00gle-login.com",
        "Attackers can sneak data out hidden inside DNS queries (tunnelling)",
        "DNS itself gets attacked: fake answers, hijacked settings, floods. DNSSEC helps",
        "WHOIS and old DNS records help link domains back to the same attacker",
        "Blocking bad domains protects a whole network",
        "On a PC, the DNS cache and hosts file are worth checking"], size=18)
    slide("What I took away", [
        "My IP address is my laptop's identity: private at home, public to the world",
        "DHCP sets me up, the gateway lets me out, DNS finds the way",
        "A lookup goes: cache → resolver → root → .com → the site's own server",
        "When something goes wrong, DNS is one of the first places to look",
        "Thank you. Any questions?"], size=22)
    p.save(path)


if __name__ == "__main__":
    build_redesign(os.path.join(HERE, "Group3_Network_Identity_DNS_Redesign.pptx"))
    build_v1(os.path.join(HERE, "Group3_Network_Identity_DNS.pptx"))
