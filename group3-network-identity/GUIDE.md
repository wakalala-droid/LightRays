# Group 3 — Tasks 2.1 & 2.2: Your Network Identity; DNS in Action

Members: Wakalala Mulyokela (Leader), Ruth Chipompo, Hampande Luyando, Loveless Chanda, Dumisani Zulu

---

## PART A — Do it on the laptop (Windows) and take the screenshots

**How to take a screenshot on Windows:** press `Windows + Shift + S`, drag over the area, then paste into the slide (Ctrl+V).
Or press `PrtScn` to capture the whole screen.

**Open Command Prompt:** press `Windows + R`, type `cmd`, press Enter.

| # | Type this command | What it shows | Screenshot name |
|---|---|---|---|
| 1 | `ipconfig` | Your IPv4 address, subnet mask, default gateway | `01-ipconfig.png` |
| 2 | `ipconfig /all` | Also shows DHCP Enabled, DHCP Server, DNS Servers, MAC (Physical) address, lease times | `02-ipconfig-all.png` |
| 3 | Open a browser → go to `https://whatismyipaddress.com` (or search "what is my ip") | Your **public** IP (the one the Internet sees) | `03-public-ip.png` |
| 4 | `nslookup google.com` | DNS turning a name into IP addresses | `04-nslookup-google.png` |
| 5 | `nslookup unza.zm` (or any local site) | Another lookup example | `05-nslookup-local.png` |
| 6 | `ipconfig /displaydns` | The DNS cache — names your PC already looked up | `06-dns-cache.png` |
| 7 | `ipconfig /flushdns` then `nslookup google.com` again | Clearing the cache, then a fresh query | `07-flushdns.png` |
| 8 | `ping google.com` | DNS resolves the name, then the PC talks to that IP | `08-ping.png` |
| 9 | `tracert google.com` | Packets leave through the **default gateway** first (first hop) | `09-tracert.png` |
| 10 | `nslookup -type=MX gmail.com` | Other DNS record types (mail servers) | `10-mx-record.png` |

**Tip:** In screenshot 2, highlight these lines before capturing (or circle them in the slide):
`IPv4 Address`, `Default Gateway`, `DHCP Enabled`, `DHCP Server`, `DNS Servers`.

**Safety tip:** Your **private** IP (e.g. 192.168.x.x / 10.x.x.x) is fine to show in class.
Consider blurring the **public** IP and the **Physical (MAC) address** in the slides.

*(Mac users: Terminal → `ifconfig`, `ipconfig getifaddr en0`, `scutil --dns`, `nslookup google.com`, `netstat -nr | grep default`.)*

---

## PART B — Answers for the presentation

### 1. What is your IP address?
An **IP (Internet Protocol) address** is the number that identifies a device on a network, like a postal address for a computer.
- **IPv4** example: `192.168.1.105` (four numbers 0–255). **IPv6** example: `fe80::1c2a:3bff:fe4d:12a0`.
- **Private IP**: used inside your home/campus network (ranges `10.0.0.0/8`, `172.16.0.0/12`, `192.168.0.0/16`). Shown by `ipconfig`.
- **Public IP**: the address your router/ISP presents to the Internet. Many devices share it through **NAT**. Shown by "what is my IP" websites.
- **Subnet mask** (e.g. `255.255.255.0`) says which part of the address is the network and which part is the device.
- *Our result:* "Our laptop's private IPv4 is ____, subnet mask ____, and our public IP is ____." (fill in from screenshots 1–3)

### 2. What are DNS, DHCP and the default gateway?
- **DHCP (Dynamic Host Configuration Protocol):** automatically gives a device its network settings when it joins: IP address, subnet mask, default gateway and DNS servers, for a limited time (the *lease*). The process is **DORA**: Discover → Offer → Request → Acknowledge. Evidence: `DHCP Enabled: Yes` and `DHCP Server` in `ipconfig /all`.
- **Default gateway:** the router your device sends traffic to when the destination is **outside** your local network. It is the "exit door" to the Internet. Usually the router, e.g. `192.168.1.1`. Evidence: the first hop in `tracert`.
- **DNS (Domain Name System):** the Internet's "phone book". It translates human-friendly names (`google.com`) into IP addresses (`142.250.x.x`) that computers use. Evidence: `DNS Servers` in `ipconfig /all` and `nslookup` output.

### 3. How does DNS convert names into IP addresses?
When you type `www.google.com`:
1. **Local cache check:** the browser and the OS check whether they already know the answer (`ipconfig /displaydns`). The `hosts` file is also checked.
2. **Ask the recursive resolver:** if the name isn't cached, the PC asks its configured DNS server (from DHCP: often the router, the ISP, or e.g. 8.8.8.8). It uses **UDP port 53**.
3. **Root server:** the resolver asks a root server, which replies "ask the `.com` TLD servers".
4. **TLD server:** the `.com` server replies "ask Google's **authoritative** name servers (ns1.google.com…)".
5. **Authoritative server:** it returns the actual record, e.g. **A record** → IPv4 address (AAAA → IPv6).
6. **Answer and cache:** the resolver returns the IP to the PC and both cache it for the record's **TTL** (time to live).
7. The browser connects to that IP.

In `nslookup`, a **"Non-authoritative answer"** means the reply came from a resolver's cache, not directly from Google's own server.
Common record types: **A** (IPv4), **AAAA** (IPv6), **CNAME** (alias), **MX** (mail server), **NS** (name server), **TXT** (text/verification, e.g. SPF).

### 4. Why is DNS important for cybersecurity investigations?
- **Almost everything starts with a DNS lookup**, malware included. DNS logs show which domains a device contacted and when. They are a timeline of activity.
- **Detecting malware and command-and-control (C2):** infected machines look up attacker domains. Signs include unusual, random-looking domain names (DGA: domain generation algorithms) and many failed lookups (NXDOMAIN).
- **Phishing investigations:** look-alike domains (e.g. `g00gle-login.com`), newly registered domains, and checking where a link really points.
- **DNS tunnelling / data exfiltration:** attackers hide stolen data in long DNS queries or TXT records. Unusually long or frequent queries are a red flag.
- **Attacks on DNS itself:** DNS spoofing / cache poisoning (fake answers send users to fake sites), DNS hijacking (changed DNS settings on a PC or router), and DDoS on DNS servers. **DNSSEC** helps verify that answers are genuine.
- **Attribution and threat intelligence:** WHOIS, passive DNS history, and which IPs and domains share infrastructure help link an attack to a group.
- **Defence:** DNS filtering/sinkholing blocks known-bad domains for a whole organisation.
- **Evidence:** the local DNS cache (`ipconfig /displaydns`) and the `hosts` file are checked during incident response, because malware sometimes edits the hosts file.

---

## PART C — Suggested slide order & who presents
1. Title + group members — *Wakalala (Leader)*
2. Our network identity: IP address (screenshots 1–3) — *Ruth*
3. DHCP & default gateway (screenshots 2, 9) — *Hampande*
4. DNS: what it is and how lookups work (diagram + screenshots 4–7) — *Loveless*
5. DNS in action: ping, MX records (screenshots 8, 10) — *Loveless*
6. Why DNS matters for cybersecurity investigations — *Dumisani*
7. Summary & questions — *Wakalala*
