"""Canonical facts for Chris Childs career documents.

Every claim here is taken from the v14 resume on
https://github.com/chchilds/cxvx/pull/9, from the LinkedIn profile
https://www.linkedin.com/in/chrischildsraleigh (public, labeled), or from
open PRs in this repository. Nothing here invents employment, awards,
certifications, or closed-win amounts.
"""

from __future__ import annotations

AS_OF = "November 2026"
VERSION = "v15"

NAME = "Christopher E. Childs"
NAME_SHORT = "Chris Childs"
HEADLINE = (
    "Premium Services Expert  ·  Strategic Software & Services Architect  ·  Cisco Sales Champion"
)

# Street address and phone are from resume v14 (PR #9), confirmed there by Chris.
ADDRESS = "6507 Glendower Rd, Raleigh, North Carolina 27613"
PHONE = "(919) 523-8426"
EMAIL_CISCO = "Chchilds@cisco.com"
EMAIL_PERSONAL = "Chris.E.Childs@gmail.com"
LINKEDIN_TEXT = "linkedin.com/in/chrischildsraleigh"
LINKEDIN_URL = "https://www.linkedin.com/in/chrischildsraleigh"

LOCATION_WORK = "Raleigh, NC"
COMPANY = "Cisco Systems, Inc."

# May 2012 start → November 2026 is 14 years 6 months. Keep the v14 "14+" wording.
SUMMARY = (
    "Customer-focused Cisco combined seller and solutions engineer based in Raleigh, "
    "North Carolina, with 14+ years at Cisco spanning premium services, strategic software, "
    "Enterprise Agreements, and pre-sales architecture. Designs consumption models that "
    "attach CX Enterprise Agreements (CXEA), Expert Care, Lifecycle Services, and Advanced "
    "Services to software ELAs — and converts that design work into competitive wins across "
    "security, collaboration, data center, wireless, and enterprise networking."
)

# LinkedIn (Nov 2026) lists Kansas City, Kansas; resume v14 and PR #9 use Raleigh.
# This document set keeps Raleigh because Chris directed that update.
LOCATION_NOTE = (
    "Work location on these documents is Raleigh, NC, matching resume v14 / PR #9. "
    "The public LinkedIn profile currently shows Kansas City, Kansas. "
    "Confirm with Chris before changing either."
)

ROLES = [
    {
        "title": "EA Architect & Premium Service Seller (Combined Seller)",
        "dates": "August 2023 – Present",
        "company": COMPANY,
        "location": LOCATION_WORK,
        "bullets": [
            "Combined seller for Cisco Premium Services and Strategic Software: architect CX Enterprise Agreements (CXEAs), Expert Care, Lifecycle Services (LCS), and Advanced Services (AS) as a single customer motion.",
            "Design software-plus-services consumption models that attach Premium Services to Enterprise Agreements and drive recurring-revenue, True Forward, and high-adoption outcomes.",
            "Partner with account teams, partners, and customers to position One Cisco architectures across networking, security, collaboration, and data center software.",
            "Produce Enterprise Agreement and CX services-attach collateral — proposal decks, deal-financial analysis, and seller one-pagers — so account teams can quote software and services as one motion.",
            "Recognized as a Cisco Sales Champion and GSX winner for software and services sales; completed IMPACT Ready IC (Aug 2023) and FY24 Services Sales (Oct 2023).",
        ],
    },
    {
        "title": "Strategic Software Sales Specialist",
        "dates": "October 2019 – Present",
        "company": COMPANY,
        "location": LOCATION_WORK,
        "note": "Concurrent with the Combined Seller role from August 2023. Both titles are current on LinkedIn and on resume v14.",
        "bullets": [
            "Lead strategic software sales for Enterprise Agreements, ELA consumption, and subscription / recurring-revenue offers across Cisco’s software portfolio.",
            "Named FY22 Americas Sales Champion, Services & Software — recognized among Cisco’s top Americas software and services sellers.",
            "Translate security, collaboration, networking, and data center software into customer consumption and adoption plans; completed Cisco Recurring Revenue Metrics (Feb 2022).",
            "Carry forward SE-honed competitive displacement skills — SD-WAN, Meraki, security, contact center, and API/scripting use cases — into software-led deals.",
        ],
    },
    {
        "title": "Systems Engineer, Carolina Select",
        "dates": "January 2015 – October 2019",
        "company": COMPANY,
        "location": LOCATION_WORK,
        "bullets": [
            "Technical and engineering lead for Cisco Select healthcare, financial, and manufacturing accounts in North and South Carolina. Competitive wins in security, collaboration, data center, wireless, enterprise, management, and software, delivered with Advanced Services or partner-led engagements.",
            "Drove blended 100% ELA / software consumption spanning UC, ONE (C1), and Security; used CloudCenter, ACI, and UCS Director in competitive data-center-to-cloud conversations modeled on Cisco-on-Cisco catalog offerings.",
            "SE of multiple quarters in Commercial; seven Carolina Select customers; customer-referenced architectures, EBCs, Global CRO travel, Cisco Live Expo for LED / UPOE with CREE, plus scripting and API use cases.",
            "Major wins in SD-WAN, Meraki, Security portfolio, and Contact Center; full post-sales and customer-excellence focus.",
        ],
    },
    {
        "title": "Network Consulting Engineer",
        "dates": "March 2013 – January 2015",
        "company": COMPANY,
        "location": LOCATION_WORK,
        "bullets": [
            "Designed and implemented WAAS and ACE for production data centers (HLD, LLD, NIP); 90%+ customer-facing time; six Advanced Services enterprise customers; UCS deploy and on-site knowledge transfers with TAC and BU partners.",
            "Configured and optimized data-center networks; worked with Cisco TAC and business units on enterprise deployments; planning and deliverables for project management; high customer response and satisfaction feedback.",
        ],
    },
    {
        "title": "Network Engineer, Remote Management Services",
        "dates": "May 2012 – March 2013",
        "company": COMPANY,
        "location": LOCATION_WORK,
        "bullets": [
            "NOC incident handling for Cisco premier customers (Remedy, Cisco MAP, TAC): LAN/WLAN, Nexus 7000 NX-OS, ASR5k 3G, ASR9k IOS XR LTE, backhaul, TACACS+/ACS, IronPort, IPsec/VPN, BGP/OSPF/EIGRP; built handover, outage-mailer, and NMS-access processes with SLA/SLO reporting.",
            "Global customer support including data-center RMS for a mobile wireless provider; developed web interface for interactive shift and incident handover; improved NOC immersion for new team members.",
        ],
    },
]

EARLIER = [
    {
        "title": "Realtor",
        "org": "Smart Choice Realty and ERA United Home Realty",
        "location": "Raleigh, NC",
        "dates": "2006–2012",
        "text": "Built and maintained the firm website and email server; client database search, digital photos, PDF contracts, and e-sign storage.",
    },
    {
        "title": "Sales Associate / Geek Squad",
        "org": "Best Buy",
        "location": "Cary, NC",
        "dates": "2004–2006",
        "text": "Computer sales and home/small-business networking support; MVP award for Exemplary Service.",
    },
    {
        "title": "General Manager",
        "org": "Papa John’s",
        "location": "Cary, NC",
        "dates": "2000–2004",
        "text": "POS/credit-card connectivity, customer database, web-ordering coupon sync, and productivity / P&L / inventory reporting.",
    },
]

EDUCATION = [
    {
        "school": "East Carolina University",
        "credential": "B.S., Industrial Technology",
        "dates": "Completed 2018",
        "note": "ECU: recruiting for Cisco Advanced Services internships.",
    },
    {
        "school": "Wake Technical Community College",
        "credential": "A.A.S., Networking Technology, GPA 3.7",
        "dates": "2013",
        "note": "Wake Tech: student leader / peer tutor; Disability Support Services presentation on transitioning for students with disabilities.",
    },
]

# v14 and PR #9: "valid through December 2026". As of 27 Nov 2026 that window
# has not closed. Recertification after December 2026 is not in the sources.
CERTS_CURRENT_LABEL = "Cisco certifications (listed as valid through December 2026)"
CERTS_CURRENT = "CCNP Routing & Switching; CCDP Design; CCNP Data Center; CCNA Voice."
SALES_ENABLEMENT = (
    "FY24 Services Sales (Oct 2023); IMPACT Ready ICs (Aug 2023); "
    "Cisco Recurring Revenue Metrics (Feb 2022)."
)
CERTS_ADDITIONAL = "CCIE Data Center written passed (lab date Nov 2017); VCP 6.0 DCV."

SKILLS = [
    (
        "Software & services sales",
        "Enterprise Agreements / ELA, CXEA, Premium Services, Expert Care, LCS / PS-Subscription, Advanced Services, consumption modeling, recurring revenue, True Forward, One Cisco",
    ),
    (
        "Architecture",
        "SD-WAN, Meraki, Security portfolio, Contact Center / UC, ACI, UCS, CloudCenter, wireless, campus / WAN, UPOE",
    ),
    (
        "Engineering",
        "Routing & switching (BGP, OSPF, EIGRP, STP/VLANs), NX-OS / IOS / IOS XR, ASA / IronPort, IPsec VPN, QoS, IPv6, packet analysis",
    ),
    (
        "Automation & tools",
        "Scripting and Cisco APIs, Wireshark, SNMP, Remedy, documentation, customer-referenced architectures",
    ),
    (
        "Personal",
        "Customer-facing, multi-account ownership, solution-oriented, communicative, time management, flexible",
    ),
]

MOTIONS = [
    {
        "name": "CX Enterprise Agreement (CXEA) + software ELA",
        "text": "Architect software-plus-services consumption so Premium Services attach to the Enterprise Agreement rather than sitting as a follow-on quote.",
    },
    {
        "name": "Expert Care, Lifecycle Services / PS-Subscription, Advanced Services",
        "text": "Position Support, PS-Sub (formerly LCS), and PS-Transactional as distinct vehicles: keep-it-running, govern-advise-measure, and defined-scope hands-on work.",
    },
    {
        "name": "Recurring revenue and True Forward",
        "text": "Design adoption and True Forward outcomes into the commercial model, not as a back-end cleanup.",
    },
    {
        "name": "Services attach on product-led deals",
        "text": "Partner with account teams so software and CX services quote on the same close date, with a documented exception path when a customer refuses attach.",
    },
]

# Open PRs in this public repo. These are seller/proposal materials, not
# attested closed wins. Do not copy TCV or pipeline figures into the resume.
WORK_SAMPLES = [
    {
        "pr": 16,
        "title": "Polaris Inc Cisco EA proposal (EAMP 9354137)",
        "kind": "Enterprise Agreement adder proposal deck",
        "url": "https://github.com/chchilds/cxvx/pull/16",
    },
    {
        "pr": 15,
        "title": "Tallgrass H1 Services Remediation Action Plan",
        "kind": "Regional services-attach action plan",
        "url": "https://github.com/chchilds/cxvx/pull/15",
    },
    {
        "pr": 13,
        "title": "Hallmark Cards Cisco Enterprise Agreement proposal PPTX",
        "kind": "EA 3.0 proposal deck",
        "url": "https://github.com/chchilds/cxvx/pull/13",
    },
    {
        "pr": 12,
        "title": "Halifax Health DFA Enterprise Agreement PPTX (Proposal 9322988)",
        "kind": "Deal-financial-analysis EA deck",
        "url": "https://github.com/chchilds/cxvx/pull/12",
    },
    {
        "pr": 10,
        "title": "PS-Sub / LCS abilities one-pager PDF",
        "kind": "Seller conversation sheet (Cisco Confidential)",
        "url": "https://github.com/chchilds/cxvx/pull/10",
    },
    {
        "pr": 9,
        "title": "Update Chris Childs 2026 resume with recent Cisco seller roles",
        "kind": "Resume v14 source (predecessor of this v15 set)",
        "url": "https://github.com/chchilds/cxvx/pull/9",
    },
    {
        "pr": 8,
        "title": "Add Cisco Tallgrass Region RIS Assessment execution plan",
        "kind": "RIS execution plan",
        "url": "https://github.com/chchilds/cxvx/pull/8",
    },
    {
        "pr": 7,
        "title": "Add Halifax Health Enterprise Agreement PPTX (Proposal 9327623)",
        "kind": "EA proposal deck",
        "url": "https://github.com/chchilds/cxvx/pull/7",
    },
    {
        "pr": 4,
        "title": "Add Hallmark Cards Enterprise Agreement summary PowerPoint",
        "kind": "EA summary deck",
        "url": "https://github.com/chchilds/cxvx/pull/4",
    },
    {
        "pr": 1,
        "title": "Add Seaboard Foods Enterprise Agreement summary PowerPoint",
        "kind": "EA summary deck",
        "url": "https://github.com/chchilds/cxvx/pull/1",
    },
]

BIO_PARAS = [
    (
        f"{NAME} is a Cisco Premium Services and Strategic Software combined seller "
        f"based in Raleigh, North Carolina. He designs CX Enterprise Agreements, Expert Care, "
        f"Lifecycle Services, and Advanced Services as a single customer motion, attaching "
        f"those services to software Enterprise Agreements rather than treating them as a "
        f"follow-on quote."
    ),
    (
        "Chris has been at Cisco since May 2012. He moved from Remote Management Services "
        "and Advanced Services consulting into a Carolina Select Systems Engineer role, then "
        "into Strategic Software Sales (October 2019–present) and the Combined Seller / "
        "EA Architect seat (August 2023–present). He was named FY22 Americas Sales Champion, "
        "Services & Software, and is a Cisco Sales Champion and GSX winner for software and "
        "services sales."
    ),
    (
        "Earlier Cisco work includes competitive displacements in SD-WAN, Meraki, security, "
        "and contact center; blended ELA consumption across UC, ONE, and Security; and "
        "production data-center design and implementation of WAAS, ACE, and UCS. He holds "
        "a B.S. in Industrial Technology from East Carolina University (2018) and an A.A.S. "
        "in Networking Technology from Wake Technical Community College (GPA 3.7, 2013). "
        "Cisco certifications listed on the 2026 resume — CCNP Routing & Switching, CCDP, "
        "CCNP Data Center, and CCNA Voice — are recorded as valid through December 2026."
    ),
]

OPEN_QUESTIONS = [
    "LinkedIn currently lists Kansas City, Kansas; resume v14 uses Raleigh, NC. These documents keep Raleigh.",
    "Cisco certifications are listed as valid through December 2026. Recertification after that date is not in the source materials.",
    "Customer EA decks and the Tallgrass action plan are Cisco Confidential indicative / pipeline materials. They are indexed as work samples, not closed-win claims, and dollar amounts from those PRs are not copied here.",
    "Confirm whether the street address should remain on public GitHub copies of the resume.",
]
