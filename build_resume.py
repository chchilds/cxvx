#!/usr/bin/env python3
"""Fill the original Word resume template with 2026 content and write a .docx."""

from __future__ import annotations

import os
import shutil
from copy import deepcopy
from zipfile import ZipFile, ZIP_DEFLATED
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
NSMAP = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
}

ET.register_namespace("w", NSMAP["w"])
ET.register_namespace("w14", NSMAP["w14"])
ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
ET.register_namespace("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas")
ET.register_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")

SRC = "/home/ubuntu/.cursor/projects/workspace/uploads/Resume_-_Chris_Childs_2026__v13__a881.docx"
OUT = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "Resume - Chris Childs 2026 (v14).docx",
)


def q(tag: str) -> str:
    return f"{W}{tag}"


def clone_para(template: ET.Element, text: str) -> ET.Element:
    """Clone a paragraph and replace its runs with a single text run."""
    p = deepcopy(template)
    rpr = None
    for child in list(p):
        if child.tag == q("r"):
            if rpr is None:
                found = child.find(q("rPr"))
                if found is not None:
                    rpr = deepcopy(found)
            p.remove(child)
        elif child.tag.endswith("}hyperlink"):
            p.remove(child)
    run = ET.SubElement(p, q("r"))
    if rpr is not None:
        run.append(rpr)
    t = ET.SubElement(run, q("t"))
    if text[:1].isspace() or text[-1:].isspace():
        t.set(XML_SPACE, "preserve")
    t.text = text
    return p


def clone_job_meta(template: ET.Element, company: str, dates: str) -> ET.Element:
    """Company on the left, dates after a tab, matching the original layout."""
    p = deepcopy(template)
    rpr = None
    date_rpr = None
    runs = p.findall(q("r"))
    if runs:
        found = runs[0].find(q("rPr"))
        if found is not None:
            rpr = deepcopy(found)
        # last non-tab run often holds the date formatting
        for r in reversed(runs):
            if r.find(q("tab")) is None and r.find(q("t")) is not None:
                found = r.find(q("rPr"))
                if found is not None:
                    date_rpr = deepcopy(found)
                break
    for child in list(p):
        if child.tag in (q("r"),) or child.tag.endswith("}hyperlink"):
            p.remove(child)

    def add_run(txt: str, props):
        run = ET.SubElement(p, q("r"))
        if props is not None:
            run.append(deepcopy(props))
        t = ET.SubElement(run, q("t"))
        if txt[:1].isspace() or txt[-1:].isspace():
            t.set(XML_SPACE, "preserve")
        t.text = txt

    add_run(company, rpr)
    tab_run = ET.SubElement(p, q("r"))
    if rpr is not None:
        tab_run.append(deepcopy(rpr))
    ET.SubElement(tab_run, q("tab"))
    add_run(dates, date_rpr if date_rpr is not None else rpr)
    return p


def empty_para(template: ET.Element) -> ET.Element:
    p = deepcopy(template)
    for child in list(p):
        if child.tag == q("r") or child.tag.endswith("}hyperlink"):
            p.remove(child)
    return p


def cell_tcpr(cell: ET.Element) -> ET.Element:
    tcpr = cell.find(q("tcPr"))
    if tcpr is None:
        raise RuntimeError("cell missing tcPr")
    return deepcopy(tcpr)


def replace_cell(cell: ET.Element, paragraphs: list[ET.Element]) -> None:
    tcpr = cell_tcpr(cell)
    for child in list(cell):
        cell.remove(child)
    cell.append(tcpr)
    for p in paragraphs:
        cell.append(p)


def build() -> None:
    if not os.path.exists(SRC):
        raise SystemExit(f"Original Word template not found: {SRC}")

    shutil.copy2(SRC, OUT)

    with ZipFile(SRC) as z:
        xml = z.read("word/document.xml")
        other = {n: z.read(n) for n in z.namelist() if n != "word/document.xml"}

    root = ET.fromstring(xml)
    tbl = root.find(f".//{q('tbl')}")
    rows = tbl.findall(q("tr"))
    name_cell, contact_cell = rows[0].findall(q("tc"))
    body_cell, skills_cell = rows[1].findall(q("tc"))
    refs_cell, empty_right = rows[2].findall(q("tc"))

    left_paras = body_cell.findall(q("p"))
    heading = left_paras[1]  # OBJECTIVE with bottom border
    body_txt = left_paras[2]
    job_title = left_paras[16]
    job_meta = left_paras[17]
    highlights = left_paras[18]
    bullet = left_paras[19]
    cert_line = left_paras[6]
    spacer = left_paras[0]
    skill_heading = rows[1].findall(q("tc"))[1].findall(q("p"))[1]
    skill_cat = rows[1].findall(q("tc"))[1].findall(q("p"))[3]
    skill_item = rows[1].findall(q("tc"))[1].findall(q("p"))[4]
    name_p = name_cell.findall(q("p"))[0]
    contact_p = contact_cell.findall(q("p"))[0]
    refs_h = refs_cell.findall(q("p"))[2]
    refs_b = refs_cell.findall(q("p"))[4]

    # --- Header ---
    replace_cell(name_cell, [clone_para(name_p, "Christopher E. Childs")])
    replace_cell(
        contact_cell,
        [
            clone_para(contact_p, "6507 Glendower Rd"),
            clone_para(contact_p, "Raleigh, North Carolina 27613"),
            clone_para(contact_p, "Chchilds@cisco.com"),
            clone_para(contact_p, "Chris.E.Childs@gmail.com"),
            clone_para(contact_p, "(919) 523-8426"),
            clone_para(contact_p, "linkedin.com/in/chrischildsraleigh"),
        ],
    )

    left: list[ET.Element] = []
    left.append(empty_para(spacer))
    left.append(clone_para(heading, "OBJECTIVE"))
    left.append(
        clone_para(
            body_txt,
            "I am a solutions-driven, customer-savvy combined seller and EA architect focused on "
            "ELA / Enterprise Agreement software and Premium Services consumption modeling, seeking "
            "an opportunity to test and expand my skills through delivering superior client outcomes.",
        )
    )
    left.append(empty_para(spacer))
    left.append(clone_para(heading, "EDUCATION"))
    left.append(clone_para(cert_line, "CCNP Routing & Switching          Valid through 12/2026"))
    left.append(clone_para(cert_line, "CCDP Design                       Valid through 12/2026"))
    left.append(clone_para(cert_line, "CCNP Data Center                  Valid through 12/2026"))
    left.append(clone_para(cert_line, "CCNA Voice                        Valid through 12/2026"))
    left.append(clone_para(cert_line, "CCIE Data Center Written passed — lab date 11/2017"))
    left.append(clone_para(cert_line, "VCP 6.0 DCV; FY24 Services Sales; IMPACT Ready ICs; Recurring Revenue Metrics"))
    left.append(clone_para(cert_line, "Wake Technical Community College — Networking Technology AAS — 2013, GPA 3.7"))
    left.append(clone_para(cert_line, "East Carolina University — B.S. Industrial Technology — Completed 2018"))
    left.append(empty_para(spacer))
    left.append(clone_para(heading, "EXPERIENCE"))

    def add_job(title, company, dates, highlight, bullets):
        left.append(clone_para(job_title, title))
        left.append(clone_job_meta(job_meta, company, dates))
        left.append(clone_para(highlights, highlight))
        for b in bullets:
            left.append(clone_para(bullet, b))
        left.append(empty_para(spacer))

    add_job(
        "EA Architect and Premium Service Seller — Combined Seller",
        "Cisco Systems, Inc., Raleigh, NC",
        "Aug 2023 – Present",
        "Project Highlights include: Combo seller for Cisco Premium Services and Strategic Software — "
        "CX Enterprise Agreements (CXEAs), Expert Care, Lifecycle Services (LCS), and Advanced Services (AS). "
        "Architect software-plus-services consumption models that attach Premium Services to Enterprise "
        "Agreements and drive recurring-revenue outcomes.",
        [
            "One Cisco architectures across networking, security, collaboration, and data center software",
            "Cisco Sales Champion and GSX winner for software and services sales",
            "IMPACT Ready IC (Aug 2023) and FY24 Services Sales (Oct 2023)",
            "Full customer-facing combined seller motion with account teams and partners",
        ],
    )
    add_job(
        "Strategic Software Sales Specialist",
        "Cisco Systems, Inc., Raleigh, NC",
        "Oct 2019 – Present",
        "Project Highlights include: Strategic software sales for Enterprise Agreements, ELA consumption, "
        "and subscription / recurring-revenue offers across Cisco’s software portfolio.",
        [
            "FY22 Americas Sales Champion, Services & Software",
            "Security, collaboration, networking, and data center software consumption plans",
            "Cisco Recurring Revenue Metrics (Feb 2022)",
            "SD-WAN, Meraki, security, contact center, and API/scripting use cases in software-led deals",
        ],
    )
    add_job(
        "Systems Engineer, Carolina Select",
        "Cisco Systems, Inc., Raleigh, NC",
        "Jan 2015 – Oct 2019",
        "Project Highlights include: Technical and Engineering coordination for Cisco Select Healthcare, "
        "Financial, and Manufacturing accounts based in South Carolina and North Carolina. Numerous "
        "competitive wins in Security, Collaboration, Data Center, Wireless, Enterprise, Management, and "
        "Software, blanketed in Advanced Services or partner-led engagements. Blended 100% consumption of "
        "ELA/software adoption consisting of UC, C1, and Security. CloudCenter, ACI, and UCS Director "
        "leveraged in competitive data-center-to-cloud conversations and catalog offerings modeled after "
        "Cisco-on-Cisco use cases. Driving UPOE and LED lighting with CREE.",
        [
            "Customer Referenced Architecture delivered to every customer",
            "Onsite and customer facing; travel for EBCs and Global CRO experiences",
            "Full post-sales and customer excellence focused",
            "Cisco Live Expo for LED and UPOE experiences",
            "SE of multiple quarters in the Commercial region",
            "Maintain and support 7 Carolina Select customers",
            "High customer response and satisfaction feedback",
            "SD-WAN, Meraki, Security portfolio, Contact Center, huge wins",
            "Scripting and API use cases for Cisco offerings",
        ],
    )
    add_job(
        "Network Consulting Engineer",
        "Cisco Systems, Inc., Raleigh, NC",
        "Mar 2013 – Jan 2015",
        "Project Highlights include: Data Center Design and Implement WAAS and ACE appliances for customer "
        "production environments. This includes HLD and LLD as well as Network Implementation Plan "
        "documentation. Configure and Optimize Data Center networks. Work closely with Cisco TAC Engineers "
        "and BU to develop customer solutions, and troubleshoot various Cisco Data Center scenarios for "
        "enterprise deployments.",
        [
            "Daily customer contact (Customer facing time +90% per quarter)",
            "Planning and Deliverables for Project Management",
            "Maintain and support 6 Advanced Services Enterprise Customers",
            "High customer response and satisfaction feedback",
            "UCS support and deployment",
            "On-site for Knowledge Transfers for customer trainings",
            "WAAS and ACE design, configuration, and deployment",
        ],
    )
    add_job(
        "Network Engineer",
        "Cisco Systems, Inc., Raleigh, NC",
        "May 2012 – Mar 2013",
        "Project Highlights include: Handling incidents for Cisco customers in a NOC environment, "
        "troubleshooting circuits and internal customer networks, utilizing Remedy and Cisco MAP for "
        "managing incidents, working closely with Cisco TAC Engineers to develop customer solutions, and "
        "troubleshooting various Cisco routers, switches and security appliances for both large campus and "
        "local branch deployments.",
        [
            "Configured and solve connectivity issues for LANs and Wireless LANs for Cisco premier customers",
            "Nexus NX-OS 7000, ASR5k 3G, ASR9k IOS XR LTE, cell tower backhaul, TACACS+/ACS, IronPort, IPsec/VPN",
            "BGP, OSPF, EIGRP troubleshooting; SLA/SLO reporting and customer delivery",
            "Developed NOC processes, outage email templates, and a web interface for shift/incident handover",
            "Technical Environment (Development): Scripting, web pages",
            "Technical Environment (Creation): Numerous NOC processes",
        ],
    )
    add_job(
        "Realtor",
        "Smart Choice Realty and ERA United Home Realty, Raleigh, NC",
        "2006 – 2012",
        "Project Highlights: Created and maintained a website and email server for large Real Estate firm. "
        "Data Entry and digital photo uploads, client database searches and search parameter settings; "
        "Website/Email server design and upkeep; PDF contracts and digital document signing and image "
        "storage and contract retrieval.",
        [],
    )
    add_job(
        "Sales Associate / Geek Squad",
        "Best Buy, Cary, NC",
        "2004 – 2006",
        "Project Highlights: Computer Department — Customer Sales; technical assistance with home and "
        "small business networking. MVP award for Exemplary Service.",
        [],
    )
    add_job(
        "General Manager",
        "Papa John's, Cary, NC",
        "2000 – 2004",
        "Project Highlights: Verification and troubleshooting end device and credit card machine internet "
        "connectivity, customer database upkeep, synchronized in store coupons to website online ordering, "
        "created productivity/profit & loss reports and inventory database reporting.",
        [],
    )

    left.append(clone_para(heading, "REFERENCES"))
    left.append(
        clone_para(
            refs_b,
            "Professional and personal references - available on request.",
        )
    )
    replace_cell(body_cell, left)

    skills: list[ET.Element] = [
        empty_para(spacer),
        clone_para(skill_heading, "Skills"),
        empty_para(spacer),
        clone_para(skill_cat, "Software & Services"),
        clone_para(skill_item, "ENTERPRISE AGREEMENTS"),
        clone_para(skill_item, "ELA / CXEA"),
        clone_para(skill_item, "PREMIUM SERVICES"),
        clone_para(skill_item, "EXPERT CARE / LCS / AS"),
        clone_para(skill_item, "CONSUMPTION MODELING"),
        clone_para(skill_item, "RECURRING REVENUE"),
        empty_para(spacer),
        clone_para(skill_cat, "Personal"),
        clone_para(skill_item, "MULTI-TASKING"),
        clone_para(skill_item, "SOLUTION ORIENTED"),
        clone_para(skill_item, "COMMUNICATIVE"),
        clone_para(skill_item, "TIME-MANAGEMENT"),
        clone_para(skill_item, "FLEXIBLE"),
        empty_para(spacer),
        clone_para(skill_cat, "Networking"),
        clone_para(skill_item, "ACCESS CONTROL"),
        clone_para(skill_item, "ACL/VACL/RACL"),
        clone_para(skill_item, "DHCP"),
        clone_para(skill_item, "DNS"),
        clone_para(skill_item, "DOCUMENTATION"),
        clone_para(skill_item, "ETHERNET"),
        clone_para(skill_item, "IPv6"),
        clone_para(skill_item, "NAT / PAT"),
        clone_para(skill_item, "PACKET ANALYSIS"),
        clone_para(skill_item, "QoS"),
        clone_para(skill_item, "OSPF"),
        clone_para(skill_item, "BGP"),
        clone_para(skill_item, "VTP"),
        clone_para(skill_item, "PORT CHANNEL"),
        clone_para(skill_item, "ROUTING"),
        clone_para(skill_item, "SWITCH SECURITY"),
        clone_para(skill_item, "SNMP"),
        clone_para(skill_item, "STP, MST, PVST"),
        clone_para(skill_item, "TCP/IP"),
        clone_para(skill_item, "UC SUPPORT"),
        clone_para(skill_item, "VLAN / PVLAN"),
        clone_para(skill_item, "VOICE"),
        clone_para(skill_item, "VPN"),
        clone_para(skill_item, "WIRELESS"),
        clone_para(skill_item, "SD-WAN"),
        clone_para(skill_item, "MERAKI"),
        clone_para(skill_item, "WAN"),
        empty_para(spacer),
        clone_para(skill_cat, "Firewalls"),
        clone_para(skill_item, "CISCO ASA"),
        clone_para(skill_item, "SECURITY PORTFOLIO"),
        empty_para(spacer),
        clone_para(skill_cat, "Monitoring"),
        clone_para(skill_item, "REMEDY"),
        clone_para(skill_item, "CISCO MAP"),
        clone_para(skill_item, "EMAIL REPORTS"),
        clone_para(skill_item, "SNMP"),
        clone_para(skill_item, "WIRESHARK"),
        empty_para(spacer),
        clone_para(skill_cat, "Cisco"),
        clone_para(skill_item, "IOS"),
        clone_para(skill_item, "XR"),
        clone_para(skill_item, "NX-OS"),
        clone_para(skill_item, "ASA"),
        clone_para(skill_item, "CATALYST"),
        clone_para(skill_item, "ACI"),
        clone_para(skill_item, "UCS"),
        clone_para(skill_item, "ARUBA"),
        empty_para(spacer),
        clone_para(skill_cat, "Virtualization"),
        clone_para(skill_item, "VMWARE (ESXi)"),
        empty_para(spacer),
        clone_para(skill_cat, "Automation"),
        clone_para(skill_item, "SCRIPTING"),
        clone_para(skill_item, "CISCO APIs"),
    ]
    replace_cell(skills_cell, skills)

    # Drop the original third table row (it was only REFERENCES and created a blank extra page).
    tbl.remove(rows[2])

    # Original template kept an empty paragraph after the table, which forces a blank
    # last page. Fold sectPr into the last body paragraph instead.
    body = root.find(q("body"))
    sect = body.find(q("sectPr"))
    for child in list(body):
        if child.tag == q("p"):
            body.remove(child)
    last_p = body_cell.findall(q("p"))[-1]
    ppr = last_p.find(q("pPr"))
    if ppr is None:
        ppr = ET.Element(q("pPr"))
        last_p.insert(0, ppr)
    if sect is not None:
        body.remove(sect)
        ppr.append(sect)

    # Update core title
    if "docProps/core.xml" in other:
        core = ET.fromstring(other["docProps/core.xml"])
        # namespaces vary; set dc:title if present
        for el in core.iter():
            if el.tag.endswith("}title"):
                el.text = "Resume — Christopher E. Childs 2026 (v14)"
            if el.tag.endswith("}lastModifiedBy"):
                el.text = "Christopher E. Childs"
        other["docProps/core.xml"] = ET.tostring(core, encoding="utf-8", xml_declaration=True)

    document_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    tmp = OUT + ".tmp"
    with ZipFile(tmp, "w", ZIP_DEFLATED) as z:
        for name, data in other.items():
            z.writestr(name, data)
        z.writestr("word/document.xml", document_xml)
    os.replace(tmp, OUT)
    print("Wrote", OUT)


if __name__ == "__main__":
    build()
