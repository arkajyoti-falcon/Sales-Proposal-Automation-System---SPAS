import os
from io import BytesIO

import streamlit as st
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="Falcon WCS CONTROLIT – Section Builder", page_icon="🧠")

st.title("Falcon WCS CONTROLIT – DOCX Section Generator")

st.write(
    "Enter client name and download a ready-to-insert DOCX section "
    "for *Falcon’s WCS CONTROLIT* with fixed content and architecture diagrams."
)

# -------------------------------------------------------------------
# CONFIG – image paths (update to match your actual folder structure)
# -------------------------------------------------------------------
IMG_WCS_HEADER = "FIXED_IMAGE\\wcs1.PNG"                 # [img1] – optional title/hero image
IMG_SYSTEM_ARCH = "FIXED_IMAGE\\wcs2.PNG"                # [img2]
IMG_HA_ARCH = "FIXED_IMAGE\\wcs3.PNG"                    # [img3]
IMG_UI_DASHBOARD = "FIXED_IMAGE\\wcs4.PNG"               # Dashboard
IMG_UI_LIVE_BAGS = "FIXED_IMAGE\\wcs5.PNG"               # Live Bags
IMG_UI_BAY_STATUS = "FIXED_IMAGE\\wcs6.PNG"              # Bay Status
IMG_UI_PROCESSED = "FIXED_IMAGE\\wcs7.PNG"               # Processed Packages
IMG_UI_CONFIG = "FIXED_IMAGE\\wcs8.PNG"                  # Configuration Settings
IMG_COMM_ARCH = "FIXED_IMAGE\\wcs9.PNG"                  # Communication Architecture

IMAGE_WIDTH_INCHES = 6.0


def add_centered_image(doc: Document, path: str, width_in: float = IMAGE_WIDTH_INCHES):
    """
    Safely add a centered image if the file exists.
    If the file does not exist, it silently skips it.
    """
    if not path or not os.path.exists(path):
        return

    p = doc.add_paragraph()
    run = p.add_run()
    run.add_picture(path, width=Inches(width_in))
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER


def add_bold_paragraph(doc: Document, text: str, style: str | None = None):
    """
    Add a paragraph where the whole text is bold.
    Optionally apply a paragraph style (e.g., 'List Bullet', 'List Number').
    """
    if style:
        p = doc.add_paragraph(style=style)
    else:
        p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    return p


def build_wcs_controlit_docx(client_name: str) -> BytesIO:
    """Build the 'Falcon WCS CONTROLIT' section as a DOCX and return as BytesIO."""
    if not client_name:
        client_name = "Client"

    doc = Document()

    # Heading 2 – main section title
    doc.add_heading("Falcon’s WCS CONTROLIT", level=2)

    # Optional top image (header / key visual)
    add_centered_image(doc, IMG_WCS_HEADER, width_in=3.0)

    # Intro text
    doc.add_paragraph(
        "Falcon WCS (Warehouse Control System) is an in-house developed IT solution by Falcon Autotech, "
        "serving as the brain behind the company’s sortation solutions. It manages the real-time movement "
        "of goods and data across the system, ensuring efficient operations in high-throughput warehouses. "
        "Falcon WCS integrates seamlessly with Warehouse Management Systems (WMS), Transport Management "
        "Systems (TMS), and other external applications via APIs to enhance operational efficiency."
    )

    # ------------------------------------------------------------------
    # A. System Architecture
    # ------------------------------------------------------------------
    doc.add_heading("A. System Architecture", level=3)

    # High-Level Design
    p = doc.add_paragraph()
    p.add_run("High-Level Design (HLD) Overview").bold = True

    doc.add_paragraph(
        "The Falcon WCS integrates with external systems like the Warehouse Management System (WMS) and "
        "Transport Management System (TMS). Communication occurs via APIs / WSDL / MQ communication "
        "protocols, ensuring smooth data flow for order management, shipment tracking, and other critical "
        "operations."
    )

    # System architecture diagram
    add_centered_image(doc, IMG_SYSTEM_ARCH)

    add_bold_paragraph(doc, "Key Components:")

    # Presentation & Session Layer
    add_bold_paragraph(doc, "Presentation and Session Layer:", style="List Bullet")
    doc.add_paragraph(
        "MySQL Database: Stores operational data, shipment details, and sortation instructions.",
        style="List Bullet 2",
    )
    doc.add_paragraph(
        "Sorter Services: Responsible for managing sorting logic and directing parcels to appropriate destinations.",
        style="List Bullet 2",
    )
    doc.add_paragraph(
        "Dashboard: Provides a user interface for real-time monitoring of warehouse operations and performance metrics.",
        style="List Bullet 2",
    )
    doc.add_paragraph(
        "Integration Services: Handles communication with external systems (e.g., WMS, TMS) "
        "and ensures data consistency across platforms.",
        style="List Bullet 2",
    )

    # Application Layer
    add_bold_paragraph(doc, "Application Layer:", style="List Bullet")
    doc.add_paragraph(
        "Image Services: Processes and manages images captured during the sortation process.",
        style="List Bullet 2",
    )
    doc.add_paragraph(
        "ICR Software: Utilizes Image Character Recognition to read parcel labels and identify shipment information.",
        style="List Bullet 2",
    )
    doc.add_paragraph(
        "PLC Software: Interfaces with Programmable Logic Controllers to manage the physical movement of parcels "
        "and control sortation equipment.",
        style="List Bullet 2",
    )

    # Transport Layer
    add_bold_paragraph(doc, "Transport Layer:", style="List Bullet")
    doc.add_paragraph(
        "Sorter PLCs: Receive commands from the session layer (Sorter Services) and execute sorting operations "
        "based on real-time data.",
        style="List Bullet 2",
    )

    # System communication
    add_bold_paragraph(doc, "System Communication:", style="List Bullet")
    doc.add_paragraph(
        "All layers are connected via a stacked switch, which provides internet and intranet connectivity. "
        "Communication between the sortation system and external systems for results or shipment data occurs "
        "through this switch.",
        style="List Bullet 2",
    )

    # ------------------------------------------------------------------
    # B. High Availability Architecture
    # ------------------------------------------------------------------
    doc.add_heading("B. High Availability Architecture", level=3)

    doc.add_paragraph(
        "The Falcon WCS architecture ensures uninterrupted operations using a High Availability (HA) server setup. "
        "The system is designed to handle both planned and unplanned downtime, providing robust mechanisms for "
        "failover, replication, and data redundancy."
    )

    add_centered_image(doc, IMG_HA_ARCH)

    add_bold_paragraph(doc, "Key Components and Features of the High Availability Architecture:")

    # 1. Stacked Switch
    add_bold_paragraph(doc, "Stacked Switch:", style="List Number")
    doc.add_paragraph(
        "Centralizes data exchange between NAS, nodes, domain controller (DC), and peripherals.",
        style="List Bullet 2",
    )
    doc.add_paragraph(
        "Analyzes packet headers to reduce unnecessary data transmission, enhancing LAN efficiency.",
        style="List Bullet 2",
    )

    # 2. Domain Controller
    add_bold_paragraph(doc, "Domain Controller:", style="List Number")
    doc.add_paragraph(
        "Heartbeat Monitoring: Tracks the status of nodes and initiates VM failover when necessary.",
        style="List Bullet 2",
    )
    doc.add_paragraph(
        "Image Hosting: Stores and manages images received from the ICR (Image Character Recognition).",
        style="List Bullet 2",
    )

    # 3. NAS
    add_bold_paragraph(doc, "NAS (Network Attached Storage):", style="List Number")
    doc.add_paragraph(
        "Centralized data storage providing access to connected devices and virtual machines.",
        style="List Bullet 2",
    )
    doc.add_paragraph(
        "Redundancy: Two NAS boxes with mirrored drives ensure data protection and availability, "
        "offering a failsafe against hardware failure.",
        style="List Bullet 2",
    )

    # 4. Node
    add_bold_paragraph(doc, "Node:", style="List Number")
    doc.add_paragraph(
        "Hyper Terminals: Nodes host and manage virtual machines (VMs) to run the warehouse control systems "
        "and related applications.",
        style="List Bullet 2",
    )
    doc.add_paragraph(
        "Clustering: Nodes are clustered using Microsoft Windows Cluster to enable failover protection, "
        "ensuring continuous operation even in case of hardware failure.",
        style="List Bullet 2",
    )

    # 5. Virtual Machine & InnoDB
    add_bold_paragraph(doc, "Virtual Machine & InnoDB Cluster:", style="List Number")
    doc.add_paragraph(
        "Primary VM: Hosts Falcon WCS services, while a secondary backup on the node ensures failover through "
        "network load balancing (NLB).",
        style="List Bullet 2",
    )
    doc.add_paragraph(
        "InnoDB Cluster: Ensures data replication using a Master–Slave–Slave setup for MySQL databases, "
        "maintaining consistency and availability.",
        style="List Bullet 2",
    )

    # 6. NAS Cluster
    add_bold_paragraph(doc, "NAS Cluster:", style="List Number")
    doc.add_paragraph(
        "Unified File System: NAS nodes share files across the cluster, ensuring no data loss during failover "
        "or disaster recovery.",
        style="List Bullet 2",
    )
    doc.add_paragraph(
        "Backup NAS: Provides redundancy by replicating data between two NAS boxes, further safeguarding "
        "against failures.",
        style="List Bullet 2",
    )

    # Disaster Handling
    add_bold_paragraph(doc, "Disaster Handling:")
    doc.add_paragraph(
        "Recovery Time Objective (RTO) & Data Loss Objective (RPO):",
        style="List Bullet 2",
    )
    doc.add_paragraph(
        "VM Cluster Failure: RTO = 1 hour; RPO = 1 hour.",
        style="List Bullet 3",
    )
    doc.add_paragraph(
        "Node Failure: No impact with a single failure; RTO = 4 hours if both nodes fail.",
        style="List Bullet 3",
    )
    doc.add_paragraph(
        "NAS Failure: Backup NAS available with no downtime, ensuring continued operation.",
        style="List Bullet 3",
    )

    # ------------------------------------------------------------------
    # C. WCS User Interface
    # ------------------------------------------------------------------
    doc.add_heading("C. WCS User Interface", level=3)

   
    doc.add_paragraph(
        "The Falcon WCS features a robust, user-friendly dashboard that provides real-time visibility "
        "into warehouse and sortation operations. The dashboard serves as the primary interface for "
        "monitoring key system metrics, tracking performance, and ensuring smooth operations."
    )

    doc.add_paragraph("Dashboard Overview")
    doc.add_paragraph(
        "The WCS dashboard offers real-time data visualization, helping warehouse operators and IT teams "
        "make data-driven decisions. Users can monitor system health, performance, and detect anomalies "
        "through an intuitive graphical interface."
    )

    add_bold_paragraph(doc, "Key Features of the Dashboard:", style="List Bullet")
    doc.add_paragraph(
        "System Health Monitoring: Displays metrics such as CPU utilization, memory usage, disk performance, "
        "and system load across the infrastructure.",
        style="List Number",
    )
    doc.add_paragraph(
        "Real-Time Sortation Monitoring: Shows the real-time movement of parcels within the sortation system, "
        "including chute assignments and shipment statuses.",
        style="List Number",
    )
    doc.add_paragraph(
        "Error Reporting: Notifies users of system errors, network disruptions, and potential failures in "
        "real time, allowing for quick resolution and minimal downtime.",
        style="List Number",
    )
    doc.add_paragraph(
        "Performance Metrics: Provides detailed reports on sortation throughput, parcel handling times, "
        "and system efficiency to ensure that warehouse targets are met.",
        style="List Number",
    )
    doc.add_paragraph(
        "User Role Management: The dashboard allows different levels of access based on user roles, ensuring "
        "that the right personnel can view or manage the system as needed.",
        style="List Number",
    )

    doc.add_paragraph(
        "In the context of this IT dashboard, the following user interactive screens are provided:",
    )

    # Dashboard screens + images
    doc.add_paragraph(
        "Dashboard (Home Screen): Provides an overview of important metrics, data visualizations, and summary "
        "information related to the IT system or processes.",
        style="List Bullet 2",
    )
    add_centered_image(doc, IMG_UI_DASHBOARD)

    doc.add_paragraph(
        "Live Bags: Displays real-time information and status updates regarding bags or parcels currently in "
        "transit or being processed.",
        style="List Bullet 2",
    )
    add_centered_image(doc, IMG_UI_LIVE_BAGS)

    doc.add_paragraph(
        "Bay Status: Offers insights into the status and availability of different processing bays or areas within the system.",
        style="List Bullet 2",
    )
    add_centered_image(doc, IMG_UI_BAY_STATUS)

    doc.add_paragraph(
        "Processed Packages: Shows details and statistics related to packages or items that have been successfully "
        "processed or handled by the system.",
        style="List Bullet 2",
    )
    add_centered_image(doc, IMG_UI_PROCESSED)

    doc.add_paragraph(
        "Configuration Setting: Enables users to configure and customize various settings and parameters within "
        "the IT system or dashboard.",
        style="List Bullet 2",
    )
    add_centered_image(doc, IMG_UI_CONFIG)

    doc.add_paragraph(
        "Report & Analysis: Allows users to generate and access comprehensive reports, analytics, and insights "
        "based on the data collected by the IT dashboard.",
        style="List Bullet 2",
    )
    doc.add_paragraph(
        "Rejection Bay Mapping: Provides functionality to map and manage rejection bays or areas where packages "
        "are deemed unsuitable for processing.",
        style="List Bullet 2",
    )
    doc.add_paragraph(
        "Alarms: Displays alerts, notifications, or alarms related to system events, errors, or anomalies that "
        "require attention or investigation.",
        style="List Bullet 2",
    )
    doc.add_paragraph(
        "Calibration Settings: Allows users to adjust and calibrate system settings, parameters, or sensors to "
        "ensure accurate and reliable performance.",
        style="List Bullet 2",
    )
    doc.add_paragraph(
        "Operator Management: Offers features and tools to manage and monitor the operators or personnel "
        "responsible for operating the IT system.",
        style="List Bullet 2",
    )
    doc.add_paragraph(
        "User Management: Provides functionality to manage user accounts, permissions, roles, and access levels "
        "within the IT dashboard.",
        style="List Bullet 2",
    )
    doc.add_paragraph(
        "User Guide: The 'User Guide' page offers comprehensive documentation and instructions on how to use "
        "the IT dashboard effectively. It serves as a reference guide for users.",
        style="List Bullet 2",
    )

    # ------------------------------------------------------------------
    # D. Communication Architecture
    # ------------------------------------------------------------------
    doc.add_heading("D. Communication Architecture", level=3)

   
    doc.add_paragraph(
        "Falcon WCS operates within a highly interconnected system, ensuring seamless communication between "
        "the WCS server, on-premises devices (such as sorter PLCs, PTL devices, 1D scanners, and HHT devices), "
        "and client systems. This communication architecture facilitates real-time data exchange and operational "
        "control, optimizing sortation processes and warehouse efficiency."
    )

    add_centered_image(doc, IMG_COMM_ARCH)

    doc.add_paragraph("On-Premises Communication", style="List Bullet")

    # Sorter PLC
    doc.add_paragraph("Sorter PLC Devices:", style="List Bullet 2")
    doc.add_paragraph(
        "Protocol: Falcon WCS communicates with sorter PLCs using either the Siemens S7 protocol or the Omron "
        "communication protocol.",
        style="List Bullet 3",
    )
    doc.add_paragraph(
        "Functionality: The sorter PLC devices receive sortation instructions from the WCS and execute the sorting "
        "process by directing parcels to the appropriate chute based on the system’s real-time data.",
        style="List Bullet 3",
    )

    # PTL devices
    doc.add_paragraph("PTL (Pick-to-Light) Devices:", style="List Bullet 2")
    doc.add_paragraph(
        "Protocol: PTL devices communicate with Falcon WCS using the TCP/IP protocol.",
        style="List Bullet 3",
    )
    doc.add_paragraph(
        "Functionality: The system sends commands to the PTL devices for guiding manual picking operations by "
        "lighting up indicators at the appropriate bins or shelves, improving operational accuracy and speed.",
        style="List Bullet 3",
    )

    # 1D Scanners
    doc.add_paragraph("1D Scanners:", style="List Bullet 2")
    doc.add_paragraph(
        "Protocol: These barcode scanners also use the TCP/IP protocol to communicate with the WCS.",
        style="List Bullet 3",
    )
    doc.add_paragraph(
        "Functionality: The scanners capture barcode data from the parcels, and this information is sent to the WCS "
        "for processing, such as determining sorting destinations.",
        style="List Bullet 3",
    )

    # HHT devices
    doc.add_paragraph("HHT (Handheld Terminal) Devices:", style="List Bullet 2")
    doc.add_paragraph(
        "Protocol: The wireless HHT devices communicate with Falcon WCS over Wi-Fi.",
        style="List Bullet 3",
    )
    doc.add_paragraph(
        "Functionality:",
        style="List Bullet 3",
    )
    doc.add_paragraph(
        "The HHT devices send scan input data (e.g., barcodes) to the server over Wi-Fi.",
        style="List Bullet 3",
    )
    doc.add_paragraph(
        "The WCS processes this data and sends the required output instructions back to the HHT device and "
        "associated PTL devices.",
        style="List Bullet 3",
    )
    doc.add_paragraph(
        "The HHT device executes these instructions, facilitating real-time decision making and execution for operators.",
        style="List Bullet 3",
    )

    # ------------------------------------------------------------------
    # E. Client Communication
    # ------------------------------------------------------------------
    doc.add_heading("E. Client Communication", level=3)

    doc.add_paragraph("Data Transfer Methods:", style="List Bullet")
    doc.add_paragraph(
        "API: Falcon WCS can communicate processed data to client systems through API calls, allowing for "
        "seamless integration with external software.",
        style="List Bullet 2",
    )
    doc.add_paragraph(
        "MQ (Message Queuing): Falcon WCS can also send data via message queues, ensuring reliable delivery of "
        "messages even during network downtime.",
        style="List Bullet 2",
    )
    doc.add_paragraph(
        "WSDL/XML: For structured data exchanges, Falcon WCS supports WSDL and XML formats for client communication.",
        style="List Bullet 2",
    )
    doc.add_paragraph(
        "Other Protocols: Additional methods for data transfer may include customized protocols depending on "
        f"{client_name}'s requirements.",
        style="List Bullet 2",
    )

    doc.add_paragraph("Purpose:", style="List Bullet")
    doc.add_paragraph(
        "The data sent to the client can include sortation results, system performance reports, and operational "
        "analytics, which can be used for further processing or reporting within external systems like Warehouse "
        "Management Systems (WMS) and Transport Management Systems (TMS).",
        style="List Bullet 2",
    )

    # ------------------------------------------------------------------
    # F. HAA Server Specifications (In Client Scope)
    # ------------------------------------------------------------------
    doc.add_heading(f"F. HAA Server Specifications (In {client_name}’s Scope)", level=3)

    doc.add_paragraph(
        "20-core configuration with 128 GB RAM in T440 and 64 GB RAM in T40."
    )

    # Server spec table
    table = doc.add_table(rows=1, cols=3)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "SN"
    hdr_cells[1].text = "Description"
    hdr_cells[2].text = "Qty"

    rows = [
        ("1", "Synology_storage_DS723+ 2.6 GHz AMD Ryzen R1600 Dual-Core, 2 x Gigabit Ethernet Ports, "
              "2 GB ECC DDR4 RAM, 2 x 3.5/2.5\" Bays, 2 x M.2 2280 Slots, 10GbE RJ-45 network upgrade module.", "2"),
        ("2", "Dell Tower Model T440 – PowerEdge T440 Xeon Gold 6148 2.4GHz/20C/27.5MB/150W, 4 x 32GB RDIMM, "
              "2 x 1.2TB 10K RPM SAS, PERC H750, dual hot-plug PS, iDRAC9 Enterprise, dual-port LAN.", "2"),
        ("3", "Dell PowerEdge T40 Intel Xeon E-2224G, 4 x 16 GB RAM, 2 x 1TB SATA 7.2K, RAID 0/1, "
              "480 GB SSD, 1GbE dual LAN card, 1-year onsite NBD.", "1"),
        ("4", "Windows Server 2019", "3"),
        ("5", "Monitor", "1"),
        ("6", "KVM Switch", "1"),
        ("7", "Mouse", "1"),
        ("8", "Keyboard", "1"),
        ("9", "LAN Cable", "10"),
        ("10", "Power Cable", "8"),
        ("11", "VGA Cable", "1"),
        ("12", "42U AC Server Rack", "1"),
        ("13", "Netgear 24-Port Giga Switch", "2"),
        ("14", "2 TB SSD Micron", "4"),
        ("15", "DP to VGA Converter (Cadyce)", "1"),
    ]

    for sn, desc, qty in rows:
        r = table.add_row().cells
        r[0].text = sn
        r[1].text = desc
        r[2].text = qty

    doc.add_paragraph("")
    doc.add_paragraph(f"Below pointers to be taken care by {client_name} for servers:")

    bullet = doc.add_paragraph(
        f"{client_name} should provide servers with the server operating system (OS) pre-installed "
        "(Windows Server 2019 / 2022)."
    )
    bullet.style = "List Bullet"

    bullet = doc.add_paragraph(
        "For optimal performance and reliability, we strongly recommend setting up virtual machines (VMs) over "
        "reputed VM platforms like Hyper-V or VMware. This will enable automatic failover, ensuring high "
        "availability and minimizing downtime in case of any issues with the primary server."
    )
    bullet.style = "List Bullet"

    bullet = doc.add_paragraph(
        "The detailed architecture is described in the server architecture document."
    )
    bullet.style = "List Bullet"

    doc.add_paragraph(
        "By providing a server with the OS and VMs configured, the deployment process will be more seamless and "
        "Falcon’s team can focus on deploying WCS within minimal timelines."
    )

    # Save to buffer
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# -----------------------------
# Streamlit UI
# -----------------------------
client_name_input = st.text_input("Client Name", value="Zepto")

docx_buffer = None
if st.button("Generate Falcon WCS CONTROLIT DOCX"):
    with st.spinner("Building DOCX section..."):
        docx_buffer = build_wcs_controlit_docx(client_name_input.strip())

if docx_buffer:
    st.download_button(
        label="Download WCS CONTROLIT Section (.docx)",
        data=docx_buffer,
        file_name=f"Falcon_WCS_CONTROLIT_{client_name_input.strip() or 'Client'}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
