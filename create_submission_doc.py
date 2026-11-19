import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def set_cell_background(cell, fill_hex):
    """Sets the background color of a table cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Sets internal padding for a table cell (in twips)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'''
        <w:tcMar {nsdecls("w")}>
            <w:top w:w="{top}" w:type="dxa"/>
            <w:bottom w:w="{bottom}" w:type="dxa"/>
            <w:left w:w="{left}" w:type="dxa"/>
            <w:right w:w="{right}" w:type="dxa"/>
        </w:tcMar>
    ''')
    tcPr.append(tcMar)

def set_table_borders(table, color="CCCCCC", sz="4", val="single"):
    """Applies borders around and inside a table."""
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
        </w:tblBorders>
    ''')
    tblPr.append(borders)

def add_callout_box(doc, title, text, bg_color="F0F4F8", border_color="1B365D"):
    """Adds a stylish callout box for instructions or notes."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    
    cell = tbl.cell(0, 0)
    cell.width = Inches(6.5)
    set_cell_background(cell, bg_color)
    set_cell_margins(cell, top=120, bottom=120, left=180, right=180)
    
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:top w:val="none"/>
            <w:left w:val="single" w:sz="24" w:space="0" w:color="{border_color}"/>
            <w:bottom w:val="none"/>
            <w:right w:val="none"/>
        </w:tcBorders>
    ''')
    tcPr.append(borders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    run_t = p.add_run(title + "\n")
    run_t.bold = True
    run_t.font.name = "Calibri"
    run_t.font.size = Pt(10.5)
    run_t.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
    
    run_txt = p.add_run(text)
    run_txt.font.name = "Calibri"
    run_txt.font.size = Pt(9.5)
    run_txt.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
    
    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_before = Pt(0)
    p_after.paragraph_format.space_after = Pt(4)

def add_screenshot_slot(doc, heading, description, placeholder_text, height_twips=220):
    """Adds an explicit screenshot heading, caption/description, and an image pasting box."""
    p_head = doc.add_paragraph()
    p_head.paragraph_format.space_before = Pt(8)
    p_head.paragraph_format.space_after = Pt(2)
    r_head = p_head.add_run(heading)
    r_head.bold = True
    r_head.font.name = "Calibri"
    r_head.font.size = Pt(10.5)
    r_head.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
    
    p_desc = doc.add_paragraph()
    p_desc.paragraph_format.space_before = Pt(0)
    p_desc.paragraph_format.space_after = Pt(4)
    r_desc = p_desc.add_run("Description: " + description)
    r_desc.italic = True
    r_desc.font.name = "Calibri"
    r_desc.font.size = Pt(9)
    r_desc.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    set_table_borders(tbl, color="B0C4DE", sz="6", val="dashed")
    
    cell = tbl.cell(0, 0)
    cell.width = Inches(6.5)
    set_cell_background(cell, "FAFCFF")
    set_cell_margins(cell, top=height_twips, bottom=height_twips, left=150, right=150)
    
    p_inner = cell.paragraphs[0]
    p_inner.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_inner.paragraph_format.space_before = Pt(0)
    p_inner.paragraph_format.space_after = Pt(0)
    r_inner = p_inner.add_run(placeholder_text)
    r_inner.bold = True
    r_inner.font.name = "Calibri"
    r_inner.font.size = Pt(9.5)
    r_inner.font.color.rgb = RGBColor(0x7F, 0x92, 0xA8)
    
    p_space = doc.add_paragraph()
    p_space.paragraph_format.space_before = Pt(0)
    p_space.paragraph_format.space_after = Pt(6)

def add_code_block(doc, code_str, max_lines=None):
    """Adds a formatted code snippet box."""
    lines = code_str.strip().split("\n")
    if max_lines and len(lines) > max_lines:
        lines = lines[:max_lines] + ["// ... (refer to OrderManagementSystem.java for full class)"]
    
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    cell = tbl.cell(0, 0)
    cell.width = Inches(6.5)
    set_cell_background(cell, "F8F9FA")
    set_cell_margins(cell, top=100, bottom=100, left=140, right=140)
    
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:top w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>
            <w:left w:val="single" w:sz="18" w:space="0" w:color="4A777A"/>
            <w:bottom w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>
            <w:right w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>
        </w:tcBorders>
    ''')
    tcPr.append(borders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.05
    run = p.add_run("\n".join(lines))
    run.font.name = "Consolas"
    run.font.size = Pt(8.5)
    run.font.color.rgb = RGBColor(0x24, 0x29, 0x2F)
    
    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_before = Pt(0)
    p_after.paragraph_format.space_after = Pt(4)

def build_document():
    doc = docx.Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)
    
    # -------------------------------------------------------------
    # HEADER SECTION
    # -------------------------------------------------------------
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(2)
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = title_p.add_run("OOPJ (CS336B)\n")
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(14)
    r_sub.bold = True
    r_sub.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
    
    r_main = title_p.add_run("Moodle MSE2 Submission Format")
    r_main.font.name = "Calibri"
    r_main.font.size = Pt(16)
    r_main.bold = True
    r_main.font.color.rgb = RGBColor(0x0F, 0x1E, 0x36)
    
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    
    # -------------------------------------------------------------
    # STUDENT DETAILS TABLE
    # -------------------------------------------------------------
    student_table = doc.add_table(rows=6, cols=2)
    student_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    student_table.autofit = False
    set_table_borders(student_table, color="B0C4DE", sz="6")
    
    details_data = [
        ("Name", "Harsh [Enter Your Full Name Here]"),
        ("Branch", "[Enter Branch e.g., CSE / IT / CS / AIML]"),
        ("Roll Number", "[Enter Your University Roll Number Here]"),
        ("Section", "[Enter Your Section e.g., CSE-1 / CSE-2 / CSE-3]"),
        ("Shift", "[Enter Shift 1 / Shift 2]"),
        ("Case Study Name", "E-Commerce Order Management System (Java JDBC + MySQL)")
    ]
    
    for i, (label, val) in enumerate(details_data):
        cell_lbl = student_table.cell(i, 0)
        cell_val = student_table.cell(i, 1)
        
        cell_lbl.width = Inches(2.2)
        cell_val.width = Inches(4.3)
        
        set_cell_background(cell_lbl, "EBF3FB")
        set_cell_margins(cell_lbl, top=70, bottom=70, left=100, right=100)
        set_cell_margins(cell_val, top=70, bottom=70, left=100, right=100)
        
        p_l = cell_lbl.paragraphs[0]
        p_l.paragraph_format.space_before = Pt(0)
        p_l.paragraph_format.space_after = Pt(0)
        r_l = p_l.add_run(label)
        r_l.bold = True
        r_l.font.name = "Calibri"
        r_l.font.size = Pt(10)
        r_l.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
        
        p_v = cell_val.paragraphs[0]
        p_v.paragraph_format.space_before = Pt(0)
        p_v.paragraph_format.space_after = Pt(0)
        r_v = p_v.add_run(val)
        r_v.font.name = "Calibri"
        r_v.font.size = Pt(10)
        if "[" in val:
            r_v.italic = True
            r_v.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
        else:
            r_v.bold = True
            r_v.font.color.rgb = RGBColor(0x11, 0x11, 0x11)
            
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    
    # -------------------------------------------------------------
    # QUESTION 1 SECTION
    # -------------------------------------------------------------
    q1_p = doc.add_paragraph()
    q1_p.paragraph_format.space_before = Pt(10)
    q1_p.paragraph_format.space_after = Pt(4)
    r_q1 = q1_p.add_run("Q1. Assigned Problem Statement & Solution")
    r_q1.bold = True
    r_q1.font.name = "Calibri"
    r_q1.font.size = Pt(13)
    r_q1.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
    
    ps_text = (
        "An e-commerce company wants to develop an Order Management System using Java JDBC and MySQL to manage customer orders. "
        "The database ecommerce_db contains a table named Orders with the following attributes:\n"
        "• order_id — INT, Primary Key\n"
        "• customer_name — VARCHAR(100)\n"
        "• product_name — VARCHAR(100)\n"
        "• quantity — INT\n"
        "• unit_price — DOUBLE\n"
        "• status — VARCHAR(30)\n\n"
        "Write a menu-driven Java JDBC application to perform the following operations:\n"
        "a) Establish a connection with the MySQL database and insert at least three order records using PreparedStatement. Handle duplicate order IDs appropriately.\n"
        "b) Accept an order status from the user and display all orders with that status and a quantity greater than 2. Sort the results by unit price in descending order.\n"
        "c) Accept an order ID and an additional quantity. Increase the existing quantity of the specified order by the given amount. Display an appropriate message if the order does not exist.\n"
        "d) Accept an order ID and a new status. Retrieve the order details and update the status only if the order exists and the new status is valid, such as Pending, Shipped, or Delivered. Display the updated order details.\n"
        "e) Display the total order value using SUM(quantity * unit_price). Accept an order ID and delete the corresponding record only if its status is Cancelled. Display an appropriate message if the order cannot be deleted."
    )
    add_callout_box(doc, "Problem Statement: E-Commerce Order Management System", ps_text, bg_color="F8FAFC", border_color="1B365D")
    
    # Moodle Upload / GitHub Repository Link
    lnk_h = doc.add_paragraph()
    lnk_h.paragraph_format.space_before = Pt(6)
    lnk_h.paragraph_format.space_after = Pt(2)
    r_lh = lnk_h.add_run("Moodle Upload / GitHub Repository Link")
    r_lh.bold = True
    r_lh.font.name = "Calibri"
    r_lh.font.size = Pt(11)
    r_lh.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
    
    repo_table = doc.add_table(rows=2, cols=1)
    repo_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    repo_table.autofit = False
    set_table_borders(repo_table, color="B0C4DE", sz="6")
    
    r_c0 = repo_table.cell(0, 0)
    r_c0.width = Inches(6.5)
    set_cell_background(r_c0, "EBF3FB")
    set_cell_margins(r_c0, top=60, bottom=60, left=100, right=100)
    p0 = r_c0.paragraphs[0]
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after = Pt(0)
    rp0 = p0.add_run("GitHub Repository / Project Moodle Link:")
    rp0.bold = True
    rp0.font.name = "Calibri"
    rp0.font.size = Pt(9.5)
    rp0.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
    
    r_c1 = repo_table.cell(1, 0)
    r_c1.width = Inches(6.5)
    set_cell_margins(r_c1, top=70, bottom=70, left=100, right=100)
    p1 = r_c1.paragraphs[0]
    p1.paragraph_format.space_before = Pt(0)
    p1.paragraph_format.space_after = Pt(0)
    rp1 = p1.add_run("[INSTRUCTION: Paste your GitHub repository URL or Moodle submission link here]\nExample: https://github.com/your-username/ecommerce-order-management-system")
    rp1.font.name = "Consolas"
    rp1.font.size = Pt(9)
    rp1.font.color.rgb = RGBColor(0x00, 0x55, 0xAA)
    
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    
    # Project Code Section
    pc_h = doc.add_paragraph()
    pc_h.paragraph_format.space_before = Pt(8)
    pc_h.paragraph_format.space_after = Pt(2)
    r_pch = pc_h.add_run("Project Code")
    r_pch.bold = True
    r_pch.font.name = "Calibri"
    r_pch.font.size = Pt(11)
    r_pch.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
    
    sql_text = (
        "-- MySQL Database Setup: ecommerce_db\n"
        "CREATE DATABASE IF NOT EXISTS ecommerce_db;\n"
        "USE ecommerce_db;\n\n"
        "CREATE TABLE IF NOT EXISTS Orders (\n"
        "    order_id INT PRIMARY KEY,\n"
        "    customer_name VARCHAR(100) NOT NULL,\n"
        "    product_name VARCHAR(100) NOT NULL,\n"
        "    quantity INT NOT NULL,\n"
        "    unit_price DOUBLE NOT NULL,\n"
        "    status VARCHAR(30) NOT NULL\n"
        ");"
    )
    p_sql_desc = doc.add_paragraph()
    p_sql_desc.paragraph_format.space_before = Pt(0)
    p_sql_desc.paragraph_format.space_after = Pt(1)
    r_sd = p_sql_desc.add_run("1. Database Schema (`ecommerce_db.sql`):")
    r_sd.bold = True
    r_sd.font.size = Pt(9.5)
    add_code_block(doc, sql_text)
    
    java_file_path = r"C:\Users\harsh\eclipse-workspace\shah\demo\src\main\java\com\kiet\examproject\OrderManagementSystem.java"
    java_code = ""
    if os.path.exists(java_file_path):
        with open(java_file_path, "r", encoding="utf-8") as f:
            java_code = f.read()
    
    p_java_desc = doc.add_paragraph()
    p_java_desc.paragraph_format.space_before = Pt(4)
    p_java_desc.paragraph_format.space_after = Pt(1)
    r_jd = p_java_desc.add_run("2. Java Implementation (`com.kiet.examproject.OrderManagementSystem.java`):")
    r_jd.bold = True
    r_jd.font.size = Pt(9.5)
    add_code_block(doc, java_code)
    
    # -------------------------------------------------------------
    # SCREENSHOTS SECTION WITH EXPLICIT HEADINGS
    # -------------------------------------------------------------
    sc_h = doc.add_paragraph()
    sc_h.paragraph_format.space_before = Pt(10)
    sc_h.paragraph_format.space_after = Pt(2)
    r_sch = sc_h.add_run("Screenshots of Model, Controller, and Postman (Database & Execution)")
    r_sch.bold = True
    r_sch.font.name = "Calibri"
    r_sch.font.size = Pt(12)
    r_sch.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
    
    sc_guide = (
        "Evaluation Note: This project is a core Java JDBC application interacting directly with MySQL. "
        "Each screenshot below has an explicit title and caption detailing which exam requirement it validates. "
        "Press Win + Shift + S to take each screenshot, click inside the corresponding dashed box below, and press Ctrl + V."
    )
    add_callout_box(doc, "Screenshots Submission Guidelines", sc_guide, bg_color="FFFDF0", border_color="D4A017")
    
    # Screenshot 1
    add_screenshot_slot(
        doc,
        heading="Screenshot 1: MySQL Database & Orders Table Schema",
        description="Shows the output of 'DESCRIBE Orders;' and initial table structure in MySQL Workbench or Command Line.",
        placeholder_text="[ PASTE SCREENSHOT 1 HERE: MySQL Database & Table Schema ]"
    )
    
    # Screenshot 2
    add_screenshot_slot(
        doc,
        heading="Screenshot 2: Operation (a) – Database Connectivity, PreparedStatement Insertion & Duplicate Handling",
        description="Shows successful insertion of 5 orders, and subsequent duplicate key handling (catching SQLIntegrityConstraintViolationException) without application crash.",
        placeholder_text="[ PASTE SCREENSHOT 2 HERE: Batch Insertion & Duplicate ID Handling ]"
    )
    
    # Screenshot 3
    add_screenshot_slot(
        doc,
        heading="Screenshot 3: Operation (b) – Filter Orders by Status & Quantity > 2 (Sorted by Unit Price DESC)",
        description="Shows user entering status 'Pending' and displaying matching orders having quantity > 2, ordered from highest to lowest unit price.",
        placeholder_text="[ PASTE SCREENSHOT 3 HERE: Filtered & Sorted Query Output ]"
    )
    
    # Screenshot 4
    add_screenshot_slot(
        doc,
        heading="Screenshot 4: Operation (c) – Increase Order Quantity by Order ID (Atomic Update & Validation)",
        description="Shows increasing quantity for Order ID 103 by 2 with updated row display, and error message when an invalid Order ID (999) is entered.",
        placeholder_text="[ PASTE SCREENSHOT 4 HERE: Order Quantity Increment & Not Found Handling ]"
    )
    
    # Screenshot 5
    add_screenshot_slot(
        doc,
        heading="Screenshot 5: Operation (d) – Validate & Update Order Status with Before/After Order Details",
        description="Shows rejection of invalid status ('Refunded'), retrieval of current details before update, and successful update to 'Shipped' with updated details display.",
        placeholder_text="[ PASTE SCREENSHOT 5 HERE: Status Validation & Detailed Update Output ]"
    )
    
    # Screenshot 6
    add_screenshot_slot(
        doc,
        heading="Screenshot 6: Operation (e) – Aggregate Total Order Value (SUM) & Conditional Delete for 'Cancelled' Status",
        description="Shows total value calculated using SUM(quantity * unit_price), deletion rejection for non-cancelled orders (ID 101), and successful deletion of Cancelled order (ID 105).",
        placeholder_text="[ PASTE SCREENSHOT 6 HERE: Aggregate SUM & Cancelled Order Deletion ]"
    )
    
    # -------------------------------------------------------------
    # PROJECT STRUCTURE SCREENSHOT WITH EXPLICIT HEADING
    # -------------------------------------------------------------
    str_h = doc.add_paragraph()
    str_h.paragraph_format.space_before = Pt(10)
    str_h.paragraph_format.space_after = Pt(2)
    r_strh = str_h.add_run("Screenshot of Eclipse/Intellij Project Structure")
    r_strh.bold = True
    r_strh.font.name = "Calibri"
    r_strh.font.size = Pt(12)
    r_strh.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
    
    add_screenshot_slot(
        doc,
        heading="Screenshot 7: Eclipse / IntelliJ Package Explorer & Project Structure",
        description="Shows demo project root, package com.kiet.examproject, OrderManagementSystem.java, and pom.xml containing MySQL Connector/J dependency.",
        placeholder_text="[ PASTE SCREENSHOT 7 HERE: Eclipse/IntelliJ Project Explorer Tree ]",
        height_twips=280
    )
    
    # -------------------------------------------------------------
    # THREE AI PROMPTS SECTION
    # -------------------------------------------------------------
    ai_h = doc.add_paragraph()
    ai_h.paragraph_format.space_before = Pt(10)
    ai_h.paragraph_format.space_after = Pt(2)
    r_aih = ai_h.add_run("Write three AI Prompts used for creating the Project")
    r_aih.bold = True
    r_aih.font.name = "Calibri"
    r_aih.font.size = Pt(12)
    r_aih.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
    
    ai_table = doc.add_table(rows=4, cols=2)
    ai_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    ai_table.autofit = False
    set_table_borders(ai_table, color="B0C4DE", sz="6")
    
    hdr_prompt = ai_table.cell(0, 0)
    hdr_refl = ai_table.cell(0, 1)
    hdr_prompt.width = Inches(3.2)
    hdr_refl.width = Inches(3.3)
    set_cell_background(hdr_prompt, "1B365D")
    set_cell_background(hdr_refl, "1B365D")
    set_cell_margins(hdr_prompt, top=80, bottom=80, left=100, right=100)
    set_cell_margins(hdr_refl, top=80, bottom=80, left=100, right=100)
    
    p_hp = hdr_prompt.paragraphs[0]
    p_hp.paragraph_format.space_before = Pt(0)
    p_hp.paragraph_format.space_after = Pt(0)
    r_hp = p_hp.add_run("AI Prompt Used")
    r_hp.bold = True
    r_hp.font.name = "Calibri"
    r_hp.font.size = Pt(10)
    r_hp.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    
    p_hr = hdr_refl.paragraphs[0]
    p_hr.paragraph_format.space_before = Pt(0)
    p_hr.paragraph_format.space_after = Pt(0)
    r_hr = p_hr.add_run("Brief Reflection & Technical Rationale (For 20 Marks)")
    r_hr.bold = True
    r_hr.font.name = "Calibri"
    r_hr.font.size = Pt(10)
    r_hr.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    
    ai_data = [
        (
            "Prompt 1 (Database Architecture & Safe Batch Insertion):\n"
            "\"Design a MySQL schema for an Orders table (order_id, customer_name, product_name, quantity, unit_price, status). "
            "Write core Java JDBC code using DriverManager and PreparedStatement to establish a connection and insert at least 3 sample records. "
            "Implement proper exception handling for duplicate primary key (order_id) without crashing the application.\"",
            
            "Reflection 1:\n"
            "• Technical Insight: By using PreparedStatement placeholders (?), the query is pre-compiled, preventing SQL injection vulnerabilities.\n"
            "• Exception Handling: Catching SQLIntegrityConstraintViolationException specifically enables the application to detect duplicate primary keys gracefully and inform the user rather than terminating with an unhandled runtime error."
        ),
        (
            "Prompt 2 (Conditional Queries, Sorting & Status Validation):\n"
            "\"Write Java JDBC methods using PreparedStatement to:\n"
            "1. Accept order status from user, filter orders with that status and quantity > 2, and sort by unit_price DESC.\n"
            "2. Accept order_id and additional quantity, increasing existing quantity in the database while checking if order exists.\n"
            "3. Accept order_id and new status, validate status against ('Pending', 'Shipped', 'Delivered'), retrieve existing details before update, and display updated details.\"",
            
            "Reflection 2:\n"
            "• Atomic Update Logic: Updating quantity directly via SQL (`SET quantity = quantity + ?`) guarantees thread-safe, atomic database modification.\n"
            "• Business Validation: Validating input status before executing queries reduces unnecessary database load and ensures database integrity."
        ),
        (
            "Prompt 3 (Aggregate Analytics & Conditional Deletion):\n"
            "\"Implement the final operations in JDBC:\n"
            "1. Compute and display the total order value using aggregate function SUM(quantity * unit_price).\n"
            "2. Accept order_id and delete the record ONLY if its status is 'Cancelled'. If the order does not exist or has a non-cancelled status, reject deletion with an appropriate message.\n"
            "3. Enclose all operations in a clean, interactive, menu-driven console loop with try-with-resources.\"",
            
            "Reflection 3:\n"
            "• Aggregate Computation: Utilizing SQL's native SUM() function delegates mathematical computation to the MySQL engine for superior performance.\n"
            "• Business Constraint Enforcement: Enforcing a two-step verification (fetch status -> check if 'Cancelled' -> delete) prevents accidental deletion of active/delivered customer orders."
        )
    ]
    
    for row_idx, (p_text, r_text) in enumerate(ai_data, start=1):
        c_p = ai_table.cell(row_idx, 0)
        c_r = ai_table.cell(row_idx, 1)
        c_p.width = Inches(3.2)
        c_r.width = Inches(3.3)
        
        bg = "FBFDFF" if row_idx % 2 == 1 else "F4F8FC"
        set_cell_background(c_p, bg)
        set_cell_background(c_r, bg)
        set_cell_margins(c_p, top=70, bottom=70, left=90, right=90)
        set_cell_margins(c_r, top=70, bottom=70, left=90, right=90)
        
        p_cell_p = c_p.paragraphs[0]
        p_cell_p.paragraph_format.space_before = Pt(0)
        p_cell_p.paragraph_format.space_after = Pt(0)
        rp_run = p_cell_p.add_run(p_text)
        rp_run.font.name = "Calibri"
        rp_run.font.size = Pt(8.5)
        
        p_cell_r = c_r.paragraphs[0]
        p_cell_r.paragraph_format.space_before = Pt(0)
        p_cell_r.paragraph_format.space_after = Pt(0)
        rr_run = p_cell_r.add_run(r_text)
        rr_run.font.name = "Calibri"
        rr_run.font.size = Pt(8.5)
        
    doc.add_paragraph().paragraph_format.space_after = Pt(12)
    
    # -------------------------------------------------------------
    # QUESTION 2 SECTION (Blank template with instructions)
    # -------------------------------------------------------------
    doc.add_page_break()
    
    q2_p = doc.add_paragraph()
    q2_p.paragraph_format.space_before = Pt(6)
    q2_p.paragraph_format.space_after = Pt(4)
    r_q2 = q2_p.add_run("Q2. Assigned Problem Statement & Solution (If Applicable)")
    r_q2.bold = True
    r_q2.font.name = "Calibri"
    r_q2.font.size = Pt(13)
    r_q2.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
    
    q2_note = (
        "[INSTRUCTIONS FOR QUESTION 2]:\n"
        "• If you have been assigned a second question for MSE-2, fill in its details below.\n"
        "• If only Question 1 was assigned for your batch/slot, leave this section as 'N/A' or refer to Q1."
    )
    add_callout_box(doc, "Question 2 Placeholder & Instructions", q2_note, bg_color="F8FAFC", border_color="778899")
    
    # Q2 Moodle Upload / GitHub Link
    q2_lnk_h = doc.add_paragraph()
    q2_lnk_h.paragraph_format.space_before = Pt(4)
    q2_lnk_h.paragraph_format.space_after = Pt(2)
    r_q2lh = q2_lnk_h.add_run("Moodle upload / GitHub Repository Link")
    r_q2lh.bold = True
    r_q2lh.font.name = "Calibri"
    r_q2lh.font.size = Pt(11)
    r_q2lh.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
    
    q2_repo_table = doc.add_table(rows=2, cols=1)
    q2_repo_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    q2_repo_table.autofit = False
    set_table_borders(q2_repo_table, color="B0C4DE", sz="6")
    
    q2_rc0 = q2_repo_table.cell(0, 0)
    q2_rc0.width = Inches(6.5)
    set_cell_background(q2_rc0, "EBF3FB")
    set_cell_margins(q2_rc0, top=60, bottom=60, left=100, right=100)
    q2_p0 = q2_rc0.paragraphs[0]
    p_run = q2_p0.add_run("GitHub Repository / Project Moodle Link for Q2:")
    p_run.bold = True
    p_run.font.name = "Calibri"
    p_run.font.size = Pt(9.5)
    p_run.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
    
    q2_rc1 = q2_repo_table.cell(1, 0)
    q2_rc1.width = Inches(6.5)
    set_cell_margins(q2_rc1, top=70, bottom=70, left=100, right=100)
    q2_p1 = q2_rc1.paragraphs[0]
    q2_rp1 = q2_p1.add_run("[INSTRUCTION: Paste repository URL for Question 2 if applicable, or write N/A]")
    q2_rp1.italic = True
    q2_rp1.font.name = "Calibri"
    q2_rp1.font.size = Pt(9.5)
    q2_rp1.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    
    # Q2 Project Code
    q2_pc_h = doc.add_paragraph()
    q2_pc_h.paragraph_format.space_before = Pt(6)
    q2_pc_h.paragraph_format.space_after = Pt(2)
    r_q2pch = q2_pc_h.add_run("Project Code")
    r_q2pch.bold = True
    r_q2pch.font.name = "Calibri"
    r_q2pch.font.size = Pt(11)
    r_q2pch.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
    
    q2_code_box = doc.add_table(rows=1, cols=1)
    q2_code_box.alignment = WD_TABLE_ALIGNMENT.CENTER
    q2_code_box.autofit = False
    set_table_borders(q2_code_box, color="CCCCCC", sz="4", val="dashed")
    c_q2c = q2_code_box.cell(0, 0)
    c_q2c.width = Inches(6.5)
    set_cell_background(c_q2c, "FAFAFA")
    set_cell_margins(c_q2c, top=150, bottom=150, left=150, right=150)
    p_q2c = c_q2c.paragraphs[0]
    p_q2c.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_q2c_in = p_q2c.add_run("[ PASTE SOURCE CODE FOR QUESTION 2 HERE (OR LEAVE BLANK IF N/A) ]")
    r_q2c_in.bold = True
    r_q2c_in.font.name = "Calibri"
    r_q2c_in.font.size = Pt(10)
    r_q2c_in.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    
    # Q2 Screenshots
    q2_sc_h = doc.add_paragraph()
    q2_sc_h.paragraph_format.space_before = Pt(6)
    q2_sc_h.paragraph_format.space_after = Pt(2)
    r_q2sch = q2_sc_h.add_run("Screenshots of Model, Controller, and Postman")
    r_q2sch.bold = True
    r_q2sch.font.name = "Calibri"
    r_q2sch.font.size = Pt(11)
    r_q2sch.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
    
    q2_sc_box = doc.add_table(rows=1, cols=1)
    q2_sc_box.alignment = WD_TABLE_ALIGNMENT.CENTER
    q2_sc_box.autofit = False
    set_table_borders(q2_sc_box, color="CCCCCC", sz="4", val="dashed")
    c_q2sc = q2_sc_box.cell(0, 0)
    c_q2sc.width = Inches(6.5)
    set_cell_background(c_q2sc, "FAFAFA")
    set_cell_margins(c_q2sc, top=150, bottom=150, left=150, right=150)
    p_q2sc = c_q2sc.paragraphs[0]
    p_q2sc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_q2sc_in = p_q2sc.add_run("[ PASTE SCREENSHOTS FOR QUESTION 2 HERE (OR LEAVE BLANK IF N/A) ]")
    r_q2sc_in.bold = True
    r_q2sc_in.font.name = "Calibri"
    r_q2sc_in.font.size = Pt(10)
    r_q2sc_in.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    
    # Q2 Project Structure
    q2_str_h = doc.add_paragraph()
    q2_str_h.paragraph_format.space_before = Pt(6)
    q2_str_h.paragraph_format.space_after = Pt(2)
    r_q2strh = q2_str_h.add_run("Screenshot of Eclipse/Intellij Project Structure.")
    r_q2strh.bold = True
    r_q2strh.font.name = "Calibri"
    r_q2strh.font.size = Pt(11)
    r_q2strh.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
    
    q2_str_box = doc.add_table(rows=1, cols=1)
    q2_str_box.alignment = WD_TABLE_ALIGNMENT.CENTER
    q2_str_box.autofit = False
    set_table_borders(q2_str_box, color="CCCCCC", sz="4", val="dashed")
    c_q2str = q2_str_box.cell(0, 0)
    c_q2str.width = Inches(6.5)
    set_cell_background(c_q2str, "FAFAFA")
    set_cell_margins(c_q2str, top=150, bottom=150, left=150, right=150)
    p_q2str = c_q2str.paragraphs[0]
    p_q2str.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_q2str_in = p_q2str.add_run("[ PASTE PROJECT STRUCTURE SCREENSHOT FOR Q2 (OR LEAVE BLANK IF N/A) ]")
    r_q2str_in.bold = True
    r_q2str_in.font.name = "Calibri"
    r_q2str_in.font.size = Pt(10)
    r_q2str_in.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    
    # Q2 AI Prompts
    q2_ai_h = doc.add_paragraph()
    q2_ai_h.paragraph_format.space_before = Pt(6)
    q2_ai_h.paragraph_format.space_after = Pt(2)
    r_q2aih = q2_ai_h.add_run("Write three AI Prompts used for creating the Project.")
    r_q2aih.bold = True
    r_q2aih.font.name = "Calibri"
    r_q2aih.font.size = Pt(11)
    r_q2aih.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
    
    q2_ai_table = doc.add_table(rows=4, cols=2)
    q2_ai_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    q2_ai_table.autofit = False
    set_table_borders(q2_ai_table, color="B0C4DE", sz="6")
    
    q2_hp = q2_ai_table.cell(0, 0)
    q2_hr = q2_ai_table.cell(0, 1)
    q2_hp.width = Inches(3.2)
    q2_hr.width = Inches(3.3)
    set_cell_background(q2_hp, "778899")
    set_cell_background(q2_hr, "778899")
    set_cell_margins(q2_hp, top=60, bottom=60, left=90, right=90)
    set_cell_margins(q2_hr, top=60, bottom=60, left=90, right=90)
    
    p_q2hp = q2_hp.paragraphs[0]
    p_q2hp.add_run("AI Prompt Used (Q2)").bold = True
    p_q2hp.runs[0].font.name = "Calibri"
    p_q2hp.runs[0].font.size = Pt(9.5)
    p_q2hp.runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    
    p_q2hr = q2_hr.paragraphs[0]
    p_q2hr.add_run("Brief Reflection / Rationale (Q2)").bold = True
    p_q2hr.runs[0].font.name = "Calibri"
    p_q2hr.runs[0].font.size = Pt(9.5)
    p_q2hr.runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    
    for row_idx in range(1, 4):
        c_p = q2_ai_table.cell(row_idx, 0)
        c_r = q2_ai_table.cell(row_idx, 1)
        c_p.width = Inches(3.2)
        c_r.width = Inches(3.3)
        set_cell_background(c_p, "FAFAFA")
        set_cell_background(c_r, "FAFAFA")
        set_cell_margins(c_p, top=60, bottom=60, left=90, right=90)
        set_cell_margins(c_r, top=60, bottom=60, left=90, right=90)
        
        p_cp = c_p.paragraphs[0]
        r1 = p_cp.add_run(f"[Enter AI Prompt {row_idx} for Q2 or leave N/A]")
        r1.italic = True
        r1.font.size = Pt(9)
        r1.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
        
        p_cr = c_r.paragraphs[0]
        r2 = p_cr.add_run(f"[Enter Reflection {row_idx} for Q2 or leave N/A]")
        r2.italic = True
        r2.font.size = Pt(9)
        r2.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
        
    doc.add_paragraph().paragraph_format.space_after = Pt(12)
    
    # -------------------------------------------------------------
    # MSE-2 GUIDELINES & MARKING SCHEME FOOTER
    # -------------------------------------------------------------
    gl_text = (
        "MSE-2 OOPJ (CS336B) Exam Guidelines (19th Nov 2026):\n"
        "1. Bring your laptop with charger. Only one student per seat. Sharing of data, information, stationery, or electronic equipment is strictly prohibited.\n"
        "2. Ensure all software requirements (IDE, JDK, MySQL Server, MySQL Workbench) are pre-configured.\n"
        "3. Save all your AI prompts—they must be submitted as deliverables on Moodle.\n\n"
        "Marks, Submission and Record (Total: 40 Marks):\n"
        "• Complete Code + Three AI Prompts with Brief Reflections (Moodle Submission): 20 Marks\n"
        "• Viva demonstrating business logic & explaining collections/JDBC (In-Person Viva): 20 Marks\n\n"
        "LMS Upload Deliverables:\n"
        "• Zip file of the project (`demo.zip` / `shah.zip`)\n"
        "• Completed Word document in this format (`.docx` / `.pdf`)"
    )
    add_callout_box(doc, "Official MSE-2 Instructions & Evaluation Criteria", gl_text, bg_color="FFF8F0", border_color="E67E22")
    
    out_path = r"C:\Users\harsh\eclipse-workspace\shah\OOPJ_CS336B_MSE2_Submission_Harsh.docx"
    out_path = r"C:\Users\harsh\eclipse-workspace\shah\OOPJ_CS336B_MSE2_Submission_Harsh_Updated.docx"
    doc.save(out_path)
    print(f"Successfully generated document at: {out_path}")

if __name__ == "__main__":
    build_document()
