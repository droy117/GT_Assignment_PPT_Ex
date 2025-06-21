import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR

# For chart generation (if desired)
try:
    import matplotlib.pyplot as plt
    import io
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False
    print("Warning: Matplotlib not found. Chart generation will be skipped. Install with 'pip install matplotlib'.")


# --- Configuration ---
EXCEL_FILE_PATH = 'Sample Audit Report v1.xlsx'
POWERPOINT_TEMPLATE_PATH = 'website_assessment_template.pptx'
OUTPUT_POWERPOINT_PATH = 'website_assessment_report_generated.pptx'

# --- Function Definitions ---

def extract_data_from_excel(excel_path):
    """
    Extracts structured data from the Excel file.
    Assumes a single sheet named 'Domain Audit Report' for all data,
    and that the header is on the 4th row (index 3).
    """
    try:
        df_audit_report = pd.read_excel(excel_path, sheet_name='Domain Audit Report', header=3)

        total_issues = len(df_audit_report)

        compliance_score = 'N/A'
        if 'Compliance Score' in df_audit_report.columns:
            score_values = df_audit_report['Compliance Score'].astype(str).str.extract(r'(\d+\.?\d*)\s+out of\s+(\d+\.?\d*)')
            if not score_values.empty:
                score_sum = 0
                total_possible_sum = 0
                for _, row in score_values.iterrows():
                    try:
                        current_score = float(row[0])
                        total_possible = float(row[1])
                        score_sum += current_score
                        total_possible_sum += total_possible
                    except ValueError:
                        continue
                if total_possible_sum > 0:
                    compliance_score = round((score_sum / total_possible_sum) * 100, 2)
                else:
                    compliance_score = 'N/A'
            else:
                compliance_score = 'N/A - Could not parse Compliance Score values'
        else:
            compliance_score = 'N/A - "Compliance Score" column not found'

        overall_summary_text = f"This report summarizes the audit findings for the domain. A total of {total_issues} entries were found. The overall compliance assessment resulted in an average score of {compliance_score}% across all entries. This comprehensive audit covers various aspects including privacy policy adherence, cookie banner deployment, user consent management, and integration with compliance tools. Key areas of strength and identified gaps are presented for a clear overview of the domain's compliance posture."

        key_metrics = {
            'compliance_score': compliance_score,
            'total_issues': total_issues,
            'overall_summary': overall_summary_text
        }

        print("Data extracted successfully from Excel sheet 'Domain Audit Report'.")
        return df_audit_report, df_audit_report, key_metrics
    except FileNotFoundError:
        print(f"Error: Excel file not found at {excel_path}")
        return None, None, None
    except KeyError as e:
        print(f"Error: Missing expected sheet 'Domain Audit Report' or columns in Excel (check header=3 setting): {e}")
        return None, None, None
    except Exception as e:
        print(f"An unexpected error occurred during Excel data extraction: {e}")
        return None, None, None

def create_powerpoint_presentation(template_path):
    """
    Loads the PowerPoint template. If template is truly empty, it might just be a base .pptx file.
    """
    try:
        prs = Presentation(template_path)
        print("PowerPoint template loaded successfully.")
        return prs
    except FileNotFoundError:
        print(f"Error: PowerPoint template not found at {template_path}. Please ensure it exists.")
        return None
    except Exception as e:
        print(f"An unexpected error occurred while loading PowerPoint template: {e}")
        return None

def add_title_slide(prs, key_metrics):
    """
    Adds a title slide and populates it. Assumes 'Title Slide' layout is prs.slide_layouts[0].
    """
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)

    title = slide.shapes.title
    title.text = "Website Assessment Report"

    subtitle = slide.placeholders[1]
    subtitle.text = f"Date: {pd.Timestamp.now().strftime('%Y-%m-%d')}\nOverall Compliance Score: {key_metrics['compliance_score']}%"
    subtitle.text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT # Ensure subtitle fits
    print("Title slide added and populated.")

def add_summary_slide(prs, key_metrics):
    """
    Adds a summary slide and populates it with key insights.
    Assumes 'Title and Content' layout is prs.slide_layouts[1].
    """
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)

    title = slide.shapes.title
    title.text = "Executive Summary"

    # Add key metrics in a dedicated text box
    left, top, width, height = Inches(1), Inches(1.8), Inches(8), Inches(1.5) # Reduced height slightly
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT # Crucial for fitting text

    p = tf.add_paragraph()
    p.text = f"Overall Compliance Score: {key_metrics['compliance_score']}%"
    p.font.size = Pt(20) # Starting font size

    p = tf.add_paragraph()
    p.text = f"Total Entries Assessed: {key_metrics['total_issues']}"
    p.font.size = Pt(20) # Starting font size

    # Add overall summary text
    left, top, width, height = Inches(1), Inches(3.8), Inches(8), Inches(2.5) # Adjusted position and increased height for summary
    txBox_summary = slide.shapes.add_textbox(left, top, width, height)
    tf_summary = txBox_summary.text_frame
    tf_summary.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT # Crucial for fitting text
    tf_summary.word_wrap = True # Ensure word wrapping

    p_summary = tf_summary.add_paragraph()
    p_summary.text = key_metrics['overall_summary']
    p_summary.font.size = Pt(16) # Starting font size

    print("Summary slide added and populated.")

def add_findings_slide(prs, df_findings):
    """
    Adds a new slide for detailed findings with a table.
    """
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)

    title = slide.shapes.title
    title.text = "Detailed Audit Findings"

    rows, cols = df_findings.shape
    num_rows_with_header = rows + 1

    # Adjust table position and size to maximize space for content
    left = Inches(0.2) # Closer to left edge
    top = Inches(1.5)
    width = Inches(9.6) # Wider to accommodate more columns
    height = Inches(5.5)

    table = slide.shapes.add_table(num_rows_with_header, cols, left, top, width, height).table

    # Set column widths - more sophisticated distribution
    # You might need to manually set widths for specific columns if they have more content
    # For now, distributing evenly:
    for i in range(cols):
        table.columns[i].width = Inches(width.inches / cols)

    # Populate header row
    for col_idx, col_name in enumerate(df_findings.columns):
        cell = table.cell(0, col_idx)
        cell.text = str(col_name)
        text_frame = cell.text_frame
        text_frame.paragraphs[0].font.bold = True
        text_frame.paragraphs[0].font.size = Pt(10) # Smaller header font for more space
        text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT # Auto-size header text
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(0x1F, 0x49, 0x7D)
        text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # Populate data rows
    for row_idx, row_data in df_findings.iterrows():
        for col_idx, value in enumerate(row_data):
            cell = table.cell(row_idx + 1, col_idx)
            cell.text = str(value)
            text_frame = cell.text_frame
            text_frame.paragraphs[0].font.size = Pt(8) # Smaller data font for tables
            text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT # Auto-size cell text
            if (row_idx + 1) % 2 == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(0xE0, 0xE0, 0xE0)

    print("Detailed findings slide with table added.")

def add_charts_slide(prs, df_findings):
    """
    Generates and adds charts based on findings data using Matplotlib.
    """
    if not MATPLOTLIB_AVAILABLE:
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        title.text = "Visual Summary of Findings (Charts Not Generated)"
        left, top, width, height = Inches(1), Inches(2), Inches(8), Inches(2)
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        p = tf.add_paragraph()
        p.text = "Matplotlib is not installed. Please install it ('pip install matplotlib') to generate charts."
        print("Placeholder chart slide added due to missing matplotlib.")
        return

    if 'Level of Gap Quantity' in df_findings.columns:
        gap_quantity_counts = df_findings['Level of Gap Quantity'].value_counts()

        fig, ax = plt.subplots(figsize=(8, 4))
        gap_quantity_counts.plot(kind='bar', ax=ax, color=['#4CAF50', '#FFC107', '#F44336', '#9C27B0'])
        ax.set_title('Distribution of Gap Quantity Levels', fontsize=16)
        ax.set_xlabel('Level of Gap Quantity', fontsize=12)
        ax.set_ylabel('Number of Domains', fontsize=12)
        plt.xticks(rotation=45, ha='right')
        plt.tight_layout()

        img_buffer = io.BytesIO()
        plt.savefig(img_buffer, format='png', bbox_inches='tight')
        img_buffer.seek(0)
        plt.close(fig)

        chart_slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(chart_slide_layout)
        title = slide.shapes.title
        title.text = "Visual Summary of Findings"

        left = Inches(1)
        top = Inches(2)
        width = Inches(8)
        height = Inches(4.5)
        slide.shapes.add_picture(img_buffer, left, top, width, height)
        print("Chart slide with image added.")
    else:
        print("Skipping chart generation: 'Level of Gap Quantity' column not found in findings data.")
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        title.text = "Visual Summary of Findings (Data Missing)"
        left, top, width, height = Inches(1), Inches(2), Inches(8), Inches(2)
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        p = tf.add_paragraph()
        p.text = "To generate charts, ensure the Excel data contains relevant columns like 'Level of Gap Quantity' or 'Severity'."
        print("Placeholder chart slide added due to missing data column.")

def save_powerpoint_presentation(prs, output_path):
    """
    Saves the generated PowerPoint presentation.
    """
    try:
        prs.save(output_path)
        print(f"PowerPoint presentation saved successfully to {output_path}")
    except Exception as e:
        print(f"Error saving PowerPoint presentation: {e}")

# --- Main Automation Logic ---
def main():
    print("Starting automation of website assessment report generation...")

    df_audit_report, _, key_metrics = extract_data_from_excel(EXCEL_FILE_PATH)
    if df_audit_report is None:
        print("Exiting: Failed to extract data from Excel.")
        return

    prs = create_powerpoint_presentation(POWERPOINT_TEMPLATE_PATH)
    if prs is None:
        prs = Presentation()
        print("Created a new blank PowerPoint presentation as template was not found/loadable.")

    add_title_slide(prs, key_metrics)
    add_summary_slide(prs, key_metrics)
    add_findings_slide(prs, df_audit_report)
    add_charts_slide(prs, df_audit_report)

    save_powerpoint_presentation(prs, OUTPUT_POWERPOINT_PATH)

    print("Automation complete!")

if __name__ == "__main__":
    main()
