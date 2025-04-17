import os
from bs4 import BeautifulSoup
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

def html_to_slides(html_file, output_file):
    """
    Convert HTML with slide divs to PowerPoint presentation dynamically
    
    Args:
        html_file: Path to the HTML file
        output_file: Path to save the PowerPoint file
    """
    # Read the HTML file
    with open(html_file, 'r', encoding='utf-8') as f:
        html_content = f.read()
    
    # Parse the HTML content
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Find all slide divs
    slides = soup.find_all('div', class_='slide')
    
    # Create a new presentation
    prs = Presentation()
    
    # Process each slide
    for slide_idx, slide_html in enumerate(slides):
        # Add a new slide
        slide_layout = prs.slide_layouts[6]  # Blank layout
        slide = prs.slides.add_slide(slide_layout)
        
        # Find the left and right columns
        left_column = slide_html.find('div', class_='left-column')
        right_column = slide_html.find('div', class_='right-column')
        
        # Get column content
        left_rows = left_column.find_all('div', class_='row') if left_column else []
        right_rows = right_column.find_all('div', class_='row') if right_column else []
        
        # Determine slide dimensions and layout ratios dynamically
        slide_width = prs.slide_width
        slide_height = prs.slide_height
        
        # Calculate dynamic margins and spacing based on slide size
        margin_ratio = 0.05  # 5% of slide dimensions for margins
        h_margin = slide_width * margin_ratio
        v_margin = slide_height * margin_ratio
        spacing = min(h_margin, v_margin) * 0.5  # Spacing between elements
        
        # Determine column widths dynamically
        column_ratio = 0.5  # Default to 50% each column
        left_column_width = (slide_width - (3 * h_margin)) * column_ratio
        right_column_width = (slide_width - (3 * h_margin)) * column_ratio
        
        # Create column containers dynamically
        # Left column container
        left_container = {
            'x': h_margin,
            'y': v_margin * 2,  # Space for title
            'width': left_column_width,
            'height': slide_height - (v_margin * 3)
        }
        
        # Right column container
        right_container = {
            'x': h_margin * 2 + left_column_width,
            'y': v_margin * 2,  # Space for title
            'width': right_column_width,
            'height': slide_height - (v_margin * 3)
        }
        
        # Add containers to slide
        left_shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            left_container['x'],
            left_container['y'],
            left_container['width'],
            left_container['height']
        )
        left_shape.fill.solid()
        left_shape.fill.fore_color.rgb = RGBColor(245, 245, 245)  # Light gray
        left_shape.line.color.rgb = RGBColor(221, 221, 221)  # Light border
        
        right_shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            right_container['x'],
            right_container['y'],
            right_container['width'],
            right_container['height']
        )
        right_shape.fill.solid()
        right_shape.fill.fore_color.rgb = RGBColor(255, 255, 255)  # White
        right_shape.line.color.rgb = RGBColor(221, 221, 221)  # Light border
        
        # Add title dynamically
        title_height = v_margin * 1.5
        title = slide.shapes.add_textbox(
            h_margin,
            v_margin / 2,
            slide_width - (h_margin * 2),
            title_height
        )
        title_text_frame = title.text_frame
        title_text_frame.text = f"Slide {slide_idx + 1}"
        
        # Set title font size dynamically based on slide dimensions
        # Slide dimensions are in EMU (English Metric Units)
        # Convert to points for better calculation (1 inch = 72 points)
        slide_width_pt = slide_width / 914400 * 72  # Convert EMU to points
        slide_height_pt = slide_height / 914400 * 72  # Convert EMU to points
        
        title_font_size = int(min(slide_width_pt, slide_height_pt) * 0.06)  # 6% of min dimension
        title_font_size = max(16, min(title_font_size, 36))  # Keep between 16-36pt
        title_text_frame.paragraphs[0].font.size = Pt(title_font_size)
        title_text_frame.paragraphs[0].font.bold = True
        
        # Calculate row heights for left column dynamically
        if left_rows:
            available_height = left_container['height'] - (spacing * (len(left_rows) + 1))
            left_row_height = available_height / len(left_rows)
        else:
            left_row_height = 0
            
        # Calculate row heights for right column dynamically
        if right_rows:
            available_height = right_container['height'] - (spacing * (len(right_rows) + 1))
            right_row_height = available_height / len(right_rows)
        else:
            right_row_height = 0
        
        # Process left column rows
        for i, row in enumerate(left_rows):
            # Calculate position dynamically
            y_position = left_container['y'] + spacing + (i * (left_row_height + spacing))
            
            # Calculate inner margins for row
            row_margin = spacing / 2
            
            # Create a row shape with background color
            row_shape = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                left_container['x'] + row_margin,
                y_position,
                left_container['width'] - (row_margin * 2),
                left_row_height
            )
            row_shape.fill.solid()
            row_shape.fill.fore_color.rgb = RGBColor(224, 224, 224)  # Light gray for rows
            row_shape.line.color.rgb = RGBColor(200, 200, 200)
            
            # Add text to the row
            left_text_box = slide.shapes.add_textbox(
                left_container['x'] + (row_margin * 2),
                y_position + row_margin,
                left_container['width'] - (row_margin * 4),
                left_row_height - (row_margin * 2)
            )
            left_text_frame = left_text_box.text_frame
            left_text_frame.word_wrap = True
            left_text_frame.text = row.get_text().strip()
            
            # Set font size dynamically based on row height
            # Convert row height from EMU to points
            row_height_pt = left_row_height / 914400 * 72  # Convert EMU to points
            font_size = int(row_height_pt * 0.4)  # 40% of row height in points
            font_size = max(8, min(font_size, 18))  # Keep between 8-18pt
            left_text_frame.paragraphs[0].font.size = Pt(font_size)
            
        # Process right column rows
        for i, row in enumerate(right_rows):
            # Calculate position dynamically
            y_position = right_container['y'] + spacing + (i * (right_row_height + spacing))
            
            # Calculate inner margins for row
            row_margin = spacing / 2
            
            # Create a row shape with background color
            row_shape = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                right_container['x'] + row_margin,
                y_position,
                right_container['width'] - (row_margin * 2),
                right_row_height
            )
            row_shape.fill.solid()
            row_shape.fill.fore_color.rgb = RGBColor(240, 240, 240)  # Lighter gray for rows
            row_shape.line.color.rgb = RGBColor(221, 221, 221)
            
            # Add text to the row
            right_text_box = slide.shapes.add_textbox(
                right_container['x'] + (row_margin * 2),
                y_position + row_margin,
                right_container['width'] - (row_margin * 4),
                right_row_height - (row_margin * 2)
            )
            right_text_frame = right_text_box.text_frame
            right_text_frame.word_wrap = True
            right_text_frame.text = row.get_text().strip()
            
            # Set font size dynamically based on row height
            # Convert row height from EMU to points
            row_height_pt = right_row_height / 914400 * 72  # Convert EMU to points
            font_size = int(row_height_pt * 0.4)  # 40% of row height in points
            font_size = max(8, min(font_size, 18))  # Keep between 8-18pt
            right_text_frame.paragraphs[0].font.size = Pt(font_size)
    
    # Save the presentation
    prs.save(output_file)
    print(f"Presentation saved to {output_file}")
    
def html_string_to_slides(html_content, output_file):
    """Process HTML content directly from a string"""
    # Save the HTML content to a temporary file
    temp_file = "temp_html_file.html"
    with open(temp_file, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    # Process the HTML file
    html_to_slides(temp_file, output_file)
    
    # Remove the temporary file
    os.remove(temp_file)

if __name__ == "__main__":
    # Example usage
    html_file = "sample1.html"  # Path to the HTML file
    output_file = "generated_slides.pptx"  # Output PowerPoint file
    
    html_to_slides(html_file, output_file)