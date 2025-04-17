import os
import re
import base64
import requests
from bs4 import BeautifulSoup
from tinycss2 import parse_stylesheet
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR
from pptx.dml.color import RGBColor
from io import BytesIO
from PIL import Image
import math

class HTMLToPPTX:
    def __init__(self):
        self.prs = Presentation()
        self.css_rules = {}
        self.image_counter = 0
        
        # Default slide dimensions (16:9 aspect ratio)
        self.slide_width = 13.33
        self.slide_height = 7.5
        
        # Dynamic layout settings
        self.current_slide = None
        self.current_slide_content_height = 0
        self.content_scale_factor = 1.0
        self.total_height_on_current_slide = 0
        self.y_position = 0
        
        # Margins
        self.margin_left = 0.5
        self.margin_right = 0.5
        self.margin_top = 0.5
        self.margin_bottom = 0.5
        
        # Content area dimensions
        self.content_width = self.slide_width - self.margin_left - self.margin_right
        self.content_height = self.slide_height - self.margin_top - self.margin_bottom
        
    def extract_internal_css(self, soup):
        """Extract CSS rules from internal style tags"""
        style_tags = soup.find_all('style')
        
        for style_tag in style_tags:
            css_text = style_tag.string
            if css_text:
                # Parse CSS rules
                self.parse_css(css_text)
    
    def parse_css(self, css_text):
        """Parse CSS text into a dictionary of rules"""
        # Simple CSS parser
        rule_pattern = r'([^{]+){([^}]*)}'
        rules = re.findall(rule_pattern, css_text)
        
        for selector, properties in rules:
            selector = selector.strip()
            properties_dict = {}
            
            # Extract individual properties
            for prop in properties.split(';'):
                if ':' in prop:
                    key, value = prop.split(':', 1)
                    properties_dict[key.strip()] = value.strip()
            
            self.css_rules[selector] = properties_dict
    
    def get_style_for_element(self, element):
        """Get computed style for an element based on CSS rules"""
        computed_style = {}
        
        # Check for matching CSS rules
        for selector, properties in self.css_rules.items():
            # Very simplified selector matching - only handles class and element selectors
            if '.' in selector:
                # Class selector
                class_name = selector.split('.')[1]
                if element.get('class') and class_name in element.get('class'):
                    computed_style.update(properties)
            elif selector.lower() == element.name:
                # Element selector
                computed_style.update(properties)
        
        # Apply inline style (higher priority)
        if element.get('style'):
            inline_styles = {}
            for prop in element.get('style').split(';'):
                if ':' in prop:
                    key, value = prop.split(':', 1)
                    inline_styles[key.strip()] = value.strip()
            computed_style.update(inline_styles)
        
        return computed_style
    
    def convert_color(self, color_str):
        """Convert CSS color string to RGB tuple"""
        if not color_str:
            return None
            
        # Handle hex colors
        if color_str.startswith('#'):
            color_str = color_str.lstrip('#')
            if len(color_str) == 3:
                color_str = ''.join([c*2 for c in color_str])
            return RGBColor(int(color_str[0:2], 16), int(color_str[2:4], 16), int(color_str[4:6], 16))
        
        # Handle rgb colors
        rgb_match = re.match(r'rgb\((\d+),\s*(\d+),\s*(\d+)\)', color_str)
        if rgb_match:
            return RGBColor(int(rgb_match.group(1)), int(rgb_match.group(2)), int(rgb_match.group(3)))
        
        # Handle named colors (simplified)
        color_map = {
            'white': RGBColor(255, 255, 255),
            'black': RGBColor(0, 0, 0),
            'red': RGBColor(255, 0, 0),
            'green': RGBColor(0, 128, 0),
            'blue': RGBColor(0, 0, 255),
            'yellow': RGBColor(255, 255, 0),
            'gray': RGBColor(128, 128, 128),
            'lightgray': RGBColor(211, 211, 211),
            'darkgray': RGBColor(169, 169, 169),
            # Add more as needed
        }
        
        return color_map.get(color_str.lower())
    
    def apply_text_style(self, text_frame, style):
        """Apply CSS style to a PowerPoint text frame"""
        if not style:
            return
            
        # Apply text alignment
        if 'text-align' in style:
            alignment = style['text-align'].lower()
            if alignment == 'center':
                text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            elif alignment == 'right':
                text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT
            elif alignment == 'justify':
                text_frame.paragraphs[0].alignment = PP_ALIGN.JUSTIFY
                
        # Apply text color
        if 'color' in style:
            color = self.convert_color(style['color'])
            if color:
                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.color.rgb = color
        
        # Apply font family
        if 'font-family' in style:
            # Extract first font family from the list
            font_family = style['font-family'].split(',')[0].strip("'").strip('"')
            for paragraph in text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.name = font_family
        
        # Apply font size
        if 'font-size' in style:
            size_str = style['font-size']
            size_match = re.match(r'(\d+(\.\d+)?)([a-z]+|%)?', size_str)
            if size_match:
                size_val = float(size_match.group(1))
                size_unit = size_match.group(3) if size_match.group(3) else 'px'
                
                # Convert to points (approximate)
                if size_unit == 'px':
                    size_pt = size_val * 0.75  # 1px ≈ 0.75pt
                elif size_unit == 'em' or size_unit == 'rem':
                    size_pt = size_val * 12  # Assuming 1em = 12pt
                elif size_unit == '%':
                    size_pt = size_val * 0.12  # Assuming 100% = 12pt
                elif size_unit == 'pt':
                    size_pt = size_val
                else:
                    size_pt = 12  # Default
                
                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(size_pt * self.content_scale_factor)
        
        # Apply font weight
        if 'font-weight' in style:
            weight = style['font-weight']
            is_bold = weight in ['bold', 'bolder'] or (weight.isdigit() and int(weight) >= 700)
            for paragraph in text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.bold = is_bold
        
        # Apply font style
        if 'font-style' in style:
            style_val = style['font-style']
            is_italic = style_val == 'italic'
            for paragraph in text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.italic = is_italic
        
        # Apply text decoration
        if 'text-decoration' in style:
            decoration = style['text-decoration']
            is_underlined = 'underline' in decoration
            for paragraph in text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.underline = is_underlined
    
    def estimate_text_height(self, text, font_size=12, width=8):
        """Estimate the height needed for a text block in inches"""
        if not text:
            return 0.1
            
        # Count lines
        lines = text.split('\n')
        
        # Calculate average chars per line based on width and font size
        chars_per_line = int(width * 120 / (font_size/12))
        
        # Count total lines needed after wrapping
        total_lines = 0
        for line in lines:
            if not line.strip():
                total_lines += 1
                continue
                
            line_chars = len(line)
            line_count = math.ceil(line_chars / chars_per_line) if chars_per_line > 0 else 1
            total_lines += line_count
        
        # Calculate height based on font size and line count
        line_height = font_size * 1.2 / 72  # Convert points to inches with 1.2 line spacing
        padding = 0.1  # Padding in inches
        
        return (total_lines * line_height) + padding
    
    def get_available_height(self):
        """Get available height on current slide"""
        return self.content_height - self.y_position
    
    def create_new_slide(self):
        """Create a new slide and reset positioning"""
        slide_layout = self.prs.slide_layouts[6]  # Blank slide
        self.current_slide = self.prs.slides.add_slide(slide_layout)
        self.y_position = self.margin_top
        self.content_scale_factor = 1.0  # Reset scale factor for new slide
        return self.current_slide
    
    def add_element_with_auto_positioning(self, element_callable, *args, **kwargs):
        """Add an element to the current slide with automatic positioning and pagination
        
        Args:
            element_callable: Function to call to add the element
            *args, **kwargs: Arguments to pass to the element_callable
            
        Returns:
            Height used by the element
        """
        # Initialize current_slide if not already set
        if self.current_slide is None:
            self.create_new_slide()
        
        # Get estimated height for the element
        estimated_height = self.estimate_element_height(args[0])
        
        # Check if we need a new slide
        if self.y_position > self.margin_top and self.y_position + estimated_height > self.content_height:
            self.create_new_slide()
        
        # Calculate positioning
        left = self.margin_left
        top = self.y_position
        width = self.content_width
        
        # Add the element
        actual_height = element_callable(args[0], self.current_slide, left, top, width, *args[1:], **kwargs)
        
        # Update position for next element
        self.y_position += actual_height + 0.2  # Add spacing
        
        return actual_height
    
    def estimate_element_height(self, element):
        """Estimate height needed for an element based on its type and content"""
        if element.name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
            level = int(element.name[1])
            font_size = 36 if level == 1 else 28 if level == 2 else 24 if level == 3 else 18
            text = element.get_text(strip=True)
            return self.estimate_text_height(text, font_size, self.content_width) + 0.1
            
        elif element.name == 'p':
            text = element.get_text(strip=True)
            return self.estimate_text_height(text, 12, self.content_width) + 0.1
            
        elif element.name == 'img':
            # Get image dimensions from attributes
            height = element.get('height')
            if height:
                return float(height) / 96 + 0.2  # Convert pixels to inches + padding
            return 2.0  # Default height
            
        elif element.name == 'table':
            rows = len(element.find_all('tr'))
            return rows * 0.3 + 0.3  # Approximate height
            
        elif element.name in ['ul', 'ol']:
            items = len(element.find_all('li'))
            item_text = " ".join([item.get_text(strip=True) for item in element.find_all('li')])
            return self.estimate_text_height(item_text, 12, self.content_width) + (items * 0.05) + 0.2
            
        elif element.name == 'div':
            if element.get('class') and 'code-block' in element.get('class'):
                text = element.get_text()
                lines = len(text.split('\n'))
                return lines * 0.15 + 0.3  # Approximate height for code
                
            # For other divs, estimate based on content
            height = 0.1  # Base height
            for child in element.children:
                if child.name:  # Skip text nodes
                    height += self.estimate_element_height(child) + 0.1
            return height
            
        # Default height estimate for other elements
        return 0.5
    
    def process_image(self, img_element, slide, left, top, width, height=None):
        """Process an image element and add it to the slide"""
        src = img_element.get('src')
        if not src:
            return 0.1  # Minimal height if no image
        
        # Get image dimensions from attributes or CSS
        if width is None:
            width = img_element.get('width')
            if width:
                width = float(width) / 96  # Convert pixels to inches (approximate)
            else:
                width = self.content_width / 2  # Default width
        
        if height is None:
            height = img_element.get('height')
            if height:
                height = float(height) / 96  # Convert pixels to inches (approximate)
            else:
                height = 2  # Default height
        
        # Ensure image fits within content width
        if width > self.content_width:
            # Scale height proportionally
            height = height * (self.content_width / width)
            width = self.content_width
        
        # Apply content scaling
        width *= self.content_scale_factor
        height *= self.content_scale_factor
        
        try:
            # Handle different image sources
            if src.startswith('http'):
                # Remote image
                response = requests.get(src)
                if response.status_code == 200:
                    img_content = BytesIO(response.content)
                    slide.shapes.add_picture(img_content, Inches(left), Inches(top), 
                                             width=Inches(width), height=Inches(height))
            elif src.startswith('data:image'):
                # Data URL image
                _, encoded = src.split(',', 1)
                img_data = base64.b64decode(encoded)
                img_content = BytesIO(img_data)
                slide.shapes.add_picture(img_content, Inches(left), Inches(top), 
                                          width=Inches(width), height=Inches(height))
            else:
                # Local image - handle case where image doesn't exist gracefully
                if os.path.exists(src):
                    slide.shapes.add_picture(src, Inches(left), Inches(top), 
                                              width=Inches(width), height=Inches(height))
                else:
                    # Add placeholder box
                    shape = slide.shapes.add_shape(
                        1, Inches(left), Inches(top), Inches(width), Inches(height)
                    )
                    shape.fill.solid()
                    shape.fill.fore_color.rgb = RGBColor(200, 200, 200)
                    
                    # Add text to placeholder
                    text_frame = shape.text_frame
                    text_frame.text = f"Image: {os.path.basename(src)}"
        except Exception as e:
            print(f"Error processing image {src}: {e}")
            # Add placeholder box
            shape = slide.shapes.add_shape(
                1, Inches(left), Inches(top), Inches(width), Inches(height)
            )
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(200, 200, 200)
            
            # Add text to placeholder
            text_frame = shape.text_frame
            text_frame.text = f"Image: {os.path.basename(src)}"
        
        return height  # Return actual height used
    
    def process_table(self, table_element, slide, left, top, width, height=None):
        """Process a table element and add it to the slide"""
        # Count rows and columns
        rows = table_element.find_all('tr')
        if not rows:
            return 0.1
        
        # Count columns from the first row
        for row in rows:
            cells = row.find_all(['th', 'td'])
            if cells:
                num_cols = len(cells)
                break
        else:
            num_cols = 1  # Default if no cells found
        
        num_rows = len(rows)
        
        # Calculate dynamic row height
        row_height = 0.25  # Default row height
        
        # Calculate average content per cell to adjust row height
        total_content = 0
        for row in rows:
            cells = row.find_all(['th', 'td'])
            for cell in cells:
                total_content += len(cell.get_text(strip=True))
        
        avg_content_per_cell = total_content / (num_rows * num_cols) if num_rows * num_cols > 0 else 0
        
        # Adjust row height based on average content
        if avg_content_per_cell > 30:
            row_height = 0.35
        if avg_content_per_cell > 60:
            row_height = 0.45
            
        # Apply content scaling
        row_height *= self.content_scale_factor
        
        # Create table with dimensions that fit the slide
        table_width = width
        table_height = num_rows * row_height
        
        # Check if table fits on slide
        max_table_height = self.content_height - top + self.margin_top
        if table_height > max_table_height:
            # Scale down if needed
            scale_factor = max_table_height / table_height
            table_height *= scale_factor
            row_height *= scale_factor
        
        table_shape = slide.shapes.add_table(
            num_rows, num_cols, 
            Inches(left), Inches(top), 
            Inches(table_width), Inches(table_height)
        )
        table = table_shape.table
        
        # Process table rows
        for i, row in enumerate(rows):
            cells = row.find_all(['th', 'td'])
            
            # Process cells in this row
            for j, cell in enumerate(cells):
                if j >= num_cols:  # Skip excess cells
                    break
                
                # Get cell style
                cell_style = self.get_style_for_element(cell)
                is_header = cell.name == 'th'
                
                # Get cell text
                cell_text = cell.get_text(strip=True)
                
                # Set cell text
                table_cell = table.cell(i, j)
                table_cell.text = cell_text
                
                # Apply styling
                if is_header:
                    for paragraph in table_cell.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.bold = True
                
                # Apply cell style
                if 'background-color' in cell_style:
                    bg_color = self.convert_color(cell_style['background-color'])
                    if bg_color:
                        table_cell.fill.solid()
                        table_cell.fill.fore_color.rgb = bg_color
                
                # Apply text style
                self.apply_text_style(table_cell.text_frame, cell_style)
        
        return table_height  # Return actual height used
    
    def process_code_block(self, code_element, slide, left, top, width, height=None):
        """Process a code block element and add it to the slide"""
        # Get code text
        code_text = code_element.get_text()
        
        # Create text box
        code_style = self.get_style_for_element(code_element)
        
        # Determine height based on content
        lines = len(code_text.split('\n'))
        code_height = lines * 0.15 + 0.2  # Approximate height
        
        # Apply content scaling
        code_height *= self.content_scale_factor
        
        # Create textbox for code
        textbox = slide.shapes.add_textbox(
            Inches(left), Inches(top), Inches(width), Inches(code_height)
        )
        
        # Add code text
        text_frame = textbox.text_frame
        text_frame.text = code_text
        text_frame.word_wrap = True
        
        # Apply monospace font
        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.name = 'Courier New'
                run.font.size = Pt(10 * self.content_scale_factor)
        
        # Apply styling
        self.apply_text_style(text_frame, code_style)
        
        # Apply background if specified
        if code_style and 'background-color' in code_style:
            textbox.fill.solid()
            textbox.fill.fore_color.rgb = self.convert_color(code_style['background-color'])
        
        return code_height  # Return actual height used
    
    def process_list(self, list_element, slide, left, top, width, height=None):
        """Process an ordered or unordered list element and add it to the slide"""
        is_ordered = list_element.name == 'ol'
        list_items = list_element.find_all('li')
        
        if not list_items:
            return 0.1
        
        # Estimate height based on content
        items_text = " ".join([item.get_text(strip=True) for item in list_items])
        font_size = 12  # Default
        
        # Adjust font size based on list size
        if len(list_items) > 6:
            font_size = 11
        if len(list_items) > 10:
            font_size = 10
            
        # Apply content scaling
        font_size *= self.content_scale_factor
        
        # Estimate required height
        list_height = self.estimate_text_height(items_text, font_size, width) + (len(list_items) * 0.05)
        
        # Create text box for list
        textbox = slide.shapes.add_textbox(
            Inches(left), Inches(top), Inches(width), Inches(list_height)
        )
        
        text_frame = textbox.text_frame
        text_frame.word_wrap = True
        
        # Add list items
        for i, item in enumerate(list_items):
            # Get item text
            item_text = item.get_text(strip=True)
            
            # Add prefix based on list type
            if is_ordered:
                prefix = f"{i+1}. "
            else:
                prefix = "• "
            
            # Add paragraph for list item
            if i == 0:
                paragraph = text_frame.paragraphs[0]
            else:
                paragraph = text_frame.add_paragraph()
                
            paragraph.text = prefix + item_text
            
            # Apply styling
            item_style = self.get_style_for_element(item)
            for run in paragraph.runs:
                if item_style:
                    if 'color' in item_style:
                        color = self.convert_color(item_style['color'])
                        if color:
                            run.font.color.rgb = color
                    
                    if 'font-weight' in item_style and item_style['font-weight'] == 'bold':
                        run.font.bold = True
                
                # Apply font size
                run.font.size = Pt(font_size)
        
        # Apply list style
        list_style = self.get_style_for_element(list_element)
        self.apply_text_style(text_frame, list_style)
        
        return list_height  # Return actual height used
    
    def process_heading(self, heading_element, slide, left, top, width, height=None):
        """Process a heading element and add it to the slide"""
        # Get heading level
        level = int(heading_element.name[1])
        
        # Get heading text
        heading_text = heading_element.get_text(strip=True)
        
        # Scale font size based on heading level and text length
        base_font_size = 36 if level == 1 else 28 if level == 2 else 24 if level == 3 else 18
        
        # Reduce font size for long headings
        if len(heading_text) > 40 and level <= 2:
            base_font_size -= 4
        if len(heading_text) > 60 and level <= 2:
            base_font_size -= 4
            
        # Apply content scaling
        font_size = base_font_size * self.content_scale_factor
        
        # Estimate height
        heading_height = self.estimate_text_height(heading_text, font_size, width)
        
        # Create text box
        textbox = slide.shapes.add_textbox(
            Inches(left), Inches(top), Inches(width), Inches(heading_height)
        )
        
        # Add heading text
        text_frame = textbox.text_frame
        text_frame.text = heading_text
        text_frame.word_wrap = True
        
        # Apply heading styling
        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(font_size)
                run.font.bold = True
        
        # Apply custom styling
        style = self.get_style_for_element(heading_element)
        self.apply_text_style(text_frame, style)
        
        # Center H1 and H2
        if level <= 2:
            text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            
        return heading_height  # Return actual height used
    
    def process_paragraph(self, p_element, slide, left, top, width, height=None):
        """Process a paragraph element and add it to the slide"""
        # Get paragraph text
        p_text = p_element.get_text(strip=True)
        
        # Scale font size based on content length
        base_font_size = 12
        if len(p_text) > 200:
            base_font_size = 11
        if len(p_text) > 300:
            base_font_size = 10
            
        # Apply content scaling
        font_size = base_font_size * self.content_scale_factor
        
        # Estimate text height
        text_height = self.estimate_text_height(p_text, font_size, width)
        
        # Create text box
        textbox = slide.shapes.add_textbox(
            Inches(left), Inches(top), Inches(width), Inches(text_height)
        )
        
        # Add paragraph text
        text_frame = textbox.text_frame
        text_frame.text = p_text
        text_frame.word_wrap = True
        
        # Set font size
        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(font_size)
        
        # Apply styling
        style = self.get_style_for_element(p_element)
        self.apply_text_style(text_frame, style)
        
        # Check for special classes
        if p_element.get('class'):
            if 'caption' in p_element.get('class'):
                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.italic = True
                        run.font.size = Pt(10 * self.content_scale_factor)
                
                # Center captions
                text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        return text_height  # Return actual height used
    
    def process_two_column(self, div_element, slide, left, top, width, height=None):
        """Process a div with two columns"""
        columns = div_element.find_all(class_='column')
        if len(columns) < 2:
            # Process as regular div if not enough columns
            return self.process_div_content(div_element, slide, left, top, width)
        
        col_width = (width / 2) - 0.1  # Half width with small gap
        max_height = 0
        
        # Process first column
        if columns[0]:
            col1_height = self.process_div_content(columns[0], slide, left, top, col_width)
            max_height = max(max_height, col1_height)
        
        # Process second column
        if columns[1]:
            col2_height = self.process_div_content(columns[1], slide, left + col_width + 0.2, top, col_width)
            max_height = max(max_height, col2_height)
        
        return max_height  # Return the taller height
    
    def process_div(self, div_element, slide, left, top, width, height=None):
        """Process a div element and add it to the slide"""
        style = self.get_style_for_element(div_element)
        
        # Check for two-column layout
        if div_element.get('class') and 'two-column' in div_element.get('class'):
            return self.process_two_column(div_element, slide, left, top, width)
            
        # Special handling for code blocks
        if div_element.get('class') and 'code-block' in div_element.get('class'):
            return self.process_code_block(div_element, slide, left, top, width)
        
        # Process regular div content
        return self.process_div_content(div_element, slide, left, top, width)
    
    def process_div_content(self, div_element, slide, left, top, width):
        """Process the content of a div element"""
        current_top = top
        total_height = 0
        
        # Apply div background if specified
        div_style = self.get_style_for_element(div_element)
        if div_style and 'background-color' in div_style:
            # We'll estimate the div height based on content height
            estimated_height = 0
            for child in div_element.children:
                if child.name:  # Skip text nodes
                    estimated_height += self.estimate_element_height(child) + 0.1
            
            # Create shape for background
            if estimated_height > 0:
                shape = slide.shapes.add_shape(
                    1,  # Rectangle
                    Inches(left), 
                    Inches(top), 
                    Inches(width), 
                    Inches(estimated_height)
                )
                
                # Apply background color
                shape.fill.solid()
                shape.fill.fore_color.rgb = self.convert_color(div_style['background-color'])
                
                # Make border transparent if not specified
                if 'border' not in div_style:
                    shape.line.fill.background()
        
        # Process all child elements
        for child in div_element.children:
            if child.name is None:
                # Skip text nodes
                continue
                
            element_height = 0
            
            # Process element based on type
            if child.name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
                element_height = self.process_heading(child, slide, left, current_top, width)
            elif child.name == 'p':
                element_height = self.process_paragraph(child, slide, left, current_top, width)
            elif child.name == 'img':
                element_height = self.process_image(child, slide, left, current_top, width)
            elif child.name == 'table':
                element_height = self.process_table(child, slide, left, current_top, width)
            elif child.name in ['ol', 'ul']:
                element_height = self.process_list(child, slide, left, current_top, width)
            elif child.name == 'div':
                if child.get('class'):
                    if 'code-block' in child.get('class'):
                        element_height = self.process_code_block(child, slide, left, current_top, width)
                    elif 'two-column' in child.get('class'):
                        element_height = self.process_two_column(child, slide, left, current_top, width)
                    elif 'image-container' in child.get('class'):
                        element_height = self.process_image_container(child, slide, left, current_top, width)
                    else:
                        element_height = self.process_div(child, slide, left, current_top, width)
                else:
                    element_height = self.process_div(child, slide, left, current_top, width)
            
            # Move down for next element and add spacing
            current_top += element_height + 0.1
            total_height += element_height + 0.1
        
        return total_height
    
    def process_image_container(self, container_element, slide, left, top, width):
        """Process an image container with potential caption"""
        img = container_element.find('img')
        caption = container_element.find(class_='caption')
        container_height = 0
        
        if img:
            img_width = float(img.get('width', 300)) / 96 if img.get('width') else width * 0.8
            img_height = float(img.get('height', 200)) / 96 if img.get('height') else 2
            
            # Center image
            img_left = left + (width - img_width) / 2 if img_width < width else left
            container_height += self.process_image(img, slide, img_left, top, img_width, img_height)
        
        if caption:
            caption_style = self.get_style_for_element(caption)
            caption_text = caption.get_text(strip=True)
            
            caption_box = slide.shapes.add_textbox(
                Inches(left), 
                Inches(top + container_height + 0.1), 
                Inches(width), 
                Inches(0.3)
            )
            caption_box.text_frame.text = caption_text
            
            # Apply styling
            for paragraph in caption_box.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.italic = True
                    run.font.size = Pt(10 * self.content_scale_factor)
            
            # Center caption
            caption_box.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            
            # Apply custom styling if available
            self.apply_text_style(caption_box.text_frame, caption_style)
            
            container_height += 0.4
        
        return container_height
    
    def auto_fit_content(self, slide_element):
        """Analyze content and determine if scaling is needed"""
        # Count all content elements
        content_elements = []
        for element in slide_element.find_all(['h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'p', 'img', 'table', 'ul', 'ol', 'div']):
            # Skip nested elements in two-column layouts to avoid double counting
            if element.parent.get('class') and 'column' in element.parent.get('class'):
                continue
                
            # Skip divs that are just containers for other counted elements
            if element.name == 'div' and element.get('class'):
                if any(cls in element.get('class') for cls in ['row', 'slide']):
                    continue
            
            content_elements.append(element)
        
        # Make initial estimate of content height
        estimated_height = 0
        for element in content_elements:
            estimated_height += self.estimate_element_height(element) + 0.2  # Add spacing
        
        # If content exceeds available height, calculate scaling factor
        available_height = self.content_height
        
        # Reserve space for title if present
        title = slide_element.find(['h1', 'h2'])
        if title:
            title_height = self.estimate_element_height(title) + 0.3
            available_height -= title_height
            # Remove title from estimation since we accounted for it
            if title in content_elements:
                content_elements.remove(title)
                estimated_height -= (self.estimate_element_height(title) + 0.2)
        
        # Calculate scale factor if needed
        if estimated_height > available_height and estimated_height > 0:
            return max(0.6, available_height / estimated_height)  # Don't scale below 60%
        
        return 1.0  # No scaling needed
    
    def process_slide(self, slide_element):
        """Process a single slide element and create a PowerPoint slide for it"""
        # Create a new slide
        slide = self.create_new_slide()
        
        # Reset y position for new slide
        self.y_position = self.margin_top
        
        # Get slide style
        slide_style = self.get_style_for_element(slide_element)
        
        # Apply background color if specified
        if slide_style and 'background-color' in slide_style:
            bg_color = self.convert_color(slide_style['background-color'])
            if bg_color:
                slide.background.fill.solid()
                slide.background.fill.fore_color.rgb = bg_color
        
        # Auto-calculate content scaling if needed
        self.content_scale_factor = self.auto_fit_content(slide_element)
        
        if self.content_scale_factor < 1.0:
            print(f"Automatically scaling content to {self.content_scale_factor:.0%} to fit slide")
        
        # Look for a heading (h1 or h2) to use as the slide title
        slide_title = slide_element.find(['h1', 'h2'])
        if slide_title:
            title_text = slide_title.get_text(strip=True)
            if title_text:
                # Add title - don't apply content scaling to title
                title_shape = slide.shapes.add_textbox(
                    Inches(self.margin_left), 
                    Inches(self.margin_top), 
                    Inches(self.content_width), 
                    Inches(0.8)
                )
                
                title_shape.text_frame.text = title_text
                title_shape.text_frame.word_wrap = True
                
                # Apply title styling
                title_style = self.get_style_for_element(slide_title)
                
                # Keep original content_scale_factor
                original_scale = self.content_scale_factor
                self.content_scale_factor = 1.0
                self.apply_text_style(title_shape.text_frame, title_style)
                self.content_scale_factor = original_scale
                
                # Center title
                title_shape.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
                
                # Apply default styling for titles
                font_size = 36
                if len(title_text) > 30:
                    font_size = 32
                if len(title_text) > 50:
                    font_size = 28
                
                for paragraph in title_shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(font_size)
                        run.font.bold = True
                
                # Update y position to start content after title
                self.y_position += 1.0
                
                # Find all rows in the slide
                rows = slide_element.find_all(class_='row')
                
                # Check if we have rows or just process all content
                if rows:
                    # Process each row
                    for row in rows:
                        # Process the row content
                        row_height = self.process_div(row, slide, self.margin_left, self.y_position, self.content_width)
                        
                        # Move down for the next row
                        self.y_position += row_height + 0.2
                else:
                    # No rows, process all content directly
                    # Skip the title since we already processed it
                    for child in slide_element.children:
                        if child.name and child != slide_title:
                            if child.name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
                                height = self.process_heading(child, slide, self.margin_left, self.y_position, self.content_width)
                                self.y_position += height + 0.2
                            elif child.name == 'div':
                                height = self.process_div(child, slide, self.margin_left, self.y_position, self.content_width)
                                self.y_position += height + 0.2
                            elif child.name == 'p':
                                height = self.process_paragraph(child, slide, self.margin_left, self.y_position, self.content_width)
                                self.y_position += height + 0.2
                            # Add other element types as needed
            else:
                # No proper title, process all content as body
                self.process_div(slide_element, slide, self.margin_left, self.y_position, self.content_width)
        else:
            # No title, process all content as body
            self.process_div(slide_element, slide, self.margin_left, self.y_position, self.content_width)
    
    def convert(self, html_content, output_file):
        """Convert HTML with internal CSS to PowerPoint"""
        # Parse the HTML
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # Extract CSS from internal style tags
        self.extract_internal_css(soup)
        
        # Calculate content area dimensions
        self.content_width = self.slide_width - self.margin_left - self.margin_right
        self.content_height = self.slide_height - self.margin_top - self.margin_bottom
        
        # Find all slide elements
        slide_elements = soup.find_all(class_='slide')
        
        # Check if slides were found
        if not slide_elements:
            print("Warning: No slide elements (div with class 'slide') found in HTML")
            # Try using the whole document as a single slide
            self.process_slide(soup.body)
        else:
            # Process each slide
            for slide_element in slide_elements:
                # Reset for each slide
                self.y_position = self.margin_top
                self.content_scale_factor = 1.0
                
                self.process_slide(slide_element)
        
        # Save the presentation
        self.prs.save(output_file)
        print(f"Presentation saved to {output_file}")

def html_to_pptx(html_content, output_file, slide_width=13.33, slide_height=7.5):
    """Convenience function to convert HTML to PPTX
    
    Args:
        html_content (str): HTML content to convert
        output_file (str): Path to save the PowerPoint file
        slide_width (float, optional): Width of slides in inches. Defaults to 13.33 (16:9).
        slide_height (float, optional): Height of slides in inches. Defaults to 7.5 (16:9).
    """
    converter = HTMLToPPTX()
    converter.slide_width = slide_width
    converter.slide_height = slide_height
    converter.convert(html_content, output_file)
    
def file_to_pptx(html_file, output_file, slide_width=13.33, slide_height=7.5):
    """Convert an HTML file to PowerPoint
    
    Args:
        html_file (str): Path to HTML file
        output_file (str): Path to save the PowerPoint file
        slide_width (float, optional): Width of slides in inches. Defaults to 13.33 (16:9).
        slide_height (float, optional): Height of slides in inches. Defaults to 7.5 (16:9).
    """
    with open(html_file, "r", encoding="utf-8") as f:
        html_content = f.read()
    
    html_to_pptx(html_content, output_file, slide_width, slide_height)
    
if __name__ == "__main__":
    # Example usage
    import sys
    import os
    
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
        output_file = sys.argv[2] if len(sys.argv) > 2 else "output_presentation.pptx"
    else:
        # Default to sample.html if no arguments provided
        input_file = "sample.html"
        output_file = "output_presentation.pptx"
        print(f"No input file specified. Using default: {input_file}")
    
    # Check if the file exists
    if not os.path.exists(input_file):
        print(f"Error: Input file '{input_file}' not found.")
        print("Usage: python html_to_pptx_converter.py input.html [output.pptx]")
        sys.exit(1)
        
    print(f"Converting {input_file} to {output_file}...")
    file_to_pptx(input_file, output_file)
    print(f"Conversion complete! PowerPoint file saved to {output_file}")