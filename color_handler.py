from pptx.enum.dml import MSO_THEME_COLOR, MSO_FILL
from pptx.dml.color import RGBColor
import xml.etree.ElementTree as ET

class ColorHandler:
    def __init__(self, presentation):
        self.presentation = presentation
        self.theme_colors = self._extract_theme_colors()
        
    def _extract_theme_colors(self):
        """Extract all theme colors from the presentation"""
        theme_colors = {
            'BACKGROUND_1': '#FFFFFF',  # Default white
            'BACKGROUND_2': '#F2F2F2',  # Default light gray
            'TEXT_1': '#000000',        # Default black
            'TEXT_2': '#666666',        # Default dark gray
            'ACCENT_1': '#4472C4',      # Default blue
            'ACCENT_2': '#ED7D31',      # Default orange
            'ACCENT_3': '#A5A5A5',      # Default gray
            'ACCENT_4': '#FFC000',      # Default yellow
            'ACCENT_5': '#5B9BD5',      # Default light blue
            'ACCENT_6': '#70AD47'       # Default green
        }
        
        try:
            if hasattr(self.presentation, 'slides') and len(self.presentation.slides) > 0:
                slide = self.presentation.slides[0]
                if hasattr(slide, 'part'):
                    # Try different paths to get theme part
                    theme_part = None
                    if hasattr(slide.part, 'slide_layout'):
                        layout_part = slide.part.slide_layout
                        if hasattr(layout_part, 'theme'):
                            theme_part = layout_part.theme
                    
                    if theme_part and hasattr(theme_part, 'element'):
                        root = theme_part.element
                        clr_scheme = root.find('.//a:clrScheme', 
                            {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'})
                        
                        if clr_scheme is not None:
                            # Map theme color elements to their roles
                            color_mappings = {
                                'dk1': 'TEXT_1',
                                'lt1': 'BACKGROUND_1',
                                'dk2': 'TEXT_2',
                                'lt2': 'BACKGROUND_2',
                                'accent1': 'ACCENT_1',
                                'accent2': 'ACCENT_2',
                                'accent3': 'ACCENT_3',
                                'accent4': 'ACCENT_4',
                                'accent5': 'ACCENT_5',
                                'accent6': 'ACCENT_6'
                            }
                            
                            for elem_name, theme_name in color_mappings.items():
                                elem = clr_scheme.find(f'.//a:{elem_name}', 
                                    {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'})
                                if elem is not None:
                                    color = self._extract_color_from_element(elem)
                                    if color:
                                        theme_colors[theme_name] = color
        except Exception as e:
            print(f"Error extracting theme colors: {e}")
            
        print(f"Extracted theme colors: {theme_colors}")  # Debug log
        return theme_colors
    
    def _extract_color_from_element(self, element):
        """Extract color value from XML element"""
        try:
            # Check for sRGB color
            srgb = element.find('.//a:srgbClr', 
                {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'})
            if srgb is not None:
                color = f'#{srgb.get("val")}'
                print(f"Extracted color for element: {color}")  # Debug log
                return color
            
            # Check for system color
            sys_clr = element.find('.//a:sysClr', 
                {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'})
            if sys_clr is not None:
                color = f'#{sys_clr.get("lastClr", sys_clr.get("val"))}'
                print(f"Extracted color for element: {color}")  # Debug log
                return color
            
            # Check for scheme color
            scheme_clr = element.find('.//a:schemeClr', 
                {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'})
            if scheme_clr is not None:
                val = scheme_clr.get('val')
                color = self.theme_colors.get(val.upper(), None)
                print(f"Extracted color for element: {color}")  # Debug log
                return color
                
        except Exception as e:
            print(f"Error extracting color from element: {e}")
        return None
    
    def get_shape_color(self, shape):
        """Get color information for a shape"""
        try:
            print(f"Getting color for shape: {shape}")  # Debug log
            if not hasattr(shape, 'fill'):
                print("Shape has no fill attribute")  # Debug log
                return 'transparent'
                
            fill = shape.fill
            print(f"Fill type: {fill.type if fill else 'None'}")  # Debug log
            
            # Check for no fill - MSO_FILL.BACKGROUND is used when there's no fill
            if fill is None or fill.type == MSO_FILL.BACKGROUND:
                print("Shape has no fill (background)")  # Debug log
                return 'transparent'
                
            # Handle solid fill
            if fill.type == MSO_FILL.SOLID:
                if hasattr(fill, 'fore_color') and fill.fore_color:
                    color = fill.fore_color
                    print(f"Fill fore_color: {color}")  # Debug log
                    
                    # Direct RGB color
                    if hasattr(color, 'rgb') and color.rgb:
                        rgb_color = f'#{color.rgb[0]:02x}{color.rgb[1]:02x}{color.rgb[2]:02x}'
                        print(f"Using RGB color: {rgb_color}")  # Debug log
                        print(f"Extracted color for shape: {rgb_color}")  # Debug log
                        return rgb_color
                    
                    # Theme color
                    if hasattr(color, 'theme_color'):
                        theme_color = str(color.theme_color)
                        print(f"Using theme color: {theme_color}, RGB: {self.theme_colors[theme_color]}")  # Debug log
                        if theme_color in self.theme_colors:
                            return self.theme_colors[theme_color]
                        else:
                            print(f"Theme color {theme_color} not found in theme_colors")  # Debug log
                            return 'transparent'
            
            # Handle gradient fill
            if fill.type == MSO_FILL.GRADIENT:
                print("Gradient fill detected")  # Debug log
                # Implement gradient handling logic here
                return 'gradient'  # Placeholder
            
            print("No valid fill color found")  # Debug log
            return 'transparent'
            
        except Exception as e:
            print(f"Error getting shape color: {e}")
            print(f"Shape properties: {dir(shape)}")  # Debug log
            return 'transparent'
    
    def get_text_color(self, run):
        """Get text color from a run"""
        try:
            print(f"Getting text color for run: {run.text}")  # Debug log
            if hasattr(run, 'font') and hasattr(run.font, 'color'):
                color = run.font.color
                print(f"Font color object: {color}")  # Debug log
                
                # Direct RGB color
                if hasattr(color, 'rgb') and color.rgb:
                    rgb_color = f'#{color.rgb[0]:02x}{color.rgb[1]:02x}{color.rgb[2]:02x}'
                    print(f"Using RGB color: {rgb_color}")  # Debug log
                    return rgb_color
                
                # Theme color
                if hasattr(color, 'theme_color'):
                    theme_color = str(color.theme_color)
                    print(f"Using theme color: {theme_color}")  # Debug log
                    if theme_color in self.theme_colors:
                        return self.theme_colors[theme_color]
            
            print("Using default black color")  # Debug log
            return '#000000'  # Default black
            
        except Exception as e:
            print(f"Error getting text color: {e}")
            print(f"Run properties: {dir(run)}")  # Debug log
            return '#000000'  # Default black 